/**
 * Minimal Open Packaging Conventions layer: parts, relationships, content types, lazy XML.
 * Paths are package-absolute without a leading slash, e.g. `ppt/slides/slide1.xml`.
 */
import JSZip from 'jszip';
import { NS, children, parseXml, serializeXml } from './xml.ts';
import type { Document } from './xml.ts';

export interface Relationship {
  id: string;
  type: string;
  /** Resolved package path, or the raw target when external. */
  target: string;
  external: boolean;
}

const CONTENT_TYPES = '[Content_Types].xml';

export class OpcPackage {
  private readonly zip: JSZip;
  private readonly xmlParts = new Map<string, Document>();

  private constructor(zip: JSZip) {
    this.zip = zip;
  }

  static async load(data: Uint8Array): Promise<OpcPackage> {
    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch {
      throw new Error('This file is not a PowerPoint (.pptx) file, or it is damaged.');
    }
    if (!zip.file(CONTENT_TYPES)) {
      throw new Error('This file is not a PowerPoint (.pptx) file: it has no content types part.');
    }
    return new OpcPackage(zip);
  }

  has(partPath: string): boolean {
    return this.zip.file(partPath) !== null;
  }

  parts(): string[] {
    return Object.keys(this.zip.files).filter((p) => !this.zip.files[p]!.dir);
  }

  async xml(partPath: string): Promise<Document> {
    const cached = this.xmlParts.get(partPath);
    if (cached) return cached;
    const file = this.zip.file(partPath);
    if (!file) throw new Error(`The file is missing a part it refers to: ${partPath}`);
    const doc = parseXml(await file.async('string'), partPath);
    this.xmlParts.set(partPath, doc);
    return doc;
  }

  setXml(partPath: string, doc: Document): void {
    this.xmlParts.set(partPath, doc);
    this.zip.file(partPath, ''); // placeholder so has()/parts() see it; serialized on save
  }

  async binary(partPath: string): Promise<Uint8Array> {
    const cached = this.xmlParts.get(partPath);
    if (cached) return new TextEncoder().encode(serializeXml(cached));
    const file = this.zip.file(partPath);
    if (!file) throw new Error(`The file is missing a part it refers to: ${partPath}`);
    return file.async('uint8array');
  }

  setBinary(partPath: string, data: Uint8Array): void {
    this.xmlParts.delete(partPath);
    this.zip.file(partPath, data);
  }

  removePart(partPath: string): void {
    this.xmlParts.delete(partPath);
    this.zip.remove(partPath);
    const rels = relsPathFor(partPath);
    this.xmlParts.delete(rels);
    this.zip.remove(rels);
  }

  async rels(partPath: string): Promise<Relationship[]> {
    const relsPath = relsPathFor(partPath);
    if (!this.has(relsPath)) return [];
    const doc = await this.xml(relsPath);
    return children(doc.documentElement!, NS.rel, 'Relationship').map((el) => {
      const external = el.getAttribute('TargetMode') === 'External';
      const raw = el.getAttribute('Target') ?? '';
      return {
        id: el.getAttribute('Id') ?? '',
        type: el.getAttribute('Type') ?? '',
        target: external ? raw : resolveTarget(partPath, raw),
        external,
      };
    });
  }

  async relById(partPath: string, id: string): Promise<Relationship | undefined> {
    return (await this.rels(partPath)).find((r) => r.id === id);
  }

  /** Replaces a part's relationships. Internal targets are written package-absolute. */
  writeRels(partPath: string, rels: Relationship[]): void {
    const body = rels
      .map((r) => {
        const target = r.external ? r.target : `/${r.target}`;
        const mode = r.external ? ' TargetMode="External"' : '';
        return `<Relationship Id="${esc(r.id)}" Type="${esc(r.type)}" Target="${esc(target)}"${mode}/>`;
      })
      .join('');
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rel}">${body}</Relationships>`;
    this.setXml(relsPathFor(partPath), parseXml(xml, relsPathFor(partPath)));
  }

  /** Adds a relationship and returns its new id. */
  async addRel(partPath: string, type: string, target: string): Promise<string> {
    const rels = await this.rels(partPath);
    const used = new Set(rels.map((r) => r.id));
    let n = rels.length + 1;
    while (used.has(`rId${n}`)) n++;
    const id = `rId${n}`;
    this.writeRels(partPath, [...rels, { id, type, target, external: false }]);
    return id;
  }

  async removeRel(partPath: string, id: string): Promise<void> {
    const rels = await this.rels(partPath);
    this.writeRels(partPath, rels.filter((r) => r.id !== id));
  }

  /** A free part name like `ppt/slides/slide7.xml`, numbered after the existing ones. */
  freeName(dir: string, stem: string, ext: string): string {
    const re = new RegExp(`^${escapeRe(dir)}/${escapeRe(stem)}(\\d+)${escapeRe(ext)}$`, 'i');
    let max = 0;
    for (const p of this.parts()) {
      const m = re.exec(p);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `${dir}/${stem}${max + 1}${ext}`;
  }

  /** Copies a part (and its content-type override) to a new free name in the same folder. */
  async clonePart(partPath: string): Promise<string> {
    const slash = partPath.lastIndexOf('/');
    const dir = partPath.slice(0, slash);
    const file = partPath.slice(slash + 1);
    const m = /^(.*?)(\d*)(\.[^.]+)$/.exec(file);
    const stem = m?.[1] ?? file;
    const ext = m?.[3] ?? '';
    const dest = this.freeName(dir, stem, ext);
    const cached = this.xmlParts.get(partPath);
    if (cached) this.setXml(dest, parseXml(serializeXml(cached), dest));
    else this.zip.file(dest, await this.binary(partPath));
    // Same folder, so relative targets in a copied .rels stay valid.
    const rels = await this.rels(partPath);
    if (rels.length) this.writeRels(dest, rels);
    const type = await this.overrideType(partPath);
    if (type) await this.setOverride(dest, type);
    return dest;
  }

  async overrideType(partPath: string): Promise<string | undefined> {
    const ct = await this.xml(CONTENT_TYPES);
    const el = children(ct.documentElement!, NS.ct, 'Override').find((o) => o.getAttribute('PartName')?.toLowerCase() === `/${partPath}`.toLowerCase());
    return el?.getAttribute('ContentType') ?? undefined;
  }

  async setOverride(partPath: string, contentType: string): Promise<void> {
    const ct = await this.xml(CONTENT_TYPES);
    const root = ct.documentElement!;
    const el = ct.createElementNS(NS.ct, 'Override');
    el.setAttribute('PartName', `/${partPath}`);
    el.setAttribute('ContentType', contentType);
    root.appendChild(el);
  }

  /**
   * Removes every part that can no longer be reached by following relationships from the
   * package root, together with its relationships and content-type override.
   */
  async collectGarbage(): Promise<string[]> {
    const reachable = new Set<string>();
    const queue = [''];
    while (queue.length) {
      const part = queue.shift()!;
      for (const r of await this.rels(part)) {
        if (r.external || reachable.has(r.target) || !this.has(r.target)) continue;
        reachable.add(r.target);
        queue.push(r.target);
      }
    }
    const removed: string[] = [];
    for (const p of this.parts()) {
      if (p === CONTENT_TYPES || p.endsWith('.rels') || reachable.has(p)) continue;
      this.removePart(p);
      removed.push(p);
    }
    // Relationship parts whose source part is gone.
    for (const p of this.parts()) {
      if (!p.endsWith('.rels') || p === '_rels/.rels') continue;
      const source = p.replace(/_rels\/([^/]+)\.rels$/, '$1');
      if (!this.has(source)) this.removePart(p);
    }
    // Overrides for parts that don't exist (removed now, or already dangling in the template).
    const ct = await this.xml(CONTENT_TYPES);
    const present = new Set(this.parts().map((p) => `/${p}`.toLowerCase()));
    for (const o of children(ct.documentElement!, NS.ct, 'Override')) {
      if (!present.has((o.getAttribute('PartName') ?? '').toLowerCase())) ct.documentElement!.removeChild(o);
    }
    return removed;
  }

  async save(): Promise<Uint8Array> {
    for (const [partPath, doc] of this.xmlParts) {
      this.zip.file(partPath, serializeXml(doc));
    }
    return this.zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash >= 0 ? partPath.slice(0, slash + 1) : '';
  const name = partPath.slice(slash + 1);
  return `${dir}_rels/${name}.rels`;
}

/** Resolves a relationship target relative to its source part (RFC 3986 style, POSIX separators). */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith('/')) return normalize(target.slice(1));
  const slash = sourcePart.lastIndexOf('/');
  const baseDir = slash >= 0 ? sourcePart.slice(0, slash) : '';
  return normalize(baseDir ? `${baseDir}/${target}` : target);
}

function normalize(p: string): string {
  const out: string[] = [];
  for (const seg of decodeURI(p).split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}
