/**
 * Minimal Open Packaging Conventions layer: parts, relationships, lazy XML parsing.
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
    if (!zip.file('[Content_Types].xml')) {
      throw new Error('This file is not a PowerPoint (.pptx) file: it has no content types part.');
    }
    return new OpcPackage(zip);
  }

  has(partPath: string): boolean {
    return this.zip.file(partPath) !== null;
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

  async binary(partPath: string): Promise<Uint8Array> {
    const file = this.zip.file(partPath);
    if (!file) throw new Error(`The file is missing a part it refers to: ${partPath}`);
    return file.async('uint8array');
  }

  setBinary(partPath: string, data: Uint8Array): void {
    this.xmlParts.delete(partPath);
    this.zip.file(partPath, data);
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
