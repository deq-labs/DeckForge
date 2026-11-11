/** Slide-level package operations: list, duplicate, delete. See design-log 0005. */
import type { OpcPackage } from './opc.ts';
import { NS, REL, children, descendants, removeNode } from './xml.ts';
import type { Element } from './xml.ts';

export const REL_NOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const NS_P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

export interface SlideRef {
  path: string;
  relId: string;
  /** Numeric `p:sldId/@id`. */
  id: number;
}

export async function presentationPath(pkg: OpcPackage): Promise<string> {
  const rootRels = await pkg.rels('');
  return rootRels.find((r) => r.type === REL.officeDocument)?.target ?? 'ppt/presentation.xml';
}

function sldIdList(presDoc: Element): Element[] {
  const lst = children(presDoc, NS.p, 'sldIdLst')[0];
  return lst ? children(lst, NS.p, 'sldId') : [];
}

export async function listSlides(pkg: OpcPackage): Promise<SlideRef[]> {
  const main = await presentationPath(pkg);
  const pres = (await pkg.xml(main)).documentElement!;
  const rels = await pkg.rels(main);
  const out: SlideRef[] = [];
  for (const el of sldIdList(pres)) {
    const relId = el.getAttributeNS(NS.r, 'id') ?? '';
    const rel = rels.find((r) => r.id === relId && r.type === REL.slide);
    if (rel && pkg.has(rel.target)) out.push({ path: rel.target, relId, id: Number(el.getAttribute('id')) });
  }
  return out;
}

export async function notesPathOf(pkg: OpcPackage, slidePath: string): Promise<string | undefined> {
  const rel = (await pkg.rels(slidePath)).find((r) => r.type === REL_NOTES && !r.external);
  return rel && pkg.has(rel.target) ? rel.target : undefined;
}

/** Section lists (PowerPoint 2010+) list slides by numeric id. */
function sectionSlideIds(presDoc: Element): Element[] {
  return descendants(presDoc, NS_P14, 'sldId');
}

/**
 * Duplicates a slide and inserts the copy right after `afterPath` (default: the source).
 * Charts and the notes slide are deep-copied; layouts and media are shared.
 */
export async function duplicateSlide(pkg: OpcPackage, sourcePath: string, afterPath = sourcePath): Promise<string> {
  const copy = await pkg.clonePart(sourcePath);
  const rels = await pkg.rels(sourcePath);
  const newRels = [];
  for (const r of rels) {
    if (r.external || !pkg.has(r.target)) {
      newRels.push(r);
    } else if (r.type === REL.chart) {
      newRels.push({ ...r, target: await deepClone(pkg, r.target) });
    } else if (r.type === REL_NOTES) {
      const notes = await pkg.clonePart(r.target);
      const notesRels = (await pkg.rels(r.target)).map((nr) => (nr.type === REL.slide ? { ...nr, target: copy } : nr));
      pkg.writeRels(notes, notesRels);
      newRels.push({ ...r, target: notes });
    } else {
      newRels.push(r);
    }
  }
  pkg.writeRels(copy, newRels);

  const main = await presentationPath(pkg);
  const presDoc = await pkg.xml(main);
  const pres = presDoc.documentElement!;
  const relId = await pkg.addRel(main, REL.slide, copy);
  const ids = sldIdList(pres);
  const newId = Math.max(255, ...ids.map((e) => Number(e.getAttribute('id')))) + 1;
  const presRels = await pkg.rels(main);
  const afterRel = presRels.find((r) => r.type === REL.slide && r.target === afterPath);
  const afterEl = ids.find((e) => e.getAttributeNS(NS.r, 'id') === afterRel?.id);
  const el = presDoc.createElementNS(NS.p, afterEl?.tagName ?? 'p:sldId');
  el.setAttribute('id', String(newId));
  el.setAttributeNS(NS.r, `${pres.lookupPrefix(NS.r) ?? 'r'}:id`, relId);
  const lst = children(pres, NS.p, 'sldIdLst')[0]!;
  lst.insertBefore(el, afterEl?.nextSibling ?? null);

  // Keep the copy in the same PowerPoint section as the slide it follows.
  const afterId = afterEl?.getAttribute('id');
  for (const s of sectionSlideIds(pres)) {
    if (s.getAttribute('id') === afterId) {
      const ns = presDoc.createElementNS(NS_P14, s.tagName);
      ns.setAttribute('id', String(newId));
      s.parentNode!.insertBefore(ns, s.nextSibling);
    }
  }
  return copy;
}

/** Copies a part and every part its relationships point to (one level), e.g. a chart and its workbook. */
async function deepClone(pkg: OpcPackage, partPath: string): Promise<string> {
  const copy = await pkg.clonePart(partPath);
  const rels = [];
  for (const r of await pkg.rels(partPath)) {
    rels.push(r.external || !pkg.has(r.target) ? r : { ...r, target: await pkg.clonePart(r.target) });
  }
  if (rels.length) pkg.writeRels(copy, rels);
  return copy;
}

/** Removes a slide from the presentation. Its parts are dropped by `collectGarbage`. */
export async function deleteSlide(pkg: OpcPackage, slidePath: string): Promise<void> {
  const main = await presentationPath(pkg);
  const presDoc = await pkg.xml(main);
  const pres = presDoc.documentElement!;
  const rel = (await pkg.rels(main)).find((r) => r.type === REL.slide && r.target === slidePath);
  if (!rel) return;
  const el = sldIdList(pres).find((e) => e.getAttributeNS(NS.r, 'id') === rel.id);
  const id = el?.getAttribute('id');
  if (el) removeNode(el);
  for (const s of sectionSlideIds(pres)) if (s.getAttribute('id') === id) removeNode(s);
  // Custom shows reference slides by relationship id.
  for (const s of descendants(pres, NS.p, 'sld')) if (s.getAttributeNS(NS.r, 'id') === rel.id) removeNode(s);
  await pkg.removeRel(main, rel.id);
}
