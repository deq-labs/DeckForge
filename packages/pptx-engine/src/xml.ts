import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element, Node } from '@xmldom/xmldom';

export type { Document, Element, Node };

export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
} as const;

export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  package: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
} as const;

const ELEMENT_NODE = 1;

export function parseXml(text: string, partName: string): Document {
  const parser = new DOMParser({
    onError: (level, message) => {
      if (level !== 'warning') throw new Error(`Could not read ${partName}: ${message}`);
    },
  });
  return parser.parseFromString(text, 'text/xml');
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

export function isElement(node: Node | null | undefined): node is Element {
  return node != null && node.nodeType === ELEMENT_NODE;
}

export function is(node: Node | null | undefined, ns: string, local: string): node is Element {
  return isElement(node) && node.namespaceURI === ns && node.localName === local;
}

/** Direct child elements, optionally filtered by namespace + local name. */
export function children(el: Element, ns?: string, local?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (!isElement(n)) continue;
    if (ns !== undefined && (n.namespaceURI !== ns || (local !== undefined && n.localName !== local))) continue;
    out.push(n);
  }
  return out;
}

export function child(el: Element, ns: string, local: string): Element | undefined {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (is(n, ns, local)) return n;
  }
  return undefined;
}

/** Follows a path of local names in one namespace, e.g. path(sp, NS.p, 'nvSpPr', 'cNvPr'). */
export function path(el: Element, ns: string, ...locals: string[]): Element | undefined {
  let cur: Element | undefined = el;
  for (const local of locals) {
    if (!cur) return undefined;
    cur = child(cur, ns, local);
  }
  return cur;
}

export function descendants(el: Element | Document, ns: string, local: string): Element[] {
  return Array.from(el.getElementsByTagNameNS(ns, local));
}

/** Creates an element in `ns`, reusing whatever prefix `context` already uses for it. */
export function create(context: Element, ns: string, local: string): Element {
  const doc = context.ownerDocument!;
  const prefix = context.lookupPrefix(ns);
  return doc.createElementNS(ns, prefix ? `${prefix}:${local}` : local);
}

export function removeNode(node: Node): void {
  node.parentNode?.removeChild(node);
}

export function removeChildren(el: Element): void {
  while (el.firstChild) el.removeChild(el.firstChild);
}

export function setText(el: Element, text: string): void {
  removeChildren(el);
  el.appendChild(el.ownerDocument!.createTextNode(text));
}

export function textOf(el: Element): string {
  return el.textContent ?? '';
}
