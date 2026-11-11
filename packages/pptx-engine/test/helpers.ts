import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Document, Element } from '@xmldom/xmldom';
import { NS } from '../src/xml.ts';

const here = dirname(fileURLToPath(import.meta.url));
export const FIXTURES = resolve(here, '../../../fixtures/templates');

export function fixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(FIXTURES, name)));
}

export async function unzip(bytes: Uint8Array): Promise<JSZip> {
  return JSZip.loadAsync(bytes);
}

export async function partXml(zip: JSZip, path: string): Promise<Document> {
  const f = zip.file(path);
  if (!f) throw new Error(`no part ${path}`);
  return new DOMParser().parseFromString(await f.async('string'), 'text/xml');
}

function paragraphText(p: Element): string {
  let s = '';
  for (let n = p.firstChild; n; n = n.nextSibling) {
    const el = n as Element;
    if (el.nodeType !== 1) continue;
    if (el.localName === 'r' || el.localName === 'fld') s += el.getElementsByTagNameNS(NS.a, 't')[0]?.textContent ?? '';
    else if (el.localName === 'br') s += '⏎';
  }
  return s;
}

/** Text of every paragraph on every slide, in order: `["1: Title", "2: OEE: 85%", ...]`. */
export async function slideTexts(bytes: Uint8Array): Promise<string[]> {
  const zip = await unzip(bytes);
  const slides = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => Number(/(\d+)\.xml$/.exec(a)![1]) - Number(/(\d+)\.xml$/.exec(b)![1]));
  const out: string[] = [];
  for (const [i, path] of slides.entries()) {
    const doc = await partXml(zip, path);
    for (const p of Array.from(doc.getElementsByTagNameNS(NS.a, 'p'))) {
      const t = paragraphText(p);
      if (t !== '') out.push(`${i + 1}: ${t}`);
    }
  }
  return out;
}

/** The `c:ser` elements of a chart part, serialized with insignificant whitespace removed. */
export async function chartSeriesXml(bytes: Uint8Array, chartPath: string): Promise<string[]> {
  const doc = await partXml(await unzip(bytes), chartPath);
  return Array.from(doc.getElementsByTagNameNS(NS.c, 'ser')).map((ser) =>
    new XMLSerializer().serializeToString(ser).replace(/>\s+</g, '><').replace(/ xmlns:\w+="[^"]*"/g, ''),
  );
}

/** Builds a minimal `a:p` element for unit tests. */
export function paragraph(inner: string): Element {
  const xml = `<a:p xmlns:a="${NS.a}">${inner}</a:p>`;
  return new DOMParser().parseFromString(xml, 'text/xml').documentElement!;
}

export function serialize(el: Element): string {
  return new XMLSerializer().serializeToString(el).replace(/ xmlns:a="[^"]*"/, '');
}
