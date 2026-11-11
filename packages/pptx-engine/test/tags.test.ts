import { describe, expect, it } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import type { Element } from '@xmldom/xmldom';
import { findTags, parseTag, removeLines, replaceTags } from '../src/tags.ts';
import { NS } from '../src/xml.ts';
import { paragraph, serialize } from './helpers.ts';

const r = (text: string, color?: string): string =>
  `<a:r>${color ? `<a:rPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>` : '<a:rPr/>'}<a:t>${text}</a:t></a:r>`;

const values: Record<string, string> = { oee: '85%', site: 'Plant A', a: 'X', b: 'Y', multi: 'line 1\nline 2' };
const resolve = (inner: string): string | undefined => {
  const t = parseTag(inner);
  return t.ok ? values[t.name] : undefined;
};

describe('parseTag', () => {
  it('accepts names, spaces and suffixes', () => {
    expect(parseTag(' oee ')).toEqual({ ok: true, name: 'oee', field: 'value' });
    expect(parseTag('oee.previous')).toEqual({ ok: true, name: 'oee', field: 'previous' });
    expect(parseTag('scrap-rate.TREND')).toEqual({ ok: true, name: 'scrap-rate', field: 'trend' });
  });
  it('rejects empty, spaced and unknown-suffix tags with a reason', () => {
    expect(parseTag('  ').ok).toBe(false);
    expect(parseTag('bad name').ok).toBe(false);
    expect(parseTag('1st').ok).toBe(false);
    const t = parseTag('oee.prev');
    expect(t.ok).toBe(false);
    if (!t.ok) expect(t.reason).toContain('.previous');
  });
});

describe('findTags', () => {
  it('finds tags split across runs but not across line breaks', () => {
    const p = paragraph(r('Site: {{si') + r('te}} and {{a') + '<a:br/>' + r('}}'));
    expect(findTags(p)).toEqual(['site']);
  });
});

describe('replaceTags', () => {
  it('replaces a tag inside one run', () => {
    const p = paragraph(r('OEE: {{oee}} today'));
    replaceTags(p, resolve);
    expect(serialize(p)).toBe('<a:p><a:r><a:rPr/><a:t>OEE: 85% today</a:t></a:r></a:p>');
  });

  it('rejoins a tag split across three runs and keeps the first run formatting', () => {
    const p = paragraph(r('OEE: {{o', 'FF0000') + r('e', '00FF00') + r('e}} now', '0000FF'));
    replaceTags(p, resolve);
    expect(serialize(p)).toBe(
      '<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>OEE: 85%</a:t></a:r>' +
        '<a:r><a:rPr><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t> now</a:t></a:r></a:p>',
    );
  });

  it('handles several tags, including two that share a run boundary', () => {
    const p = paragraph(r('{{a}}{{') + r('b}}-{{site}}'));
    replaceTags(p, resolve);
    expect(serialize(p)).toBe('<a:p><a:r><a:rPr/><a:t>XY</a:t></a:r><a:r><a:rPr/><a:t>-Plant A</a:t></a:r></a:p>');
  });

  it('leaves unknown and invalid tags exactly as written', () => {
    const p = paragraph(r('{{nope}} {{bad name}} {single}'));
    replaceTags(p, resolve);
    expect(serialize(p)).toBe('<a:p><a:r><a:rPr/><a:t>{{nope}} {{bad name}} {single}</a:t></a:r></a:p>');
  });

  it('removes a run that only held part of a tag', () => {
    const p = paragraph(r('{{', 'FF0000') + r('oee') + r('}}'));
    replaceTags(p, resolve);
    expect(serialize(p)).toBe('<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>85%</a:t></a:r></a:p>');
  });

  it('turns newlines in values into line breaks with the same formatting', () => {
    const p = paragraph(r('>{{multi}}<', 'FF0000'));
    replaceTags(p, resolve);
    const rpr = '<a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>';
    expect(serialize(p)).toBe(`<a:p><a:r>${rpr}<a:t>&gt;line 1</a:t></a:r><a:br>${rpr}</a:br><a:r>${rpr}<a:t>line 2&lt;</a:t></a:r></a:p>`);
  });

  it('escapes XML special characters in values', () => {
    const p = paragraph(r('{{site}}'));
    replaceTags(p, () => 'R&D <east>');
    expect(serialize(p)).toContain('<a:t>R&amp;D &lt;east&gt;</a:t>');
  });
});

describe('removeLines', () => {
  it('drops matching lines split by breaks or newlines', () => {
    const p = paragraph(r('#repeat sites') + '<a:br/>' + r('Keep me'));
    removeLines(p, (l) => l.startsWith('#'));
    expect(serialize(p)).toBe('<a:p><a:r><a:rPr/><a:t>Keep me</a:t></a:r></a:p>');
  });

  it('empties, but keeps, the only paragraph of a text body', () => {
    const body = new DOMParser().parseFromString(`<a:txBody xmlns:a="${NS.a}"><a:p>${r('#repeat x')}<a:endParaRPr/></a:p></a:txBody>`, 'text/xml').documentElement!;
    removeLines(body.firstChild as Element, (l) => l.startsWith('#'));
    expect(new XMLSerializer().serializeToString(body)).toBe(`<a:txBody xmlns:a="${NS.a}"><a:p><a:endParaRPr/></a:p></a:txBody>`);
  });

  it('removes a directive-only paragraph when others remain', () => {
    const body = new DOMParser().parseFromString(`<a:txBody xmlns:a="${NS.a}"><a:p>${r('#hide-if x')}</a:p><a:p>${r('Note')}</a:p></a:txBody>`, 'text/xml').documentElement!;
    removeLines(body.firstChild as Element, (l) => l.startsWith('#'));
    expect(body.getElementsByTagNameNS(NS.a, 'p').length).toBe(1);
  });
});
