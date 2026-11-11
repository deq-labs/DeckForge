/** Data-size fuzzing: 0, 1 and 500 rows/points must all produce a deck that reads back correctly. */
import { describe, expect, it } from 'vitest';
import { Template } from '../src/index.ts';
import type { DeckData } from '../src/index.ts';
import { fixture } from './helpers.ts';

function rng(seed: number): () => number {
  // mulberry32: deterministic so failures reproduce
  return () => {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function data(n: number, seed: number): DeckData {
  const rand = rng(seed);
  const num = (): number | null => (rand() < 0.1 ? null : Math.round(rand() * 10000) / 100);
  return {
    values: { title: 'Fuzz', site: `S${seed}`, period_end: 'today', downtime_total: num() },
    tables: {
      downtime: {
        columns: ['Line', 'Cause', 'Minutes'],
        rows: Array.from({ length: n }, (_, i) => [`Line ${i}`, rand() < 0.5 ? 'Jam & <stop>' : '', num()]),
      },
    },
    charts: {
      downtime_by_line: {
        categories: Array.from({ length: n }, (_, i) => `L${i}`),
        series: [{ name: 'Planned', values: Array.from({ length: n }, num) }],
      },
      oee_trend: {
        categories: Array.from({ length: n }, (_, i) => i),
        series: [
          { name: 'OEE', values: Array.from({ length: n }, num) },
          { name: 'Target', values: Array.from({ length: n }, () => 0.75) },
        ],
      },
    },
  };
}

describe('fuzz: data sizes', () => {
  for (const n of [0, 1, 2, 500]) {
    it(`${n} rows / points`, async () => {
      const input = data(n, n + 1);
      const tpl = await Template.load(fixture('ops-review.pptx'));
      const deck = await tpl.fill(input);
      const back = (await Template.load(deck.toBuffer())).slots();

      expect(back.tables[0]!.sampleRows).toHaveLength(n);
      for (const chart of back.charts) {
        const sent = input.charts![chart.name]!;
        expect(chart.sample.categories).toEqual(sent.categories.map(String));
        expect(chart.sample.series).toEqual(sent.series);
      }
    });
  }

  it('fills a 500-row deck quickly', async () => {
    const tpl = await Template.load(fixture('ops-review.pptx'));
    const start = performance.now();
    await tpl.fill(data(500, 7));
    expect(performance.now() - start).toBeLessThan(5000);
  });
});
