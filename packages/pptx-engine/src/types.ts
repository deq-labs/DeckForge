/** A single cell or KPI value as it arrives from a data source. */
export type Scalar = string | number | boolean | Date | null;

/** A text-tag value. `previous` feeds `{{key.previous}}`, `.change`, `.change_pct`, `.trend`. */
export interface ValueEntry {
  value: Scalar;
  previous?: Scalar;
  /** `text`, `number[:d]`, `percent[:d]`, `currency[:d]`, `date`; empty = automatic. */
  format?: string;
}

/** A table cell: a bare value, or a value with a format (see `ValueEntry.format`). */
export type TableCell = Scalar | { value: Scalar; format?: string };

export interface TableData {
  /** Header texts; matched to the template's header row case-insensitively. */
  columns: string[];
  rows: TableCell[][];
}

export interface ChartSeries {
  name: string;
  values: (number | null)[];
}

export interface ChartData {
  /** Label for the category column in the embedded workbook (cell A1). */
  categoryLabel?: string;
  categories: Scalar[];
  series: ChartSeries[];
}

/** One row of a repeat source: tag name -> value. */
export type RowRecord = Record<string, ValueEntry | Scalar>;

/** Everything needed to fill one deck. Keys are slot names, matched case-insensitively. */
export interface DeckData {
  values?: Record<string, ValueEntry | Scalar>;
  tables?: Record<string, TableData>;
  charts?: Record<string, ChartData>;
  /** Row sets for `#repeat <name>` slides. */
  rows?: Record<string, RowRecord[]>;
}

export type OnMissing = 'error' | 'keep' | 'blank';

export interface FillOptions {
  /** BCP 47 locale for numbers and dates. Default `en-US`. */
  locale?: string;
  /** ISO 4217 code for the `currency` format. Default `USD`. */
  currency?: string;
  /** What to do when a tag has no data. Default `error`. */
  onMissing?: OnMissing;
  /** If set, this text is added in small grey type to the bottom-right of every slide (trial mode). */
  watermark?: string;
}

export type IssueCode =
  | 'invalid-tag'
  | 'duplicate-slot'
  | 'unsupported'
  | 'missing-value'
  | 'missing-table'
  | 'missing-chart'
  | 'bad-value'
  | 'table-columns'
  | 'chart-series'
  | 'unused-data'
  | 'data-layout'
  | 'missing-rows'
  | 'slide-rule'
  | 'misplaced-tag'
  | 'overflow';

export interface Issue {
  severity: 'error' | 'warning';
  /** Stable machine-readable category, so apps can add their own hints. */
  code: IssueCode;
  /** 1-based slide number, when the issue is tied to a slide. */
  slide?: number;
  /** Slot name, when the issue is tied to a slot. */
  slot?: string;
  /** Plain-language message: what is wrong and how to fix it. */
  message: string;
}

export const TEXT_FIELDS = ['value', 'previous', 'change', 'change_pct', 'trend'] as const;
export type TextField = (typeof TEXT_FIELDS)[number];

export interface TextSlot {
  kind: 'text';
  name: string;
  /** Fields used anywhere in the deck, e.g. `value` for `{{oee}}`, `previous` for `{{oee.previous}}`. */
  fields: TextField[];
  /** Slides that use the tag (template numbering), excluding repeated slides. */
  slides: number[];
  /** Names of slide layouts/masters that use the tag. */
  layouts: string[];
}

/** A `#repeat <name>` source: the slides it copies and the tags those slides use. */
export interface RepeatSlot {
  kind: 'repeat';
  name: string;
  slides: number[];
  tags: { name: string; fields: TextField[] }[];
}

export interface SlideRule {
  slide: number;
  kind: 'repeat' | 'show-if' | 'hide-if';
  target: string;
}

export interface TableSlot {
  kind: 'table';
  name: string;
  slide: number;
  columns: string[];
  /** Body rows as currently written in the template (useful as sample data). */
  sampleRows: string[][];
  /** False when the table can't be filled (e.g. merged body cells); see the manifest issues. */
  fillable: boolean;
}

export type ChartType = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'combo';

export interface ChartSlot {
  kind: 'chart';
  name: string;
  slide: number;
  /** Supported chart type, or undefined when the chart can't be filled (see issues). */
  chartType: ChartType | undefined;
  /** The template's own data, read from the chart caches. */
  sample: ChartData;
}

export type Slot = TextSlot | TableSlot | ChartSlot | RepeatSlot;

export interface Manifest {
  slideCount: number;
  text: TextSlot[];
  tables: TableSlot[];
  charts: ChartSlot[];
  repeats: RepeatSlot[];
  rules: SlideRule[];
  /** Slide height in EMU, for overflow checks. */
  slideHeight: number;
  /** Problems found in the template itself (bad tag names, unsupported parts). */
  issues: Issue[];
}
