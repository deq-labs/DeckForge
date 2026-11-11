/** A single cell or KPI value as it arrives from a data source. */
export type Scalar = string | number | boolean | Date | null;

/** A text-tag value. `previous` feeds `{{key.previous}}`, `.change`, `.change_pct`, `.trend`. */
export interface ValueEntry {
  value: Scalar;
  previous?: Scalar;
  /** `text`, `number[:d]`, `percent[:d]`, `currency[:d]`, `date`; empty = automatic. */
  format?: string;
}

export interface TableData {
  /** Header texts; matched to the template's header row case-insensitively. */
  columns: string[];
  rows: Scalar[][];
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

/** Everything needed to fill one deck. Keys are slot names, matched case-insensitively. */
export interface DeckData {
  values?: Record<string, ValueEntry | Scalar>;
  tables?: Record<string, TableData>;
  charts?: Record<string, ChartData>;
}

export type OnMissing = 'error' | 'keep' | 'blank';

export interface FillOptions {
  /** BCP 47 locale for numbers and dates. Default `en-US`. */
  locale?: string;
  /** ISO 4217 code for the `currency` format. Default `USD`. */
  currency?: string;
  /** What to do when a tag has no data. Default `error`. */
  onMissing?: OnMissing;
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
  | 'unused-data';

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
  slides: number[];
}

export interface TableSlot {
  kind: 'table';
  name: string;
  slide: number;
  columns: string[];
  /** Body rows as currently written in the template (useful as sample data). */
  sampleRows: string[][];
}

export type ChartType = 'bar' | 'line';

export interface ChartSlot {
  kind: 'chart';
  name: string;
  slide: number;
  /** Supported chart type, or undefined when the chart can't be filled (see issues). */
  chartType: ChartType | undefined;
  /** The template's own data, read from the chart caches. */
  sample: ChartData;
}

export type Slot = TextSlot | TableSlot | ChartSlot;

export interface Manifest {
  slideCount: number;
  text: TextSlot[];
  tables: TableSlot[];
  charts: ChartSlot[];
  /** Problems found in the template itself (bad tag names, unsupported parts). */
  issues: Issue[];
}
