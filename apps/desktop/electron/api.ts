/**
 * The contract between the renderer (React UI) and the main process. Plain JSON-able types
 * only, no Node imports, so the renderer can share this file.
 */

export type Severity = 'error' | 'warning';

export interface ProblemView {
  severity: Severity;
  code: string;
  message: string;
  hint?: string;
  slot?: string;
  batch?: string;
}

export interface SlotView {
  kind: 'text' | 'table' | 'chart' | 'repeat';
  name: string;
  /** Where the slot is, e.g. "slide 2" or "layout 'Title Slide'". */
  where: string;
  /** What it needs from the workbook, e.g. "Line | Cause | Minutes" or "pie, 1 series". */
  detail: string;
  /** False when the template part can't be filled (unsupported chart, merged cells...). */
  fillable: boolean;
}

export interface TemplateInfo {
  path: string;
  fileName: string;
  slideCount: number;
  slots: SlotView[];
  rules: string[];
  problems: ProblemView[];
}

export type MissingMode = 'error' | 'keep' | 'blank';

export interface JobOptions {
  locale: string;
  currency: string;
  missing: MissingMode;
}

export interface JobSpec {
  templatePath: string;
  dataPath: string;
  outDir: string;
  options: JobOptions;
  /** `{template}` and `{batch}` are replaced. */
  namePattern?: string;
}

export interface CheckView {
  ok: boolean;
  problems: ProblemView[];
  batches: string[];
  deckCount: number;
}

export interface GenerateView extends CheckView {
  outDir: string;
  /** Absolute paths of the decks written. */
  files: string[];
  watermarked: boolean;
}

export interface Recipe {
  id: string;
  name: string;
  spec: JobSpec;
  lastRunAt?: string;
  lastDeckCount?: number;
}

export type PreviewTool = 'powerpoint' | 'libreoffice';

export interface LicenseView {
  mode: 'trial' | 'licensed';
  /** Shown in the header, e.g. "Trial — output has a small DeckForge footer". */
  summary: string;
  licensee?: string;
  seats?: number;
  /** ISO date when free updates end. */
  updatesUntil?: string;
  /** Present when the key on file is invalid or doesn't cover this version. */
  problem?: string;
}

export interface DeckforgeApi {
  pickTemplate(): Promise<string | null>;
  pickData(): Promise<string | null>;
  pickFolder(defaultPath?: string): Promise<string | null>;
  inspect(templatePath: string): Promise<TemplateInfo>;
  /** Asks where to save, writes a starter workbook for the template, opens it. */
  createDataFile(templatePath: string): Promise<string | null>;
  check(spec: JobSpec): Promise<CheckView>;
  generate(spec: JobSpec): Promise<GenerateView>;
  openPath(path: string): Promise<void>;
  showInFolder(path: string): Promise<void>;
  previewTool(): Promise<PreviewTool | null>;
  /** Renders the deck to PDF with PowerPoint or LibreOffice and opens it in a viewer window. */
  preview(deckPath: string): Promise<void>;
  recipes: {
    list(): Promise<Recipe[]>;
    save(recipe: Recipe): Promise<Recipe[]>;
    remove(id: string): Promise<Recipe[]>;
  };
  license: {
    status(): Promise<LicenseView>;
    /** Activates a license key (pasted text or the contents of a license file). */
    activate(key: string): Promise<LicenseView>;
    deactivate(): Promise<LicenseView>;
  };
}
