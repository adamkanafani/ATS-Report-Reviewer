/** Shared between the review engine and the results screen. */

export type ReviewCategoryKey =
  | "crossReport"
  | "proofreading"
  | "overallAssessment"
  | "crossCheck"
  | "headings"
  | "formatting"
  | "photos"
  | "toc"
  | "placeholders";

export interface ReviewItem {
  /** Where in the report, e.g. "Observations › Compressor Section › table row 3". */
  location?: string;
  before?: string;
  after?: string;
  /** Plain-language description of the change or finding. */
  note: string;
  /** "info" items are routine; "attention" items need Brett's eyes in the output file. */
  level?: "info" | "attention";
}

export interface ReviewCategory {
  key: ReviewCategoryKey;
  title: string;
  /** One-line explanation shown above the item list. */
  summary: string;
  items: ReviewItem[];
  /** Set when the step couldn't run (no API key, Word missing, ...). */
  skippedReason?: string;
}

export interface ReviewOptions {
  formatting: boolean;
  proofreading: boolean;
  crossCheck: boolean;
  overallAssessment: boolean;
  headings: boolean;
  photoTables: boolean;
  toc: boolean;
  photoResolution: boolean;
  /** With two or more reports: check that the details/observations that must match do. */
  compareReports: boolean;
}

export const DEFAULT_OPTIONS: ReviewOptions = {
  formatting: true,
  proofreading: true,
  crossCheck: true,
  overallAssessment: true,
  headings: true,
  photoTables: true,
  toc: true,
  photoResolution: true,
  compareReports: true,
};

export interface ReviewResult {
  jobId: string;
  outputFileName: string;
  /** Set when this report couldn't be reviewed (the others in the batch still were). */
  error?: string;
  categories: ReviewCategory[];
  aiEditCount: number;
  commentCount: number;
}

export type ProgressEvent =
  | { type: "step"; message: string }
  | { type: "done"; results: ReviewResult[]; comparison: ReviewCategory | null }
  | { type: "error"; message: string };
