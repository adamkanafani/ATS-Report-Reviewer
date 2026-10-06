/** Shared between the review engine and the results screen. */

export type ReviewCategoryKey =
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
};

export interface ReviewResult {
  jobId: string;
  outputFileName: string;
  categories: ReviewCategory[];
  aiEditCount: number;
  commentCount: number;
}

export type ProgressEvent =
  | { type: "step"; message: string }
  | { type: "done"; result: ReviewResult }
  | { type: "error"; message: string };
