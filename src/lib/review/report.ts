import type { ReviewCategory, ReviewCategoryKey, ReviewItem } from "./types";

const CATEGORY_META: Record<ReviewCategoryKey, { title: string; summary: string }> = {
  proofreading: {
    title: "Spelling & Grammar",
    summary: "Spelling, grammar, and ATS observation-style corrections. Every changed word is red in the report.",
  },
  overallAssessment: {
    title: "Overall Assessment",
    summary: "Readability edits (red in the report) and recommendations (also left as a Word comment on the heading).",
  },
  crossCheck: {
    title: "Photo vs. Observation Check",
    summary: "Where the photo tables and the Observations tables disagree. Each one is also a Word comment in the report.",
  },
  headings: {
    title: "Heading Styles",
    summary: "Paragraphs whose Style selection didn't match the Report Composition rules.",
  },
  formatting: {
    title: "Formatting",
    summary: "Fonts, sizes, table widths, and text color brought in line with the Report Composition standards.",
  },
  photos: {
    title: "Photos",
    summary: "Photo-table text set to Times New Roman 12 and photo resolution set to 150 ppi (operational data photos left untouched).",
  },
  toc: {
    title: "Table of Contents",
    summary: "Table of Contents update and formatting (Times New Roman 12, 1.5 spacing, black).",
  },
  placeholders: {
    title: "Unfilled Placeholders",
    summary: "Template placeholder text still in the report (X, XX, 20XX, \"Choose an item.\"). Nothing was changed here -- these need the inspector's actual values.",
  },
};

export const CATEGORY_ORDER: ReviewCategoryKey[] = [
  "crossCheck",
  "overallAssessment",
  "proofreading",
  "placeholders",
  "headings",
  "formatting",
  "photos",
  "toc",
];

export class ReviewLog {
  private categories = new Map<ReviewCategoryKey, ReviewCategory>();

  private get(key: ReviewCategoryKey): ReviewCategory {
    let cat = this.categories.get(key);
    if (!cat) {
      cat = { key, ...CATEGORY_META[key], items: [] };
      this.categories.set(key, cat);
    }
    return cat;
  }

  add(key: ReviewCategoryKey, item: ReviewItem) {
    this.get(key).items.push(item);
  }

  skip(key: ReviewCategoryKey, reason: string) {
    this.get(key).skippedReason = reason;
  }

  touch(key: ReviewCategoryKey) {
    this.get(key);
  }

  toArray(): ReviewCategory[] {
    return CATEGORY_ORDER.filter((k) => this.categories.has(k)).map((k) => this.categories.get(k)!);
  }
}
