# ATS Report Reviewer

Reviews an Advanced Turbine Support inspection report (.docx) as it comes from the inspector,
especially one-off inspections written without a template, and produces a corrected copy for
the report reviewer. Same shell and look as the ATS Borescope Report Builder.

## What it does to the report

| Check | How |
|---|---|
| Formatting (Report Composition standards) | Times New Roman throughout, 12 pt body / 11 pt Observations tables, 7" tables, centered Percent Inspected cells, standard photo size |
| Spell check and grammar | Claude AI, applied as word-level edits |
| Photo tables vs. Observations tables | Claude AI; mismatches are left as Word comments |
| Overall Assessment readability | Claude AI; edits applied, recommendations left as a Word comment |
| Heading styles | Fixed rules for the standard section names, plus an AI pass for one-off section names |
| Photo-table text | Times New Roman 12 |
| Table of Contents | Rebuilt by the installed Microsoft Word, set to TNR 12, 1.5 spacing, black |
| Photo resolution | Every photo resampled to 150 ppi at its displayed size, operational data photo(s) excluded; the document's default resolution is also set to 150 ppi |
| Text color | Everything black, **except AI edits, which are red** for review |

It also lists template placeholders still in the report (X, XX, 20XX, "Choose an item.").
It never fills those in. The original file is never modified. The reviewed copy downloads as
`<name> - Reviewed.docx`, and the app shows a read-only log of every change.

## Setup (each computer)

1. Install Node.js LTS from https://nodejs.org (default options).
2. Double-click **`Setup - Run This First.bat`**. It asks for the Anthropic API key, which it
   saves to `.env.local` (see `.env.local.example`). Press Enter to skip; the AI checks are
   then turned off.
3. Day to day: double-click **`Start Report Reviewer.bat`**. Leave the server window open, and
   the app opens at http://localhost:3200.
4. To update: **`Update.bat`**.

Microsoft Word must be installed for the Table of Contents to be rebuilt automatically. Without
it, the report is flagged so Word offers to update fields when the file is opened.

## Development

```bash
npm run dev
```

```bash
npm run test:review -- "reference/<report>.docx" --seed-errors --mock-ai
```

`--seed-errors` injects typos, wrong heading styles, and colored text into a copy of the input.
`--mock-ai` swaps Claude for canned responses so the redline and comment path can be tested
without an API key. `--no-ai` skips the AI passes. Output goes to `test-output/`.

Code map: `src/lib/review/pipeline.ts` orchestrates. `formatting.ts` holds the deterministic
rules. `ai.ts` makes the Claude calls (model `claude-opus-5-5`, structured outputs).
`redline.ts` applies red word-level edits. `images.ts` handles 150 ppi resampling. `word.ts`
rebuilds the TOC through Word COM.
