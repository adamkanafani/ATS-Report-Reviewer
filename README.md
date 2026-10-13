# ATS Report Reviewer

Reviews an Advanced Turbine Support inspection report (.docx) as it comes from the inspector,
especially one-off inspections written without a template, and produces a corrected copy for
the report reviewer. Same shell and look as the ATS Borescope Report Builder.

## What it does to the report

| Check | How |
|---|---|
| Formatting (Report Composition standards) | Times New Roman throughout, 12 pt body paragraphs / 11 pt Observations tables (other tables keep their size), 7" tables, centered Percent Inspected cells, standard photo size |
| Spell check and grammar | Claude AI, applied as word-level edits |
| Photo tables vs. Observations tables | Claude AI; mismatches are left as Word comments |
| Overall Assessment readability | Claude AI; edits applied, recommendations left as a Word comment |
| Heading styles | Fixed rules for the standard section names, plus an AI pass for one-off section names |
| Photo-table text | Times New Roman 12 |
| Photo section layout | Narrow (0.5") margins on every photo page; restores a deleted section break before Photos (with the body's standard margins and page header); removes photo-table cells hidden in the margin |
| Cover page | Title 44 pt, other lines 22 pt |
| Table of Contents | Rebuilt by the installed Microsoft Word, set to TNR 12, 1.5 spacing, black |
| Photo resolution | Every photo resampled to 150 ppi at its displayed size, operational data photo(s) excluded; the document's default resolution is also set to 150 ppi |
| Text color | Everything black, **except AI edits, which are red** for review |

### Several reports for the same unit

Tick several reports in the file browser (or upload/drop several) to review them together --
e.g. the borescope, PT, ECI and UT reports for one unit. Each gets the full review above and
downloads under its own name, and the reports are also **compared with each other**: the
Inspection Details (Utility, Site, Unit, Serial Number, Type of Fuel, Site Contact, operational
data), the visual observations for components more than one report covers (Inlet, VIGV, first
compressor stages) and the significant findings for them must match. Each mismatch becomes a
Word comment in every report involved and is listed on the results screen's Report Comparison
tab.

It also lists template placeholders still in the report (X, XX, 20XX, "Choose an item.").
It never fills those in. The original file is never modified. The reviewed copy downloads
under its original file name (to the browser's Downloads folder), and the app shows a read-only log of every change.

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
