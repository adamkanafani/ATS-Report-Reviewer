/**
 * Photo-table caption cleanup (Brett, 2026-11-10). Rule-based, no AI:
 *  - "Classification" label -> "Observation" (the two MDI software versions name that row
 *    differently; ATS standardizes on Observation).
 *  - A blank Observation value -> "Typical" (red, so the reviewer sees it was filled in).
 *  - Dashes in the Comments value removed ("-" alone, or "- text" / "text -").
 *  - A measurement row with no value ("Measurement Plane: --" -- the borescope writes "--" for
 *    none -- or blank) is deleted. Rows with a value (Point-to-Line 0.073 in) stay.
 *  - Photo file names left under the photos ("Stage_1_Rotor_Blade_..._001.JPG") are removed.
 */
import { childrenThroughWrappers, descendants, firstChild, textOf } from "../docx/ooxml";
import { DocModel, normalizeHeading } from "./model";
import { applyRedline, readParagraph } from "./redline";
import type { ReviewLog } from "./report";

const FILE_NAME = /^\s*[\w\-(). ]{1,150}\.(jpe?g|png|bmp|tiff?|gif|heic)\s*$/i;
// The caption rows every photo has; any other labeled row is a measurement row the borescope adds.
const CORE_LABELS = new Set(["section", "component", "location", "observation", "classification", "comments", "comment", "serial number"]);
// Measurement rows are named for what they measure ("Point-to-Line", "Measurement Plane", "Depth").
const MEASUREMENT_LABEL = /measure|plane|point|line|depth|length|width|gap|distance|height|thickness|diameter|radius|angle|area/;
const NO_VALUE = /^[\s\-–—.]*$/;
const DASHES = /^\s*[-–—]+\s*|\s*[-–—]+\s*$/g;

function label(cell: Element): string {
  return textOf(cell).trim().toLowerCase().replace(/:$/, "");
}

export function cleanPhotoCaptions(model: DocModel, log: ReviewLog) {
  let renamed = 0;
  let typical = 0;
  let dashes = 0;
  let planes = 0;
  let fileNames = 0;

  for (const table of model.photoTables) {
    // Operational data / data plate tables hold readings, where "--" means "not recorded" -- keep them.
    const readings = table.isOperationalData || table.isDataPlate;
    for (const row of descendants(table.topTbl, "w:tr")) {
      const cells = childrenThroughWrappers(row, "w:tc");
      if (cells.length < 2) continue;
      const name = label(cells[0]);
      const values = cells.slice(1);
      const valueText = values.map((c) => textOf(c).trim()).join("");

      if (name === "classification") {
        for (const p of descendants(cells[0], "w:p")) {
          const text = readParagraph(p).text;
          if (/classification/i.test(text) && applyRedline(p, text.replace(/classification/i, "Observation"), false).applied) renamed++;
        }
      }

      // Some MDI exports put a lone "-" in the Observation value; treat that as blank.
      if ((name === "observation" || name === "classification") && !valueText.replace(/[-–—\s]/g, "")) {
        const paras = values.flatMap((c) => descendants(c, "w:p"));
        const target = paras.find((p) => readParagraph(p).text.trim()) ?? descendants(values[values.length - 1], "w:p")[0];
        if (target && applyRedline(target, "Typical").applied) {
          typical++;
          for (const p of paras) {
            if (p === target) continue;
            const text = readParagraph(p).text;
            if (text.trim()) applyRedline(p, "");
          }
        }
      } else if (name === "observation" || name === "classification") {
        for (const cell of values) {
          for (const p of descendants(cell, "w:p")) {
            const text = readParagraph(p).text;
            const cleaned = text.replace(DASHES, "");
            if (cleaned !== text && applyRedline(p, cleaned).applied) dashes++;
          }
        }
      }

      if (name === "comments" || name === "comment") {
        for (const cell of values) {
          for (const p of descendants(cell, "w:p")) {
            const text = readParagraph(p).text;
            const cleaned = text.replace(DASHES, "");
            if (cleaned !== text && applyRedline(p, cleaned).applied) dashes++;
          }
        }
      }

      const isPhotoRow = descendants(row, "w:drawing").length > 0 || descendants(row, "w:pict").length > 0;
      if (!readings && name && !CORE_LABELS.has(name) && MEASUREMENT_LABEL.test(name) && NO_VALUE.test(valueText) && !isPhotoRow) {
        const tbl = row.parentNode as Element | null;
        if (tbl && childrenThroughWrappers(tbl, "w:tr").length > 1) {
          tbl.removeChild(row);
          planes++;
        }
      }
    }
  }

  // File names under the photos: in the Photos section or inside photo tables.
  for (const p of model.paragraphs) {
    const inPhotos = p.table?.kind === "photo" || (p.zone === "body" && p.h1 && normalizeHeading(p.h1) === "photos");
    if (!inPhotos || !p.el.parentNode) continue;
    const read = readParagraph(p.el);
    const visible = read.text.replace(/⟦\d+⟧/g, "");
    if (!FILE_NAME.test(visible)) continue;
    if (!applyRedline(p.el, read.text.replace(visible.trim(), "").trim()).applied) continue;
    fileNames++;
    // Drop the now-empty paragraph unless something depends on it: a picture, a section break,
    // or being the only paragraph left in a table cell (Word requires one).
    const parent = p.el.parentNode as Element;
    const pPr = firstChild(p.el, "w:pPr");
    const keep =
      descendants(p.el, "w:drawing").length > 0 ||
      descendants(p.el, "w:pict").length > 0 ||
      (pPr && firstChild(pPr, "w:sectPr")) ||
      (parent.tagName === "w:tc" && descendants(parent, "w:p").length <= 1);
    if (!keep && !textOf(p.el).trim()) parent.removeChild(p.el);
  }

  const notes: string[] = [];
  if (renamed) notes.push(`renamed ${renamed} "Classification" label${renamed === 1 ? "" : "s"} to "Observation"`);
  if (typical) notes.push(`filled ${typical} blank Observation${typical === 1 ? "" : "s"} with "Typical" (in red)`);
  if (dashes) notes.push(`removed the dash from ${dashes} Comments/Observation entr${dashes === 1 ? "y" : "ies"}`);
  if (planes) notes.push(`deleted ${planes} empty measurement row${planes === 1 ? "" : "s"} (Measurement Plane etc.)`);
  if (fileNames) notes.push(`removed ${fileNames} photo file name${fileNames === 1 ? "" : "s"} left under the photos`);
  if (notes.length) {
    const text = notes.join("; ");
    log.add("photos", { note: `Photo tables: ${text[0].toUpperCase()}${text.slice(1)}.` });
  }
}
