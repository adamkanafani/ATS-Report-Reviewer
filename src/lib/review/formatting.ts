/**
 * Deterministic (non-AI) formatting pass -- the parts of Brett's Report Composition standards
 * that are mechanical enough to just apply: fonts, sizes, text color, 7" table widths, centered
 * Percent Inspected cells, standard photo size, heading styles for the standard section names,
 * and the Table of Contents styles. Runs before the AI pass, so the only colored text left in the
 * finished report is the AI's own red edits.
 */
import {
  createW,
  descendants,
  elementChildren,
  ensureParagraphMarkRPr,
  ensurePropsFirst,
  firstChild,
  forceSize,
  forceTimesNewRoman,
  getVal,
  PPR_ORDER,
  RPR_ORDER,
  setAttr,
  setOrderedChild,
  STYLE_ORDER,
  stripColor,
  TCPR_ORDER,
  TBLPR_ORDER,
} from "../docx/ooxml";
import { DocModel, locationLabel, normalizeHeading, ParaInfo, PHOTO_CAPTION_LABELS, refreshHeadingContext } from "./model";
import type { ReviewLog } from "./report";
import type { ReviewOptions } from "./types";

const SEVEN_INCHES_DXA = 10080;
const EMU_PER_INCH = 914400;
const EMU_PER_DXA = 635;
const PHOTO_HEIGHT_EMU = Math.round(3.2 * EMU_PER_INCH);

const H1_NAMES = new Set([
  "documentation and photographs",
  "documents and photographs",
  "overall assessment",
  "inspection details",
  "significant observations and recommendations",
  "observations",
  "photos",
]);
const H2_NAMES = new Set(["purpose", "inspection areas", "applicable til's", "applicable tils", "inspection access"]);
const NORMAL_NAMES = new Set(["notes", "note"]);
const SECTION_NAME = /^(inlet|compressor|combustion|turbine|exhaust|generator|diffuser|bearing|load|auxiliary|accessory)\b/;

function runsWithText(p: Element): Element[] {
  return descendants(p, "w:r").filter((r) => descendants(r, "w:t").some((t) => (t.textContent ?? "").length > 0));
}

function enclosing(el: Element, tagName: string): Element | null {
  let node = el.parentNode as Element | null;
  while (node && node.nodeType === 1) {
    if (node.tagName === tagName) return node;
    node = node.parentNode as Element | null;
  }
  return null;
}

export interface FormattingContext {
  documentXml: Document;
  stylesXml: Document | null;
  model: DocModel;
  log: ReviewLog;
  options: ReviewOptions;
}

export function applyFormatting(ctx: FormattingContext): { hasToc: boolean } {
  const { options } = ctx;
  if (options.formatting) {
    applyColor(ctx);
    applyFonts(ctx);
    applyBodySizes(ctx);
    applyTableWidths(ctx);
    applyPercentCells(ctx);
  }
  if (options.photoTables) applyPhotoTables(ctx);
  if (options.headings) applyKnownHeadings(ctx);
  scanPlaceholders(ctx);
  const hasToc = descendants(ctx.documentXml.documentElement, "w:instrText").some((t) => /\bTOC\b/.test(t.textContent ?? ""));
  if (options.toc) applyTocFormatting(ctx, hasToc);
  return { hasToc };
}

// --- Text color: everything black (light-on-dark header text excepted) ---

function applyColor({ documentXml, stylesXml, log }: FormattingContext) {
  let runs = 0;
  for (const rPr of descendants(documentXml.documentElement, "w:rPr")) if (stripColor(rPr)) runs++;
  let styles = 0;
  if (stylesXml) for (const rPr of descendants(stylesXml.documentElement, "w:rPr")) if (stripColor(rPr)) styles++;
  if (runs || styles) {
    const parts = [
      runs ? `${runs} colored text run${runs === 1 ? "" : "s"}` : "",
      styles ? `${styles} colored style definition${styles === 1 ? "" : "s"}` : "",
    ].filter(Boolean);
    log.add("formatting", { note: `Changed ${parts.join(" and ")} to black.` });
  }
}

// --- Fonts: Times New Roman throughout ---

function applyFonts({ documentXml, stylesXml, log }: FormattingContext) {
  let runs = 0;
  for (const rPr of descendants(documentXml.documentElement, "w:rPr")) if (forceTimesNewRoman(rPr, false)) runs++;
  let styles = 0;
  if (stylesXml) {
    const root = stylesXml.documentElement;
    const rPrDefault = descendants(root, "w:rPrDefault")[0];
    if (rPrDefault && forceTimesNewRoman(ensurePropsFirst(rPrDefault, "w:rPr"), true)) styles++;
    for (const style of descendants(root, "w:style")) {
      const rPr = firstChild(style, "w:rPr");
      if (rPr && forceTimesNewRoman(rPr, false)) styles++;
    }
    // The Normal style anchors every inherited size: 12 pt per the standards.
    const normal = descendants(root, "w:style").find((s) => s.getAttribute("w:type") === "paragraph" && s.getAttribute("w:default") === "1");
    if (normal) {
      const rPr = setOrderedChild(normal, "w:rPr", STYLE_ORDER, {}, false);
      if (forceSize(rPr, 24)) styles++;
    }
  }
  if (runs || styles) {
    const parts = [runs ? `${runs} text run${runs === 1 ? "" : "s"}` : "", styles ? `${styles} style${styles === 1 ? "" : "s"}` : ""].filter(Boolean);
    log.add("formatting", { note: `Set ${parts.join(" and ")} to Times New Roman.` });
  }
}

// --- Sizes: 12 pt body text and tables, 11 pt Observations tables ---

function applyBodySizes({ model, log }: FormattingContext) {
  let body = 0;
  let obs = 0;
  for (const p of model.paragraphs) {
    if (p.zone !== "body" && p.zone !== "llc") continue;
    if (p.headingLevel) continue;
    if (p.table?.kind === "photo") continue; // handled by applyPhotoTables
    if (!p.table && p.h1 && normalizeHeading(p.h1) === "photos") continue;
    const target = p.table?.kind === "observation" ? 22 : 24;
    for (const run of runsWithText(p.el)) {
      const rPr = firstChild(run, "w:rPr");
      // Observations tables need an explicit 11 pt; everything else only needs a stray explicit
      // size corrected, since 12 pt is what Normal already gives it.
      if (target === 22) {
        if (forceSize(ensurePropsFirst(run, "w:rPr"), 22)) obs++;
      } else if (rPr && firstChild(rPr, "w:sz") && forceSize(rPr, 24)) body++;
    }
  }
  if (body) log.add("formatting", { note: `Corrected the font size of ${body} text run${body === 1 ? "" : "s"} to 12 pt.` });
  if (obs) log.add("formatting", { note: `Set ${obs} Observations-table text run${obs === 1 ? "" : "s"} to 11 pt.` });
}

// --- Table widths: 7 inches ---

function tableWidthDxa(tbl: Element): number | null {
  const grid = firstChild(tbl, "w:tblGrid");
  const tblW = firstChild(firstChild(tbl, "w:tblPr") ?? tbl, "w:tblW");
  if (tblW && tblW.getAttribute("w:type") === "dxa") {
    const w = Number(tblW.getAttribute("w:w"));
    if (w > 0) return w;
  }
  if (!grid) return null;
  const sum = elementChildren(grid)
    .filter((c) => c.tagName === "w:gridCol")
    .reduce((acc, c) => acc + Number(c.getAttribute("w:w") || 0), 0);
  return sum > 0 ? sum : null;
}

function scaleTable(tbl: Element, target: number): boolean {
  const current = tableWidthDxa(tbl);
  if (!current || Math.abs(current - target) / target < 0.01) return false;
  const factor = target / current;
  const tblPr = ensurePropsFirst(tbl, "w:tblPr");
  setOrderedChild(tblPr, "w:tblW", TBLPR_ORDER, { "w:w": String(target), "w:type": "dxa" });
  const grid = firstChild(tbl, "w:tblGrid");
  if (grid) {
    for (const col of elementChildren(grid)) {
      if (col.tagName !== "w:gridCol") continue;
      setAttr(col, "w:w", String(Math.round(Number(col.getAttribute("w:w") || 0) * factor)));
    }
  }
  // Only this table's own cells -- a nested table keeps its own widths.
  const cells = descendants(tbl, "w:tc").filter((tc) => enclosing(tc, "w:tbl") === tbl);
  for (const tc of cells) {
    const tcW = firstChild(firstChild(tc, "w:tcPr") ?? tc, "w:tcW");
    if (tcW && tcW.getAttribute("w:type") === "dxa") {
      setAttr(tcW, "w:w", String(Math.round(Number(tcW.getAttribute("w:w") || 0) * factor)));
    }
  }
  return true;
}

function applyTableWidths({ model, log }: FormattingContext) {
  const done = new Set<Element>();
  for (const p of model.paragraphs) {
    if (!p.table || done.has(p.table.topTbl)) continue;
    done.add(p.table.topTbl);
    if (p.zone !== "body" && p.zone !== "llc") continue;
    if (p.table.kind === "photo") continue;
    const before = tableWidthDxa(p.table.topTbl);
    if (scaleTable(p.table.topTbl, SEVEN_INCHES_DXA)) {
      log.add("formatting", {
        location: locationLabel(p),
        note: `Resized table from ${before ? (before / 1440).toFixed(2) : "?"}" to 7" wide.`,
      });
    }
  }
}

// --- Percent Inspected cells: centered horizontally and vertically ---

function applyPercentCells({ model, log }: FormattingContext) {
  let count = 0;
  for (const p of model.paragraphs) {
    if (p.table?.kind !== "observation") continue;
    if (!/^\s*\d{1,3}(\.\d+)?\s?%\s*$/.test(p.text)) continue;
    const pPr = ensurePropsFirst(p.el, "w:pPr");
    const jc = firstChild(pPr, "w:jc");
    const tc = enclosing(p.el, "w:tc");
    const tcPr = tc ? ensurePropsFirst(tc, "w:tcPr") : null;
    const vAlign = tcPr ? firstChild(tcPr, "w:vAlign") : null;
    if (getVal(jc) === "center" && (!tcPr || getVal(vAlign) === "center")) continue;
    setOrderedChild(pPr, "w:jc", PPR_ORDER, { "w:val": "center" });
    if (tcPr) setOrderedChild(tcPr, "w:vAlign", TCPR_ORDER, { "w:val": "center" });
    count++;
  }
  if (count) log.add("formatting", { note: `Centered ${count} Percent Inspected cell${count === 1 ? "" : "s"} horizontally and vertically.` });
}

// --- Photo tables: Times New Roman 12, standard 3.2" photo height ---

function applyPhotoTables({ model, log }: FormattingContext) {
  let runs = 0;
  for (const p of model.paragraphs) {
    if (p.table?.kind !== "photo") continue;
    const textRuns = runsWithText(p.el);
    for (const run of textRuns) {
      const rPr = ensurePropsFirst(run, "w:rPr");
      const a = forceTimesNewRoman(rPr, true);
      const b = forceSize(rPr, 24);
      if (a || b) runs++;
    }
    if (textRuns.length) {
      const markRPr = ensureParagraphMarkRPr(p.el);
      forceTimesNewRoman(markRPr, true);
      forceSize(markRPr, 24);
      // Raw borescope captions carry a left=-108/firstLine=+108 indent pair that only looks
      // right while the text fits on one line; at 12 pt more captions wrap, and every wrapped
      // line would hang out past the cell edge. Zeroing both keeps the first line exactly where
      // it was.
      const ind = firstChild(ensurePropsFirst(p.el, "w:pPr"), "w:ind");
      const left = Number(ind?.getAttribute("w:left") ?? ind?.getAttribute("w:start") ?? 0);
      const firstLine = Number(ind?.getAttribute("w:firstLine") ?? 0);
      if (ind && left < 0 && left + firstLine === 0) {
        for (const attr of ["w:left", "w:start", "w:firstLine"]) if (ind.getAttribute(attr) !== null) ind.removeAttribute(attr);
        if (ind.attributes.length === 0) ind.parentNode?.removeChild(ind);
      }
    }
  }
  log.add("photos", {
    note: runs
      ? `Set ${runs} photo-table text run${runs === 1 ? "" : "s"} to Times New Roman 12.`
      : "Photo-table text was already Times New Roman 12.",
  });

  let resized = 0;
  for (const table of model.photoTables) {
    if (table.isOperationalData || table.isDataPlate) continue;
    for (const inline of descendants(table.topTbl, "wp:inline")) {
      const extent = firstChild(inline, "wp:extent");
      if (!extent) continue;
      const cx = Number(extent.getAttribute("cx"));
      const cy = Number(extent.getAttribute("cy"));
      if (!cx || !cy) continue;
      const aspect = cx / cy;
      if (aspect < 1.2 || aspect > 1.45) continue; // only standard landscape borescope shots
      if (Math.abs(cy - PHOTO_HEIGHT_EMU) / PHOTO_HEIGHT_EMU <= 0.03) continue;
      const newCx = Math.round(PHOTO_HEIGHT_EMU * aspect);
      const tc = enclosing(inline, "w:tc");
      const tcW = tc ? firstChild(firstChild(tc, "w:tcPr") ?? tc, "w:tcW") : null;
      if (tcW?.getAttribute("w:type") === "dxa" && newCx > Number(tcW.getAttribute("w:w")) * EMU_PER_DXA * 1.02) continue;
      extent.setAttribute("cx", String(newCx));
      extent.setAttribute("cy", String(PHOTO_HEIGHT_EMU));
      for (const ext of descendants(inline, "a:ext")) {
        if (ext.getAttribute("cx")) {
          ext.setAttribute("cx", String(newCx));
          ext.setAttribute("cy", String(PHOTO_HEIGHT_EMU));
        }
      }
      resized++;
    }
  }
  if (resized) {
    log.add("photos", {
      note: `Resized ${resized} photo${resized === 1 ? "" : "s"} to the standard 3.2" x 4.3" size.`,
    });
  }
}

// --- Heading styles for the standard Report Composition section names ---

export function ensureHeadingStyles(stylesXml: Document | null, model: DocModel): { 1: string; 2: string } {
  const ids = { 1: model.headingStyleIds[1] ?? "Heading1", 2: model.headingStyleIds[2] ?? "Heading2" };
  if (!stylesXml) return ids;
  const root = stylesXml.documentElement;
  for (const level of [1, 2] as const) {
    if (model.headingStyleIds[level]) continue;
    // A report started from a blank Word document only has heading styles as latent styles
    // until first used -- define them here, matching ATS's own template definitions.
    const style = createW(stylesXml, "w:style", { "w:type": "paragraph", "w:styleId": ids[level] });
    style.appendChild(createW(stylesXml, "w:name", { "w:val": `heading ${level}` }));
    style.appendChild(createW(stylesXml, "w:basedOn", { "w:val": "Normal" }));
    style.appendChild(createW(stylesXml, "w:next", { "w:val": "Normal" }));
    style.appendChild(createW(stylesXml, "w:uiPriority", { "w:val": "9" }));
    style.appendChild(createW(stylesXml, "w:qFormat"));
    const pPr = createW(stylesXml, "w:pPr");
    pPr.appendChild(createW(stylesXml, "w:keepNext"));
    pPr.appendChild(createW(stylesXml, "w:jc", { "w:val": "center" }));
    pPr.appendChild(createW(stylesXml, "w:outlineLvl", { "w:val": String(level - 1) }));
    style.appendChild(pPr);
    const rPr = createW(stylesXml, "w:rPr");
    rPr.appendChild(createW(stylesXml, "w:b"));
    if (level === 1) {
      rPr.appendChild(createW(stylesXml, "w:sz", { "w:val": "28" }));
      rPr.appendChild(createW(stylesXml, "w:szCs", { "w:val": "28" }));
    }
    style.appendChild(rPr);
    root.appendChild(style);
    model.styleNames.set(ids[level], `heading ${level}`);
    model.headingStyleIds[level] = ids[level];
  }
  return ids;
}

export function setParagraphStyle(p: ParaInfo, styleId: string | null) {
  const pPr = ensurePropsFirst(p.el, "w:pPr");
  if (styleId) setOrderedChild(pPr, "w:pStyle", PPR_ORDER, { "w:val": styleId });
  else {
    const pStyle = firstChild(pPr, "w:pStyle");
    if (pStyle) pPr.removeChild(pStyle);
  }
}

function styleLabel(level: number | null): string {
  return level === 1 ? "Heading 1" : level === 2 ? "Heading 2" : "Normal";
}

/** Fixed rule set from the Report Composition standards. Returns the indices it decided, so the
 *  AI heading review doesn't second-guess them. */
function applyKnownHeadings(ctx: FormattingContext) {
  const { model, stylesXml, log } = ctx;
  const ids = ensureHeadingStyles(stylesXml, model);
  log.touch("headings");
  const seenSectionNames = new Map<string, Set<string>>();
  let currentH1: string | null = null;

  for (const p of model.paragraphs) {
    if (p.zone !== "body" && p.zone !== "llc") continue;
    if (p.table) continue;
    const norm = normalizeHeading(p.text);
    if (!norm) continue;
    let want: 1 | 2 | 0 | null = null;
    if (H1_NAMES.has(norm)) want = 1;
    else if (H2_NAMES.has(norm)) want = 2;
    else if (NORMAL_NAMES.has(norm)) want = 0;
    else if (
      currentH1 &&
      (currentH1 === "observations" || currentH1 === "photos") &&
      SECTION_NAME.test(norm) &&
      norm.length <= 60 &&
      norm.split(" ").length <= 7
    ) {
      // Only the first occurrence under each H1 is the real subheading -- ATS reports repeat
      // the section name as a plain "continued" label at the top of later pages, and those
      // must stay out of the Table of Contents.
      const seen: Set<string> = seenSectionNames.get(currentH1) ?? new Set<string>();
      if (!seen.has(norm)) want = 2;
      seen.add(norm);
      seenSectionNames.set(currentH1, seen);
    }

    if (want === 1) currentH1 = norm;
    else if (p.headingLevel === 1) currentH1 = norm;

    if (want === null) continue;
    const currentLevel = p.headingLevel ?? 0;
    if (currentLevel === want) continue;
    if (want === 0 && currentLevel > 2) continue;
    setParagraphStyle(p, want === 0 ? null : ids[want]);
    p.decided = true;
    log.add("headings", {
      location: locationLabel(p),
      before: styleLabel(p.headingLevel),
      after: styleLabel(want || null),
      note: `"${p.text.trim()}" set to ${styleLabel(want || null)}.`,
    });
  }
  refreshHeadingContext(model);
  // Mark the standard names as decided even when they were already right.
  for (const p of model.paragraphs) {
    const norm = normalizeHeading(p.text);
    if (H1_NAMES.has(norm) || H2_NAMES.has(norm) || NORMAL_NAMES.has(norm)) p.decided = true;
  }
}

// --- Unfilled template placeholders (reported, never changed) ---

const PLACEHOLDER_PATTERNS: { re: RegExp; label: string }[] = [
  { re: /choose an item\.?/i, label: '"Choose an item."' },
  { re: /click or tap here to enter (text|a date)\.?/i, label: '"Click or tap here to enter text."' },
  { re: /\b(19|20)XX\b/, label: "20XX year placeholder" },
  { re: /(^|[\s(])X{1,3}(?=[\s).,;:%"]|$)/, label: "X placeholder" },
];

function scanPlaceholders({ model, log }: FormattingContext) {
  for (const p of model.paragraphs) {
    if (p.zone === "toc") continue;
    if (PHOTO_CAPTION_LABELS.has(p.text.trim().toLowerCase())) continue;
    for (const { re, label } of PLACEHOLDER_PATTERNS) {
      if (re.test(p.text)) {
        log.add("placeholders", { location: locationLabel(p), before: p.text.trim().slice(0, 200), note: `${label} still in the report.`, level: "attention" });
        break;
      }
    }
  }
  log.touch("placeholders");
}

// --- Table of Contents: TNR 12, 1.5 line spacing, black ---

function applyTocFormatting({ model, stylesXml, log }: FormattingContext, hasToc: boolean) {
  if (!hasToc) {
    log.skip("toc", "No Table of Contents field was found in this report, so there was nothing to update.");
    return;
  }
  if (stylesXml) {
    const root = stylesXml.documentElement;
    const styles = descendants(root, "w:style");
    for (let level = 1; level <= 3; level++) {
      let style = styles.find((s) => getVal(firstChild(s, "w:name"))?.toLowerCase() === `toc ${level}`);
      if (!style) {
        style = createW(stylesXml, "w:style", { "w:type": "paragraph", "w:styleId": `TOC${level}` });
        style.appendChild(createW(stylesXml, "w:name", { "w:val": `toc ${level}` }));
        style.appendChild(createW(stylesXml, "w:basedOn", { "w:val": "Normal" }));
        style.appendChild(createW(stylesXml, "w:next", { "w:val": "Normal" }));
        style.appendChild(createW(stylesXml, "w:uiPriority", { "w:val": "39" }));
        style.appendChild(createW(stylesXml, "w:unhideWhenUsed"));
        if (level > 1) {
          const pPr = createW(stylesXml, "w:pPr");
          pPr.appendChild(createW(stylesXml, "w:ind", { "w:left": String((level - 1) * 240) }));
          style.appendChild(pPr);
        }
        root.appendChild(style);
        model.styleNames.set(`TOC${level}`, `toc ${level}`);
      }
      // autoRedefine would let a manual tweak in one entry silently restyle the whole TOC.
      const auto = firstChild(style, "w:autoRedefine");
      if (auto) style.removeChild(auto);
      const pPr = setOrderedChild(style, "w:pPr", STYLE_ORDER, {}, false);
      const spacing = setOrderedChild(pPr, "w:spacing", PPR_ORDER, {}, false);
      setAttr(spacing, "w:line", "360");
      setAttr(spacing, "w:lineRule", "auto");
      const rPr = setOrderedChild(style, "w:rPr", STYLE_ORDER, {}, false);
      forceTimesNewRoman(rPr, true);
      forceSize(rPr, 24);
      stripColor(rPr);
    }
  }
  // Existing entries too, in case Word isn't available to rebuild them.
  for (const p of model.paragraphs) {
    if (p.zone !== "toc" || !p.text.trim()) continue;
    const pPr = ensurePropsFirst(p.el, "w:pPr");
    const spacing = setOrderedChild(pPr, "w:spacing", PPR_ORDER, {}, false);
    setAttr(spacing, "w:line", "360");
    setAttr(spacing, "w:lineRule", "auto");
    for (const rPr of descendants(p.el, "w:rPr")) {
      forceSize(rPr, 24);
      forceTimesNewRoman(rPr, false);
      setOrderedChild(rPr, "w:color", RPR_ORDER, { "w:val": "000000" });
    }
  }
  log.add("toc", { note: "Table of Contents styles set to Times New Roman 12, 1.5 line spacing, black." });
}
