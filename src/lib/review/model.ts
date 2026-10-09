/**
 * A flat, ordered model of every paragraph in an inspector's report body, with enough context
 * (which part of the report it's in, which heading it sits under, whether it's in a photo or
 * Observations table) for both the deterministic formatting rules and the AI passes to target
 * the right text. Built once per review, against the parsed document.xml DOM -- `el` points at
 * the live w:p element, so later edits land directly in the document.
 *
 * Built for one-off reports with no template behind them, so nothing here assumes the 7EA
 * template's exact layout: a "photo table" is any top-level table holding a picture, an
 * "observation table" is any non-photo table under the Observations heading, and heading
 * context comes from the paragraphs' own heading styles.
 */
import { childrenThroughWrappers, descendants, elementChildren, firstChild, getVal, textOf } from "../docx/ooxml";

export type Zone = "cover" | "toc" | "body" | "llc";
export type TableKind = "photo" | "observation" | "other";

export interface TableCtx {
  /** The outermost table this paragraph sits in. */
  topTbl: Element;
  /** The innermost table (a photo block's caption table is nested inside the photo table). */
  tbl: Element;
  row: Element;
  rowIndex: number;
  cellIndex: number;
  nested: boolean;
  kind: TableKind;
  /** For photo-table caption rows: the label in the row's first cell ("component", "observation", ...). */
  rowLabel: string | null;
}

export interface ParaInfo {
  index: number;
  el: Element;
  styleId: string | null;
  /** 1 or 2 when styled Heading 1/2 (resolved through styles.xml names, not assumed IDs). */
  headingLevel: number | null;
  text: string;
  zone: Zone;
  h1: string | null;
  h2: string | null;
  table: TableCtx | null;
  /** Set once a fixed Report Composition rule has decided this paragraph's style, so the AI
   *  heading review leaves it alone. */
  decided?: boolean;
}

export interface PhotoTable {
  topTbl: Element;
  text: string;
  isOperationalData: boolean;
  isDataPlate: boolean;
  h2: string | null;
}

export interface DocModel {
  paragraphs: ParaInfo[];
  photoTables: PhotoTable[];
  /** styleId -> lowercased style name ("heading 1", "toc 2", "normal", ...). */
  styleNames: Map<string, string>;
  headingStyleIds: { 1: string | null; 2: string | null };
}

export function normalizeHeading(text: string): string {
  return text
    .toLowerCase()
    .replace(/[‘’]/g, "'")
    .replace(/&/g, "and")
    .replace(/\s+/g, " ")
    .replace(/[:.\s]+$/, "")
    .trim();
}

export const PHOTO_CAPTION_LABELS = new Set([
  "component",
  "location",
  "observation",
  "classification",
  "comments",
  "comment",
  "section",
  "measurement type",
  "measurement values (inches)",
  "measurement values",
  "serial number",
]);

export function parseStyleNames(stylesXml: Document | null): Map<string, string> {
  const map = new Map<string, string>();
  if (!stylesXml) return map;
  for (const style of descendants(stylesXml.documentElement, "w:style")) {
    const id = style.getAttribute("w:styleId");
    const name = getVal(firstChild(style, "w:name"));
    if (id) map.set(id, (name ?? id).toLowerCase());
  }
  return map;
}

function headingLevelOf(styleId: string | null, styleNames: Map<string, string>): number | null {
  if (!styleId) return null;
  const name = styleNames.get(styleId) ?? styleId.toLowerCase();
  const m = /^heading\s*(\d)$/.exec(name);
  return m ? Number(m[1]) : null;
}

function isTocStyle(styleId: string | null, styleNames: Map<string, string>): boolean {
  if (!styleId) return false;
  const name = styleNames.get(styleId) ?? styleId.toLowerCase();
  return /^toc\s*\d$/.test(name) || name === "toc heading";
}

export function paragraphStyleId(p: Element): string | null {
  const pPr = firstChild(p, "w:pPr");
  return pPr ? getVal(firstChild(pPr, "w:pStyle")) : null;
}

function hasPicture(el: Element): boolean {
  return descendants(el, "w:drawing").length > 0 || descendants(el, "w:pict").length > 0;
}

export function buildModel(documentXml: Document, stylesXml: Document | null): DocModel {
  const styleNames = parseStyleNames(stylesXml);
  const headingStyleIds: { 1: string | null; 2: string | null } = { 1: null, 2: null };
  for (const [id, name] of styleNames) {
    if (name === "heading 1" && !headingStyleIds[1]) headingStyleIds[1] = id;
    if (name === "heading 2" && !headingStyleIds[2]) headingStyleIds[2] = id;
  }

  const body = descendants(documentXml.documentElement, "w:body")[0];
  const paragraphs: ParaInfo[] = [];
  const photoTables: PhotoTable[] = [];

  // First pass: flatten every paragraph (body + table cells) in document order.
  type Raw = { el: Element; table: Omit<TableCtx, "kind"> | null };
  const raws: Raw[] = [];
  const topTables: Element[] = [];

  const walkBlock = (container: Element, tableStack: Omit<TableCtx, "kind"> | null) => {
    for (const child of elementChildren(container)) {
      if (child.tagName === "w:p") raws.push({ el: child, table: tableStack });
      else if (child.tagName === "w:tbl") walkTable(child, tableStack);
      else if (child.tagName === "w:sdt") {
        const content = firstChild(child, "w:sdtContent");
        if (content) walkBlock(content, tableStack);
      } else if (child.tagName === "w:customXml") walkBlock(child, tableStack);
    }
  };

  const walkTable = (tbl: Element, outer: Omit<TableCtx, "kind"> | null) => {
    const topTbl = outer ? outer.topTbl : tbl;
    if (!outer) topTables.push(tbl);
    childrenThroughWrappers(tbl, "w:tr").forEach((row, rowIndex) => {
      const cells = childrenThroughWrappers(row, "w:tc");
      const rowLabel = cells.length >= 2 ? textOf(cells[0]).trim().toLowerCase().replace(/:$/, "") : null;
      cells.forEach((cell, cellIndex) => {
        walkBlock(cell, { topTbl, tbl, row, rowIndex, cellIndex, nested: !!outer, rowLabel });
      });
    });
  };

  walkBlock(body, null);

  const photoTopTables = new Set(topTables.filter(hasPicture));

  // Locate the cover/TOC boundary: the TOC title or first TOC-styled paragraph. With no TOC at
  // all, the cover runs until the first Heading 1.
  let tocStart = raws.findIndex(
    (r) => isTocStyle(paragraphStyleId(r.el), styleNames) || normalizeHeading(textOf(r.el)) === "table of contents",
  );
  const hasToc = tocStart !== -1;
  if (tocStart === -1) {
    tocStart = raws.findIndex((r) => headingLevelOf(paragraphStyleId(r.el), styleNames) === 1);
    if (tocStart === -1) tocStart = 0;
  }

  let h1: string | null = null;
  let h2: string | null = null;
  let inLlc = false;
  let tocEnded = false;

  raws.forEach((raw, index) => {
    const styleId = paragraphStyleId(raw.el);
    const level = headingLevelOf(styleId, styleNames);
    const text = textOf(raw.el);
    const norm = normalizeHeading(text);

    // The TOC zone runs from its title through its entries (and any blank spacer paragraphs
    // between them); the first real non-TOC paragraph after that starts the body for good.
    let zone: Zone;
    if (index < tocStart) zone = "cover";
    else if (!tocEnded && hasToc) {
      const tocLike =
        index === tocStart || isTocStyle(styleId, styleNames) || norm === "table of contents" || (!text.trim() && !raw.table);
      if (tocLike) zone = "toc";
      else {
        tocEnded = true;
        zone = "body";
      }
    } else zone = "body";

    if (zone === "body") {
      if (level === 1) {
        h1 = text.trim();
        h2 = null;
        inLlc = norm === "llc statement" || norm === "limitation of liability";
      } else if (level === 2) {
        h2 = text.trim();
        inLlc = false;
      } else if (!raw.table && (norm === "limitation of liability" || norm === "llc statement")) {
        inLlc = true;
      }
      if (inLlc) zone = "llc";
    }

    let table: TableCtx | null = null;
    if (raw.table) {
      // The cover page is itself a table holding the ATS logo -- it must not be treated as a
      // photo table, or its 40/22 pt title block gets flattened to the photo-caption 12 pt.
      const kind: TableKind = photoTopTables.has(raw.table.topTbl) && zone !== "cover" && zone !== "toc"
        ? "photo"
        : h1 && normalizeHeading(h1) === "observations"
          ? "observation"
          : "other";
      table = { ...raw.table, kind };
    }

    paragraphs.push({ index, el: raw.el, styleId, headingLevel: level, text, zone, h1, h2, table });
  });

  // Photo tables, with the operational-data / data-plate flags the image pass relies on.
  const seen = new Set<Element>();
  for (const p of paragraphs) {
    if (!p.table || p.table.kind !== "photo" || seen.has(p.table.topTbl)) continue;
    seen.add(p.table.topTbl);
    const text = textOf(p.table.topTbl);
    photoTables.push({
      topTbl: p.table.topTbl,
      text,
      isOperationalData: /operational\s+data/i.test(text),
      isDataPlate: /data\s*plate/i.test(text),
      h2: p.h2,
    });
  }

  return { paragraphs, photoTables, styleNames, headingStyleIds };
}

/** Rebuilds the h1/h2 heading context after heading styles have been corrected. */
export function refreshHeadingContext(model: DocModel) {
  let h1: string | null = null;
  let h2: string | null = null;
  for (const p of model.paragraphs) {
    p.styleId = paragraphStyleId(p.el);
    p.headingLevel = headingLevelOf(p.styleId, model.styleNames);
    if (p.zone === "body" || p.zone === "llc") {
      if (p.headingLevel === 1) {
        h1 = p.text.trim();
        h2 = null;
      } else if (p.headingLevel === 2) h2 = p.text.trim();
    }
    p.h1 = h1;
    p.h2 = h2;
  }
}

export function locationLabel(p: ParaInfo): string {
  const parts: string[] = [];
  if (p.zone === "cover") parts.push("Cover page");
  else if (p.zone === "toc") parts.push("Table of Contents");
  else {
    if (p.h1) parts.push(p.h1);
    if (p.h2) parts.push(p.h2);
  }
  if (p.table) {
    if (p.table.kind === "photo") parts.push(p.table.rowLabel ? `photo table (${p.table.rowLabel})` : "photo table");
    else parts.push(`table row ${p.table.rowIndex + 1}`);
  }
  return parts.join(" › ") || "Report body";
}
