/**
 * Photo section page layout (Brett, 2026-10-12):
 *  - Every page of the Photos section uses Word's "Narrow" margins (0.5" all round). The photo
 *    tables are ~7.4" wide, so under 1" margins their caption column runs off the printable
 *    area. Inspectors' reports often carry a stray section break partway through the photos
 *    (Stony Brook GT1C: everything from "Combustion Section" on had 1" margins).
 *  - If the section break before Photos was deleted, put one back, so narrowing the photo pages
 *    doesn't also narrow the report body before them.
 *  - Remove photo-table cells sitting entirely in the margin. Word hides them on screen, but they
 *    print in the PDF (e.g. a stray "Basket" to the right of a caption table).
 */
import { childrenThroughWrappers, createW, descendants, elementChildren, firstChild, getVal, setAttr, textOf } from "../docx/ooxml";
import { DocModel, normalizeHeading } from "./model";
import type { ReviewLog } from "./report";

const NARROW_DXA = 720; // 0.5"
/** ATS's standard report-body margins (Batavia, Stony Brook): 0.75" top, 0.625" bottom, 1" sides. */
const BODY_MARGINS: Record<string, string> = {
  "w:top": "1080",
  "w:bottom": "900",
  "w:left": "1440",
  "w:right": "1440",
  "w:header": "540",
  "w:footer": "720",
};

const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export interface HeaderFooterRel {
  id: string;
  kind: "header" | "footer";
}

export function parseHeaderFooterRels(relsXml: string): HeaderFooterRel[] {
  const out: HeaderFooterRel[] = [];
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const kind = /\bType="[^"]*\/relationships\/(header|footer)"/.exec(m[0])?.[1] as "header" | "footer" | undefined;
    if (id && kind) out.push({ id, kind });
  }
  return out;
}

/** Deleting a section break also deletes the header/footer assignment it carried; the header
 *  and footer parts stay in the file, just attached to nothing. Hand those back to the restored
 *  section so the body (and the photo pages that inherit from it) get their header back. */
function reattachOrphanedHeaders(body: Element, sectPr: Element, rels: HeaderFooterRel[]): number {
  const referenced = new Set(
    [...descendants(body, "w:headerReference"), ...descendants(body, "w:footerReference")].map((r) => r.getAttribute("r:id")),
  );
  let attached = 0;
  for (const kind of ["footer", "header"] as const) {
    const tag = kind === "header" ? "w:headerReference" : "w:footerReference";
    const hasDefault = elementChildren(sectPr).some((c) => c.tagName === tag && c.getAttribute("w:type") === "default");
    const orphan = rels.find((r) => r.kind === kind && !referenced.has(r.id));
    if (hasDefault || !orphan) continue;
    const ref = createW(sectPr.ownerDocument, tag, { "w:type": "default" });
    ref.setAttributeNS(R_NS, "r:id", orphan.id);
    // Header references come first in a sectPr, then footer references, then everything else.
    const before =
      kind === "header"
        ? elementChildren(sectPr).find((c) => c.tagName !== "w:headerReference")
        : elementChildren(sectPr).find((c) => c.tagName !== "w:headerReference" && c.tagName !== "w:footerReference");
    sectPr.insertBefore(ref, before ?? null);
    attached++;
  }
  return attached;
}

interface Section {
  /** Index of the section's first body-level element. */
  start: number;
  /** Index of the element holding the sectPr (inclusive end). */
  end: number;
  sectPr: Element;
}

function sectionsOf(body: Element): { children: Element[]; sections: Section[] } {
  const children = elementChildren(body);
  const sections: Section[] = [];
  let start = 0;
  children.forEach((child, i) => {
    let sectPr: Element | null = null;
    if (child.tagName === "w:sectPr") sectPr = child;
    else if (child.tagName === "w:p") {
      const pPr = firstChild(child, "w:pPr");
      sectPr = pPr ? firstChild(pPr, "w:sectPr") : null;
    }
    if (sectPr) {
      sections.push({ start, end: i, sectPr });
      start = i + 1;
    }
  });
  return { children, sections };
}

function topLevel(el: Element, body: Element): Element | null {
  let node: Element | null = el;
  while (node && node.parentNode !== body) node = node.parentNode as Element | null;
  return node;
}

/** A paragraph holding nothing but a page break (or nothing at all). */
function isBreakOnly(el: Element): boolean {
  if (el.tagName !== "w:p" || textOf(el).trim()) return false;
  if (descendants(el, "w:drawing").length || descendants(el, "w:pict").length) return false;
  return true;
}

function hasPageBreak(el: Element): boolean {
  return descendants(el, "w:br").some((br) => br.getAttribute("w:type") === "page");
}

export function applyPhotoSectionLayout(documentXml: Document, model: DocModel, log: ReviewLog, headerFooterRels: HeaderFooterRel[]) {
  const body = descendants(documentXml.documentElement, "w:body")[0];
  const heading = model.paragraphs.find((p) => p.headingLevel === 1 && normalizeHeading(p.text) === "photos" && !p.table);
  const headingTop = heading ? topLevel(heading.el, body) : null;
  if (!heading || !headingTop) {
    log.add("photos", { note: "No \"Photos\" heading was found, so the photo section's margins were left as they are.", level: "attention" });
    return;
  }

  // 1. Restore a section break before Photos if the report body runs straight into it.
  let { children, sections } = sectionsOf(body);
  const headingIndex = children.indexOf(headingTop);
  const containing = sections.find((s) => headingIndex >= s.start && headingIndex <= s.end);
  if (containing) {
    const before = children.slice(containing.start, headingIndex);
    const contentBefore = before.some((el) => !isBreakOnly(el));
    if (contentBefore) {
      const earlier = containing.sectPr.cloneNode(true) as Element;
      // With the break deleted, the body inherited the photo section's settings; if those are
      // Narrow, the body's own margins are gone from the file, so restore ATS's standard ones
      // (every template-built report uses these for the body).
      const earlierMar = firstChild(earlier, "w:pgMar");
      if (earlierMar && Number(earlierMar.getAttribute("w:left") || 0) <= NARROW_DXA) {
        for (const [attr, val] of Object.entries(BODY_MARGINS)) setAttr(earlierMar, attr, val);
      }
      // The photos keep the original sectPr: carry on numbering pages, no "different first page".
      const photosSect = containing.sectPr;
      const pgNum = firstChild(photosSect, "w:pgNumType");
      if (pgNum?.getAttribute("w:start")) pgNum.removeAttribute("w:start");
      for (const tag of ["w:titlePg", "w:type"]) {
        const el = firstChild(photosSect, tag);
        if (el) photosSect.removeChild(el); // no w:type = starts on a new page
      }
      const breakPara = createW(documentXml, "w:p");
      const pPr = createW(documentXml, "w:pPr");
      pPr.appendChild(earlier);
      breakPara.appendChild(pPr);
      // The section break now provides the new page, so a page-break-only paragraph right before
      // the heading (or a page break inside it) would leave a blank page.
      const prev = before[before.length - 1];
      if (prev && isBreakOnly(prev) && hasPageBreak(prev)) body.replaceChild(breakPara, prev);
      else body.insertBefore(breakPara, headingTop);
      for (const br of descendants(heading.el, "w:br")) {
        if (br.getAttribute("w:type") === "page") br.parentNode?.removeChild(br);
      }
      const reattached = reattachOrphanedHeaders(body, earlier, headerFooterRels);
      log.add("photos", {
        note: `Added the missing section break before the Photos section, so its margins stay separate from the rest of the report${
          reattached ? " (and reattached the page header/footer the deleted break had carried)" : ""
        }.`,
      });
      ({ children, sections } = sectionsOf(body));
    }
  }

  // 2. Narrow margins on every section that holds the photos.
  const photoTops = new Set<Element>();
  for (const p of model.paragraphs) {
    if (p.h1 && normalizeHeading(p.h1) === "photos") {
      const top = topLevel(p.el, body);
      if (top) photoTops.add(top);
    }
  }
  const photoSections = sections.filter((s) => children.slice(s.start, s.end + 1).some((el) => photoTops.has(el)));
  let narrowed = 0;
  for (const s of photoSections) {
    let pgMar = firstChild(s.sectPr, "w:pgMar");
    if (!pgMar) {
      pgMar = createW(documentXml, "w:pgMar", { "w:header": "720", "w:footer": "720", "w:gutter": "0" });
      const pgSz = firstChild(s.sectPr, "w:pgSz");
      s.sectPr.insertBefore(pgMar, pgSz ? pgSz.nextSibling : s.sectPr.firstChild);
    }
    let changed = false;
    for (const side of ["w:top", "w:bottom", "w:left", "w:right"]) {
      if (pgMar.getAttribute(side) !== String(NARROW_DXA)) {
        setAttr(pgMar, side, String(NARROW_DXA));
        changed = true;
      }
    }
    if (changed) narrowed++;
  }
  log.add("photos", {
    note: narrowed
      ? `Set the Photos section's margins to Narrow (0.5") on ${narrowed} section${narrowed === 1 ? "" : "s"} that had other margins.`
      : "Photos section margins were already Narrow (0.5\").",
  });

  // 3. Remove photo-table cells that sit entirely in the right margin.
  let removed = 0;
  const removedText: string[] = [];
  for (const s of photoSections) {
    const pgSz = firstChild(s.sectPr, "w:pgSz");
    const pgMar = firstChild(s.sectPr, "w:pgMar");
    const usable = Number(pgSz?.getAttribute("w:w") ?? 12240) - Number(pgMar?.getAttribute("w:left") ?? 0) - Number(pgMar?.getAttribute("w:right") ?? 0);
    for (const el of children.slice(s.start, s.end + 1)) {
      if (el.tagName !== "w:tbl") continue;
      if (!descendants(el, "w:drawing").length && !descendants(el, "w:pict").length) continue;
      removed += removeCellsPastMargin(el, 0, usable, removedText);
    }
  }
  if (removed) {
    log.add("photos", {
      note: `Removed ${removed} photo-table cell${removed === 1 ? "" : "s"} hidden in the page margin (invisible in Word but printed in the PDF): ${removedText
        .slice(0, 6)
        .map((t) => `"${t}"`)
        .join(", ")}${removedText.length > 6 ? ", ..." : ""}.`,
    });
  }
}

function dxa(el: Element | null, attr = "w:w"): number {
  if (!el) return 0;
  const type = el.getAttribute("w:type");
  if (type && type !== "dxa") return 0;
  return Number(el.getAttribute(attr) || 0);
}

/** Walks a table's cells left to right using its grid, removing any cell (with content) that sits
 *  in the right margin. Recurses into nested caption tables. */
function removeCellsPastMargin(tbl: Element, x0: number, usable: number, removedText: string[]): number {
  const grid = firstChild(tbl, "w:tblGrid");
  const cols = grid ? elementChildren(grid).filter((c) => c.tagName === "w:gridCol").map((c) => Number(c.getAttribute("w:w") || 0)) : [];
  const tblPr = firstChild(tbl, "w:tblPr");
  const indent = dxa(tblPr ? firstChild(tblPr, "w:tblInd") : null);
  let removed = 0;

  for (const row of childrenThroughWrappers(tbl, "w:tr")) {
    const trPr = firstChild(row, "w:trPr");
    let col = Number(getVal(trPr ? firstChild(trPr, "w:gridBefore") : null) ?? 0);
    let x = x0 + indent + cols.slice(0, col).reduce((a, b) => a + b, 0);
    const cells = childrenThroughWrappers(row, "w:tc");
    cells.forEach((tc, i) => {
      const tcPr = firstChild(tc, "w:tcPr");
      const span = Number(getVal(tcPr ? firstChild(tcPr, "w:gridSpan") : null) ?? 1);
      const width = cols.length ? cols.slice(col, col + span).reduce((a, b) => a + b, 0) : dxa(tcPr ? firstChild(tcPr, "w:tcW") : null);
      const hasContent = !!textOf(tc).trim() || descendants(tc, "w:drawing").length > 0;
      // "In the margin": less than ~0.14" of the cell is inside the printable width, or most of
      // it is outside. Regular caption columns end inside the margin and are never touched.
      const visible = usable - x;
      const inMargin = visible < 200 || (width > 0 && visible < width / 2);
      if (i > 0 && inMargin && hasContent) {
        removedText.push(textOf(tc).trim().slice(0, 40));
        tc.parentNode?.removeChild(tc);
        removed++;
      } else {
        for (const inner of elementChildren(tc).filter((c) => c.tagName === "w:tbl")) {
          removed += removeCellsPastMargin(inner, x, usable, removedText);
        }
      }
      x += width;
      col += span;
    });
  }
  return removed;
}
