/**
 * DOM-level OOXML helpers on top of @xmldom/xmldom.
 *
 * Unlike the Borescope app (which splices raw document.xml text so it never has to
 * re-serialize a 100+ page report), the reviewer touches nearly every run in the document --
 * fonts, sizes, colors, redlined text -- so it works on a parsed DOM and serializes once at the
 * end. Word is strict about child-element ORDER inside property blocks (rPr/pPr/tcPr/...) even
 * where the XML is otherwise valid, so every property write goes through setOrderedChild() with
 * the schema's own sequence rather than a plain appendChild().
 */
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const XML_NS = "http://www.w3.org/XML/1998/namespace";

export function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml") as unknown as Document;
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc as unknown as Parameters<XMLSerializer["serializeToString"]>[0]);
}

export function elementChildren(el: Element): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1) out.push(n as Element);
  }
  return out;
}

export function firstChild(el: Element, tagName: string): Element | null {
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (n as Element).tagName === tagName) return n as Element;
  }
  return null;
}

/** Wrapper elements that can sit between a table/row/body and its "real" children (block-level
 *  content controls, custom XML) -- walked through transparently. */
const TRANSPARENT_WRAPPERS = new Set(["w:sdt", "w:sdtContent", "w:customXml", "w:smartTag"]);

/** Direct children named `tagName`, looking through w:sdt/w:sdtContent/w:customXml wrappers --
 *  e.g. the w:tr rows of a table even when some rows are wrapped in a row-level content control. */
export function childrenThroughWrappers(el: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (const child of elementChildren(el)) {
    if (child.tagName === tagName) out.push(child);
    else if (TRANSPARENT_WRAPPERS.has(child.tagName)) {
      const inner = child.tagName === "w:sdt" ? firstChild(child, "w:sdtContent") : child;
      if (inner) out.push(...childrenThroughWrappers(inner, tagName));
    }
  }
  return out;
}

export function descendants(el: Element, tagName: string): Element[] {
  const list = el.getElementsByTagName(tagName);
  const out: Element[] = [];
  for (let i = 0; i < list.length; i++) out.push(list[i]);
  return out;
}

/** Visible text of a paragraph/cell/table: every w:t plus tabs, ignoring deleted text. */
export function textOf(el: Element): string {
  let out = "";
  const walk = (node: Element) => {
    for (const child of elementChildren(node)) {
      if (child.tagName === "w:t") out += child.textContent ?? "";
      else if (child.tagName === "w:tab" && node.tagName === "w:r") out += "\t";
      else if (child.tagName === "w:del" || child.tagName === "w:delText" || child.tagName === "w:instrText") continue;
      else walk(child);
    }
  };
  walk(el);
  return out;
}

export function getVal(el: Element | null, attr = "w:val"): string | null {
  if (!el) return null;
  const v = el.getAttribute(attr);
  return v === "" ? null : v;
}

export function setAttr(el: Element, name: string, value: string) {
  el.setAttributeNS(W_NS, name, value);
}

export function createW(doc: Document, tagName: string, attrs: Record<string, string> = {}): Element {
  const el = doc.createElementNS(W_NS, tagName);
  for (const [k, v] of Object.entries(attrs)) setAttr(el, k, v);
  return el;
}

// Schema child orders (ECMA-376 Part 1). Only the elements this app ever writes need to be
// correct relative to each other, but the full lists keep insertion right next to anything an
// inspector's own Word wrote there.
export const RPR_ORDER = [
  "w:ins", "w:del", "w:moveFrom", "w:moveTo",
  "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike",
  "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden",
  "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
  "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
  "w:specVanish", "w:oMath", "w:rPrChange",
];
export const PPR_ORDER = [
  "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", "w:numPr",
  "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
  "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
  "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
  "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
  "w:sectPr", "w:pPrChange",
];
export const TCPR_ORDER = [
  "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar",
  "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
  "w:cellMerge", "w:tcPrChange",
];
export const TBLPR_ORDER = [
  "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize", "w:tblStyleColBandSize",
  "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
  "w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange",
];
export const STYLE_ORDER = [
  "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden", "w:uiPriority",
  "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", "w:personal", "w:personalCompose",
  "w:personalReply", "w:rsid", "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
];

/** Returns the child `tagName` of `parent`, creating it (at its schema-correct position) if
 *  missing, then applies `attrs` (replacing every existing attribute when `replaceAttrs`). */
export function setOrderedChild(
  parent: Element,
  tagName: string,
  order: string[],
  attrs: Record<string, string> = {},
  replaceAttrs = true,
): Element {
  let el = firstChild(parent, tagName);
  if (!el) {
    el = createW(parent.ownerDocument, tagName);
    const myIndex = order.indexOf(tagName);
    let before: Element | null = null;
    for (const child of elementChildren(parent)) {
      const idx = order.indexOf(child.tagName);
      if (idx > myIndex) {
        before = child;
        break;
      }
    }
    parent.insertBefore(el, before);
  } else if (replaceAttrs) {
    while (el.attributes.length > 0) el.removeAttributeNode(el.attributes[0]);
  }
  for (const [k, v] of Object.entries(attrs)) setAttr(el, k, v);
  return el;
}

export function removeChildren(parent: Element, tagName: string): number {
  let removed = 0;
  for (const child of elementChildren(parent)) {
    if (child.tagName === tagName) {
      parent.removeChild(child);
      removed++;
    }
  }
  return removed;
}

/** The property block (w:rPr in a run, w:pPr in a paragraph, ...) as the element's first child,
 *  created if missing. */
export function ensurePropsFirst(el: Element, propsTag: string): Element {
  const existing = firstChild(el, propsTag);
  if (existing) return existing;
  const props = createW(el.ownerDocument, propsTag);
  el.insertBefore(props, el.firstChild);
  return props;
}

export function ensureParagraphMarkRPr(p: Element): Element {
  const pPr = ensurePropsFirst(p, "w:pPr");
  return setOrderedChild(pPr, "w:rPr", PPR_ORDER, {}, false);
}

const SYMBOL_FONTS = /^(symbol|wingdings.*|webdings|mt extra|marlett|segoe ui symbol|zapf ?dingbats)$/i;
export const REPORT_FONT = "Times New Roman";

/** Points rFonts at Times New Roman unless it's a symbol font (bullets/checkmarks would turn
 *  into letters). Returns true if anything changed. */
export function forceTimesNewRoman(rPr: Element, addIfMissing = true): boolean {
  const rFonts = firstChild(rPr, "w:rFonts");
  if (!rFonts && !addIfMissing) return false;
  if (rFonts) {
    const ascii = rFonts.getAttribute("w:ascii") || rFonts.getAttribute("w:hAnsi");
    if (ascii && SYMBOL_FONTS.test(ascii)) return false;
    const inherits = !ascii && !rFonts.getAttribute("w:asciiTheme") && !rFonts.getAttribute("w:hAnsiTheme");
    if (inherits && !addIfMissing) return false;
    const already =
      rFonts.getAttribute("w:ascii") === REPORT_FONT &&
      rFonts.getAttribute("w:hAnsi") === REPORT_FONT &&
      !rFonts.getAttribute("w:asciiTheme") &&
      !rFonts.getAttribute("w:hAnsiTheme");
    if (already) return false;
  }
  setOrderedChild(rPr, "w:rFonts", RPR_ORDER, {
    "w:ascii": REPORT_FONT,
    "w:eastAsia": REPORT_FONT,
    "w:hAnsi": REPORT_FONT,
    "w:cs": REPORT_FONT,
  });
  return true;
}

/** Sets an explicit font size (half-points) on an rPr. Returns true if it changed. */
export function forceSize(rPr: Element, halfPoints: number): boolean {
  const v = String(halfPoints);
  const sz = firstChild(rPr, "w:sz");
  const szCs = firstChild(rPr, "w:szCs");
  if (getVal(sz) === v && getVal(szCs) === v) return false;
  setOrderedChild(rPr, "w:sz", RPR_ORDER, { "w:val": v });
  setOrderedChild(rPr, "w:szCs", RPR_ORDER, { "w:val": v });
  return true;
}

/** True for a color light enough that it's only ever used as text on a dark fill (white table
 *  header text, etc.) -- turning that black would make it unreadable, so it's left alone. */
export function isLightColor(hex: string | null): boolean {
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return false;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255 > 0.85;
}

/** Removes a text color (making it inherit black) unless it's a light-on-dark color. Returns true
 *  if something was removed. */
export function stripColor(rPr: Element): boolean {
  const color = firstChild(rPr, "w:color");
  if (!color) return false;
  const val = color.getAttribute("w:val");
  if (!val || val === "auto" || val === "000000") {
    if (color.getAttribute("w:themeColor")) {
      rPr.removeChild(color);
      return true;
    }
    return false;
  }
  if (isLightColor(val)) return false;
  rPr.removeChild(color);
  return true;
}
