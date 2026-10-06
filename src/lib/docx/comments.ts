/**
 * Adds Word review comments (word/comments.xml) anchored to whole paragraphs. Used for AI
 * findings that aren't text edits -- a photo/observation mismatch, an Overall Assessment
 * recommendation -- so they show up in Word's review pane right next to the text they're about,
 * instead of being typed into the customer-facing report itself.
 */
import type JSZip from "jszip";
import { createW, descendants, firstChild, parseXml, serializeXml, W_NS, XML_NS } from "./ooxml";

const COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
export const COMMENT_AUTHOR = "ATS Report Reviewer (AI)";

export class CommentWriter {
  private commentsDoc: Document | null = null;
  private nextId = 0;
  count = 0;

  constructor(private zip: JSZip) {}

  async init() {
    const existing = this.zip.file("word/comments.xml");
    if (existing) {
      this.commentsDoc = parseXml(await existing.async("text"));
      for (const c of descendants(this.commentsDoc.documentElement, "w:comment")) {
        this.nextId = Math.max(this.nextId, Number(c.getAttribute("w:id")) + 1);
      }
    } else {
      this.commentsDoc = parseXml(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W_NS}"></w:comments>`,
      );
    }
  }

  /** Anchors a comment spanning the whole of paragraph `p`. */
  add(p: Element, text: string) {
    if (!this.commentsDoc) throw new Error("CommentWriter.init() not called");
    const id = String(this.nextId++);
    const doc = p.ownerDocument;

    const start = createW(doc, "w:commentRangeStart", { "w:id": id });
    const pPr = firstChild(p, "w:pPr");
    p.insertBefore(start, pPr ? pPr.nextSibling : p.firstChild);
    p.appendChild(createW(doc, "w:commentRangeEnd", { "w:id": id }));
    const refRun = createW(doc, "w:r");
    refRun.appendChild(createW(doc, "w:commentReference", { "w:id": id }));
    p.appendChild(refRun);

    const cdoc = this.commentsDoc;
    const comment = createW(cdoc, "w:comment", {
      "w:id": id,
      "w:author": COMMENT_AUTHOR,
      "w:date": new Date().toISOString().replace(/\.\d+Z$/, "Z"),
      "w:initials": "AI",
    });
    for (const line of text.split(/\n+/)) {
      const cp = createW(cdoc, "w:p");
      const r = createW(cdoc, "w:r");
      const t = createW(cdoc, "w:t");
      t.setAttributeNS(XML_NS, "xml:space", "preserve");
      t.appendChild(cdoc.createTextNode(line));
      r.appendChild(t);
      cp.appendChild(r);
      comment.appendChild(cp);
    }
    cdoc.documentElement.appendChild(comment);
    this.count++;
  }

  /** Writes comments.xml plus its relationship and content-type registration. */
  async save() {
    if (!this.commentsDoc || this.count === 0) return;
    this.zip.file("word/comments.xml", serializeXml(this.commentsDoc));

    const relsPath = "word/_rels/document.xml.rels";
    let rels = (await this.zip.file(relsPath)?.async("text")) ?? "";
    if (!rels.includes(COMMENTS_REL_TYPE)) {
      const ids = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
      const relId = `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
      rels = rels.replace(
        "</Relationships>",
        `<Relationship Id="${relId}" Type="${COMMENTS_REL_TYPE}" Target="comments.xml"/></Relationships>`,
      );
      this.zip.file(relsPath, rels);
    }

    let types = (await this.zip.file("[Content_Types].xml")?.async("text")) ?? "";
    if (!types.includes('PartName="/word/comments.xml"')) {
      types = types.replace(
        "</Types>",
        `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
      );
      this.zip.file("[Content_Types].xml", types);
    }
  }
}
