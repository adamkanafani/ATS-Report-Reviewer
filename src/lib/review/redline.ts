/**
 * Applies an AI-corrected version of a paragraph's text back onto the real w:p, coloring only
 * the words that changed red -- Brett reviews the output by scanning for red text, so a fixed
 * typo should light up one word, not the whole sentence.
 *
 * A paragraph is broken into "atoms":
 *  - one char atom per character of plain run text (remembering which run it came from, so
 *    unchanged text keeps its exact original formatting -- bold labels, italics, etc.);
 *  - opaque atoms for anything that isn't plain text: fields (dates, page numbers), hyperlinks,
 *    content controls, pictures, bookmarks, tracked changes. Opaque atoms that show text are
 *    handed to the AI as ⟦n⟧ tokens it must keep, so they're never retyped as plain text; the
 *    rest are zero-width and just ride along at their original position.
 * The old and new text are diffed word-by-word (LCS), and the paragraph is rebuilt from the
 * resulting atom sequence: equal words keep their original atoms, inserted words become new runs
 * cloned from the neighboring run's formatting plus red color, deleted words are dropped.
 */
import { createW, elementChildren, firstChild, RPR_ORDER, setOrderedChild, XML_NS } from "../docx/ooxml";

export const AI_RED = "FF0000";

type CharAtom = { kind: "char"; ch: string; run: Element; red: boolean };
type OpaqueAtom = { kind: "opaque"; nodes: Element[]; token: string | null };
type Atom = CharAtom | OpaqueAtom;

export interface ParagraphText {
  /** Text as shown to the AI, with ⟦n⟧ tokens standing in for fields/links/controls. */
  text: string;
  atoms: Atom[];
  editable: boolean;
}

const SIMPLE_RUN_CHILDREN = new Set(["w:rPr", "w:t", "w:tab", "w:br", "w:lastRenderedPageBreak", "w:softHyphen"]);
/** Zero-width markers that can simply be dropped when the paragraph is rebuilt. */
const DROPPABLE = new Set(["w:proofErr", "w:lastRenderedPageBreak"]);

function hasVisibleText(nodes: Element[]): boolean {
  for (const n of nodes) {
    if (n.getElementsByTagName("w:t").length > 0) return true;
    if (n.getElementsByTagName("w:sym").length > 0) return true;
    if (n.tagName === "w:fldSimple") return true;
  }
  return false;
}

function isSimpleRun(run: Element): boolean {
  for (const child of elementChildren(run)) {
    if (!SIMPLE_RUN_CHILDREN.has(child.tagName)) return false;
    if (child.tagName === "w:br" && child.getAttribute("w:type") && child.getAttribute("w:type") !== "textWrapping") return false;
  }
  return true;
}

function fldCharType(run: Element): string | null {
  const fc = firstChild(run, "w:fldChar");
  return fc ? fc.getAttribute("w:fldCharType") : null;
}

export function readParagraph(p: Element): ParagraphText {
  const atoms: Atom[] = [];
  let text = "";
  let tokenCount = 0;
  let editable = true;

  const pushOpaque = (nodes: Element[]) => {
    const visible = hasVisibleText(nodes);
    const token = visible ? `⟦${++tokenCount}⟧` : null;
    atoms.push({ kind: "opaque", nodes, token });
    if (token) text += token;
  };

  let fieldGroup: Element[] | null = null;
  let fieldDepth = 0;

  for (const child of elementChildren(p)) {
    if (child.tagName === "w:pPr") continue;

    if (fieldGroup) {
      fieldGroup.push(child);
      if (child.tagName === "w:r") {
        const t = fldCharType(child);
        if (t === "begin") fieldDepth++;
        else if (t === "end") fieldDepth--;
      }
      if (fieldDepth === 0) {
        pushOpaque(fieldGroup);
        fieldGroup = null;
      }
      continue;
    }

    if (child.tagName === "w:r") {
      const t = fldCharType(child);
      if (t === "begin") {
        fieldGroup = [child];
        fieldDepth = 1;
        continue;
      }
      if (t) {
        // A field that started in an earlier paragraph -- not safe to rebuild around.
        editable = false;
        pushOpaque([child]);
        continue;
      }
      if (!isSimpleRun(child)) {
        pushOpaque([child]);
        continue;
      }
      for (const rc of elementChildren(child)) {
        let s = "";
        if (rc.tagName === "w:t") s = rc.textContent ?? "";
        else if (rc.tagName === "w:tab") s = "\t";
        else if (rc.tagName === "w:br") s = "\n";
        for (const ch of s) {
          atoms.push({ kind: "char", ch, run: child, red: false });
          text += ch;
        }
      }
      continue;
    }

    if (DROPPABLE.has(child.tagName)) continue;
    pushOpaque([child]);
  }

  // A field still open at the paragraph end spans paragraphs (e.g. the TOC) -- leave it alone.
  if (fieldGroup) {
    editable = false;
    pushOpaque(fieldGroup);
  }

  return { text, atoms, editable };
}

// --- Diff ---

function tokenize(s: string): string[] {
  return s.match(/⟦\d+⟧|[\p{L}\p{N}]+|\s+|[^\p{L}\p{N}\s]/gu) ?? [];
}

type Op = { op: "eq" | "del" | "ins"; tok: string };

function diffTokens(a: string[], b: string[]): Op[] {
  let start = 0;
  while (start < a.length && start < b.length && a[start] === b[start]) start++;
  let endA = a.length;
  let endB = b.length;
  while (endA > start && endB > start && a[endA - 1] === b[endB - 1]) {
    endA--;
    endB--;
  }
  const midA = a.slice(start, endA);
  const midB = b.slice(start, endB);
  const ops: Op[] = a.slice(0, start).map((tok) => ({ op: "eq" as const, tok }));

  if (midA.length * midB.length > 4_000_000) {
    for (const tok of midA) ops.push({ op: "del", tok });
    for (const tok of midB) ops.push({ op: "ins", tok });
  } else {
    const n = midA.length;
    const m = midB.length;
    const dp: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        dp[i][j] = midA[i] === midB[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (midA[i] === midB[j]) {
        ops.push({ op: "eq", tok: midA[i] });
        i++;
        j++;
      } else if (dp[i + 1][j] >= dp[i][j + 1]) {
        ops.push({ op: "del", tok: midA[i++] });
      } else {
        ops.push({ op: "ins", tok: midB[j++] });
      }
    }
    while (i < n) ops.push({ op: "del", tok: midA[i++] });
    while (j < m) ops.push({ op: "ins", tok: midB[j++] });
  }

  for (const tok of a.slice(endA)) ops.push({ op: "eq", tok });
  return ops;
}

export interface RedlineOutcome {
  applied: boolean;
  reason?: string;
  insertedWords: number;
  deletedWords: number;
}

/** Rewrites paragraph `p` to read `newText`, coloring inserted text red. */
export function applyRedline(p: Element, newText: string): RedlineOutcome {
  const para = readParagraph(p);
  if (!para.editable) return { applied: false, reason: "paragraph contains a field that spans paragraphs", insertedWords: 0, deletedWords: 0 };
  newText = newText.replace(/\r?\n/g, " ");
  if (newText === para.text) return { applied: false, reason: "no change", insertedWords: 0, deletedWords: 0 };

  const oldTokens = tokenize(para.text);
  const newTokens = tokenize(newText);
  if (oldTokens.join("") !== para.text || newTokens.join("") !== newText) {
    return { applied: false, reason: "could not tokenize text", insertedWords: 0, deletedWords: 0 };
  }
  const oldPlaceholders = oldTokens.filter((t) => t.startsWith("⟦"));
  const newPlaceholders = newTokens.filter((t) => t.startsWith("⟦"));
  if (oldPlaceholders.join() !== newPlaceholders.join()) {
    return { applied: false, reason: "edit would have altered a field, link, or content control", insertedWords: 0, deletedWords: 0 };
  }

  const ops = diffTokens(oldTokens, newTokens);
  if (ops.some((o) => o.op !== "eq" && o.tok.startsWith("⟦"))) {
    return { applied: false, reason: "edit would have moved a field, link, or content control", insertedWords: 0, deletedWords: 0 };
  }

  // Map each old token to its slice of atoms (zero-width opaque atoms attach to the token after
  // them; any trailing ones stay at the end).
  const tokenAtoms: Atom[][] = [];
  let cursor = 0;
  const atoms = para.atoms;
  for (const tok of oldTokens) {
    const slice: Atom[] = [];
    let consumed = 0;
    const need = tok.startsWith("⟦") ? 1 : [...tok].length;
    while (cursor < atoms.length && consumed < need) {
      const atom = atoms[cursor++];
      slice.push(atom);
      if (atom.kind === "char" || atom.token) consumed++;
    }
    tokenAtoms.push(slice);
  }
  const trailing = atoms.slice(cursor);

  const out: Atom[] = [];
  let oldIndex = 0;
  let lastRun: Element | null = null;
  let inserted = 0;
  let deleted = 0;
  const firstRun = (atoms.find((a) => a.kind === "char") as CharAtom | undefined)?.run ?? null;

  for (const op of ops) {
    if (op.op === "eq") {
      for (const atom of tokenAtoms[oldIndex]) {
        out.push(atom);
        if (atom.kind === "char") lastRun = atom.run;
      }
      oldIndex++;
    } else if (op.op === "del") {
      // Keep zero-width markers (bookmarks, comment anchors) that sat inside deleted text.
      for (const atom of tokenAtoms[oldIndex]) {
        if (atom.kind === "opaque") out.push(atom);
        else lastRun = atom.run;
      }
      if (/[\p{L}\p{N}]/u.test(op.tok)) deleted++;
      oldIndex++;
    } else {
      const source = lastRun ?? firstRun;
      if (!source) return { applied: false, reason: "no text run to copy formatting from", insertedWords: 0, deletedWords: 0 };
      for (const ch of op.tok) out.push({ kind: "char", ch, run: source, red: true });
      if (/[\p{L}\p{N}]/u.test(op.tok)) inserted++;
    }
  }
  out.push(...trailing);

  rebuildParagraph(p, out);
  return { applied: true, insertedWords: inserted, deletedWords: deleted };
}

function rebuildParagraph(p: Element, atoms: Atom[]) {
  const doc = p.ownerDocument;
  const pPr = firstChild(p, "w:pPr");
  // Detach everything except pPr; opaque atoms re-attach their own original nodes.
  for (const child of elementChildren(p)) if (child !== pPr) p.removeChild(child);
  for (let n = p.firstChild; n; ) {
    const next = n.nextSibling;
    if (n.nodeType !== 1) p.removeChild(n);
    n = next;
  }

  let i = 0;
  while (i < atoms.length) {
    const atom = atoms[i];
    if (atom.kind === "opaque") {
      for (const node of atom.nodes) p.appendChild(node);
      i++;
      continue;
    }
    // Group consecutive chars from the same source run with the same color.
    let j = i;
    let s = "";
    while (j < atoms.length) {
      const a = atoms[j];
      if (a.kind !== "char" || a.run !== atom.run || a.red !== atom.red) break;
      s += a.ch;
      j++;
    }
    p.appendChild(buildRun(doc, atom.run, s, atom.red));
    i = j;
  }
}

function buildRun(doc: Document, source: Element, text: string, red: boolean): Element {
  const run = createW(doc, "w:r");
  const srcRPr = firstChild(source, "w:rPr");
  if (srcRPr || red) {
    const rPr = srcRPr ? (srcRPr.cloneNode(true) as Element) : createW(doc, "w:rPr");
    if (red) {
      setOrderedChild(rPr, "w:color", RPR_ORDER, { "w:val": AI_RED });
      // Spell-check squiggles on brand-new words are noise during review.
      const noProof = firstChild(rPr, "w:noProof");
      if (noProof) rPr.removeChild(noProof);
    }
    run.appendChild(rPr);
  }
  let buf = "";
  const flush = () => {
    if (!buf) return;
    const t = createW(doc, "w:t");
    t.setAttributeNS(XML_NS, "xml:space", "preserve");
    t.appendChild(doc.createTextNode(buf));
    run.appendChild(t);
    buf = "";
  };
  for (const ch of text) {
    if (ch === "\t") {
      flush();
      run.appendChild(createW(doc, "w:tab"));
    } else if (ch === "\n") {
      flush();
      run.appendChild(createW(doc, "w:br"));
    } else buf += ch;
  }
  flush();
  return run;
}
