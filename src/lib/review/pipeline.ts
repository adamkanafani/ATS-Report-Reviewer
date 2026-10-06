/**
 * The whole review, start to finish: inspector's .docx in, reviewed .docx + change log out.
 *
 *   1. Parse and model the report (model.ts).
 *   2. Deterministic formatting (formatting.ts) -- including turning all text black, so after
 *      this point the only colored text is what the AI changes.
 *   3. AI passes, run concurrently (ai.ts): proofreading, Overall Assessment readability,
 *      photo-vs-Observations cross-check, heading style review.
 *   4. Apply AI results: style fixes, red word-level redlines (redline.ts), Word comments for
 *      findings that aren't edits (comments.ts).
 *   5. Table of Contents rebuilt by the user's own Word (word.ts), or flagged to update on open.
 *   6. Photos resampled to 150 ppi, operational data excepted (images.ts).
 */
import JSZip from "jszip";
import { CommentWriter } from "../docx/comments";
import { childrenThroughWrappers, descendants, parseXml, serializeXml, textOf } from "../docx/ooxml";
import {
  crossCheckObservations,
  hasApiKey,
  HeadingCandidate,
  mapLimit,
  ProofEdit,
  ProofItem,
  proofreadBatch,
  reviewHeadings,
  reviewOverallAssessment,
} from "./ai";
import { applyFormatting, ensureHeadingStyles, setParagraphStyle } from "./formatting";
import { resamplePhotos } from "./images";
import { buildModel, DocModel, locationLabel, normalizeHeading, ParaInfo, PHOTO_CAPTION_LABELS, refreshHeadingContext } from "./model";
import { applyRedline, readParagraph } from "./redline";
import { ReviewLog } from "./report";
import type { ReviewCategory, ReviewOptions } from "./types";
import { flagUpdateFieldsOnOpen, updateTocWithWord } from "./word";

export interface ReviewOutput {
  buffer: Buffer;
  categories: ReviewCategory[];
  aiEditCount: number;
  commentCount: number;
}

type Progress = (message: string) => void;

const PROOF_BATCH_CHARS = 9000;
const PROOF_BATCH_ITEMS = 120;
const AI_CONCURRENCY = 4;

function isOverallAssessment(p: ParaInfo): boolean {
  return !!p.h1 && normalizeHeading(p.h1) === "overall assessment" && !p.headingLevel;
}

function hasWords(text: string): boolean {
  return /[A-Za-z]{2,}/.test(text.replace(/⟦\d+⟧/g, ""));
}

function rowFirstCellText(p: ParaInfo): string | null {
  if (!p.table) return null;
  const cells = childrenThroughWrappers(p.table.row, "w:tc");
  return cells.length > 1 ? textOf(cells[0]).trim() : null;
}

function shorten(s: string, n = 220): string {
  s = s.replace(/⟦\d+⟧/g, "…").replace(/\s+/g, " ").trim();
  return s.length > n ? `${s.slice(0, n - 1)}…` : s;
}

// --- Proofreading item collection (deduplicated) ---

interface ProofGroup {
  item: ProofItem;
  paragraphs: ParaInfo[];
}

function collectProofGroups(model: DocModel, skipOverallAssessment: boolean): ProofGroup[] {
  const groups = new Map<string, ProofGroup>();
  let n = 0;
  for (const p of model.paragraphs) {
    if (p.zone !== "cover" && p.zone !== "body") continue;
    if (skipOverallAssessment && isOverallAssessment(p)) continue;
    if (!hasWords(p.text)) continue;
    if (PHOTO_CAPTION_LABELS.has(p.text.trim().toLowerCase().replace(/:$/, ""))) continue;
    const read = readParagraph(p.el);
    if (!read.editable || !hasWords(read.text)) continue;
    // Photo captions repeat the same handful of values hundreds of times ("Typical", "Stage 1
    // Rotor Blade", ...) -- send each distinct value once, apply its fix everywhere.
    const kind = p.table?.kind === "photo" ? `photo:${p.table.rowLabel ?? ""}` : (p.table?.kind ?? "paragraph");
    const key = `${kind}\u0000${read.text}`;
    const existing = groups.get(key);
    if (existing) {
      existing.paragraphs.push(p);
      continue;
    }
    let context = locationLabel(p);
    const rowLabel = p.table?.kind !== "photo" ? rowFirstCellText(p) : null;
    if (rowLabel && rowLabel !== p.text.trim()) context += ` (row: ${shorten(rowLabel, 60)})`;
    groups.set(key, { item: { id: `t${++n}`, context, text: read.text }, paragraphs: [p] });
  }
  return [...groups.values()];
}

function batchItems(groups: ProofGroup[]): ProofItem[][] {
  const batches: ProofItem[][] = [];
  let current: ProofItem[] = [];
  let chars = 0;
  for (const g of groups) {
    const size = g.item.text.length + g.item.context.length + 30;
    if (current.length && (chars + size > PROOF_BATCH_CHARS || current.length >= PROOF_BATCH_ITEMS)) {
      batches.push(current);
      current = [];
      chars = 0;
    }
    current.push(g.item);
    chars += size;
  }
  if (current.length) batches.push(current);
  return batches;
}

// --- Cross-check inputs ---

interface Anchored {
  id: string;
  anchor: Element;
  line: string;
}

function collectObservationRows(model: DocModel): Anchored[] {
  const rows: Anchored[] = [];
  const seen = new Set<Element>();
  for (const p of model.paragraphs) {
    if (p.table?.kind !== "observation" || p.table.nested || seen.has(p.table.row)) continue;
    seen.add(p.table.row);
    const cells = childrenThroughWrappers(p.table.row, "w:tc").map((c) => shorten(textOf(c), 300));
    if (!cells.some((c) => c)) continue;
    const anchorCell = childrenThroughWrappers(p.table.row, "w:tc")[cells.length > 1 ? 1 : 0];
    const anchor = descendants(anchorCell, "w:p").find((q) => textOf(q).trim()) ?? descendants(anchorCell, "w:p")[0] ?? p.el;
    const id = `O${rows.length + 1}`;
    rows.push({ id, anchor, line: `${id} | ${p.h2 ?? p.h1 ?? ""} | ${cells.join(" | ")}` });
  }
  return rows;
}

function collectPhotoEntries(model: DocModel): Anchored[] {
  const entries: Anchored[] = [];
  for (const table of model.photoTables) {
    if (table.isOperationalData || table.isDataPlate) continue;
    const nested = descendants(table.topTbl, "w:tbl").filter((t) => t !== table.topTbl);
    const units = nested.length ? nested : [table.topTbl];
    for (const unit of units) {
      const fields: string[] = [];
      let anchor: Element | null = null;
      let firstValue: Element | null = null;
      for (const row of childrenThroughWrappers(unit, "w:tr")) {
        const cells = childrenThroughWrappers(row, "w:tc");
        if (cells.length < 2) continue;
        const label = textOf(cells[0]).trim().replace(/:$/, "");
        const valueCell = cells[cells.length - 1];
        const value = textOf(valueCell).trim();
        if (!label && !value) continue;
        fields.push(`${label}: ${value}`);
        const valuePara = descendants(valueCell, "w:p").find((q) => textOf(q).trim()) ?? null;
        if (!firstValue && valuePara) firstValue = valuePara;
        if (/^(observation|classification)$/i.test(label) && valuePara) anchor = valuePara;
      }
      if (!fields.length) {
        const text = shorten(textOf(unit), 300);
        if (!text) continue;
        fields.push(text);
      }
      const fallback = descendants(unit, "w:p").find((q) => textOf(q).trim()) ?? null;
      const anchorEl = anchor ?? firstValue ?? fallback;
      if (!anchorEl) continue;
      const id = `P${entries.length + 1}`;
      entries.push({ id, anchor: anchorEl, line: `${id} | ${table.h2 ?? ""} | ${shorten(fields.join("; "), 400)}` });
    }
  }
  return entries;
}

// --- Main ---

export async function runReview(input: Buffer, options: ReviewOptions, progress: Progress): Promise<ReviewOutput> {
  const log = new ReviewLog();
  progress("Reading the report...");
  const zip = await JSZip.loadAsync(input);
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) throw new Error("This doesn't look like a Word .docx file (no word/document.xml inside).");
  const documentXml = parseXml(await docEntry.async("text"));
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? parseXml(await stylesEntry.async("text")) : null;
  const model = buildModel(documentXml, stylesXml);

  progress("Applying formatting rules...");
  const { hasToc } = applyFormatting({ documentXml, stylesXml, model, log, options });

  // --- AI passes ---
  const aiWanted = options.proofreading || options.overallAssessment || options.crossCheck || options.headings;
  const comments = new CommentWriter(zip);
  await comments.init();
  let aiEditCount = 0;

  if (aiWanted && !hasApiKey()) {
    const reason = "No Anthropic API key is set up on this computer (ANTHROPIC_API_KEY in .env.local), so the AI checks were skipped.";
    if (options.proofreading) log.skip("proofreading", reason);
    if (options.overallAssessment) log.skip("overallAssessment", reason);
    if (options.crossCheck) log.skip("crossCheck", reason);
  } else if (aiWanted) {
    const oaParas = model.paragraphs.filter((p) => isOverallAssessment(p) && hasWords(p.text));
    const doOa = options.overallAssessment && oaParas.length > 0;
    if (options.overallAssessment && !oaParas.length) log.skip("overallAssessment", "No Overall Assessment section was found in this report.");

    const proofGroups = options.proofreading ? collectProofGroups(model, doOa) : [];
    const proofBatches = batchItems(proofGroups);
    const obsRows = options.crossCheck ? collectObservationRows(model) : [];
    const photoEntries = options.crossCheck ? collectPhotoEntries(model) : [];
    const doCross = options.crossCheck && obsRows.length > 0 && photoEntries.length > 0;
    if (options.crossCheck && !doCross) {
      log.skip(
        "crossCheck",
        !obsRows.length ? "No Observations tables were found to compare against." : "No photo tables with captions were found to compare against.",
      );
    }

    const outline: HeadingCandidate[] = [];
    const outlineMap = new Map<string, ParaInfo>();
    if (options.headings) {
      for (const p of model.paragraphs) {
        if ((p.zone !== "body" && p.zone !== "llc") || p.table || !p.text.trim()) continue;
        if (p.text.trim().length > 90 && !p.headingLevel) continue;
        const id = `H${p.index}`;
        outline.push({ id, style: p.headingLevel ? `Heading ${p.headingLevel}` : "Normal", text: shorten(p.text, 90) });
        outlineMap.set(id, p);
      }
    }

    const totalJobs = proofBatches.length + (doOa ? 1 : 0) + (doCross ? 1 : 0) + (outline.length ? 1 : 0);
    let doneJobs = 0;
    const tick = (what: string) => progress(`AI review: ${what} (${++doneJobs} of ${totalJobs} done)`);
    progress(`AI review: sending ${totalJobs} request${totalJobs === 1 ? "" : "s"} to Claude...`);

    const oaItems: ProofItem[] = oaParas.map((p, i) => {
      const row = rowFirstCellText(p);
      return { id: `a${i + 1}`, context: row && row !== p.text.trim() ? `Overall Assessment (row: ${row})` : "Overall Assessment", text: readParagraph(p.el).text };
    });
    const supporting = [
      "SIGNIFICANT OBSERVATIONS AND RECOMMENDATIONS:",
      ...model.paragraphs
        .filter((p) => p.h1 && normalizeHeading(p.h1) === "significant observations and recommendations" && !p.headingLevel && p.text.trim())
        .map((p) => `- ${shorten(p.text, 600)}`),
      "",
      "OBSERVATIONS TABLES:",
      ...collectObservationRows(model).map((r) => r.line),
    ]
      .join("\n")
      .slice(0, 60000);

    const [proofResults, oaResult, crossResult, headingResult] = await Promise.allSettled([
      mapLimit(proofBatches, AI_CONCURRENCY, async (batch) => {
        const edits = await proofreadBatch(batch);
        tick("proofreading");
        return edits;
      }),
      doOa ? reviewOverallAssessment(oaItems, supporting).finally(() => tick("Overall Assessment")) : Promise.resolve(null),
      doCross
        ? crossCheckObservations(obsRows.map((r) => r.line).join("\n"), photoEntries.map((e) => e.line).join("\n")).finally(() =>
            tick("photo vs. observation check"),
          )
        : Promise.resolve(null),
      outline.length ? reviewHeadings(outline).finally(() => tick("heading styles")) : Promise.resolve(null),
    ]);

    progress("Applying AI edits in red...");

    // Heading fixes first, so later locations reflect the corrected outline.
    if (headingResult.status === "fulfilled" && headingResult.value) {
      const ids = ensureHeadingStyles(stylesXml, model);
      for (const fix of headingResult.value) {
        const p = outlineMap.get(fix.id);
        if (!p || p.decided) continue;
        const want = fix.style === "Heading 1" ? 1 : fix.style === "Heading 2" ? 2 : 0;
        if ((p.headingLevel ?? 0) === want) continue;
        if (want === 0 && (p.headingLevel ?? 0) > 2) continue;
        const before = p.headingLevel ? `Heading ${p.headingLevel}` : "Normal";
        setParagraphStyle(p, want ? ids[want as 1 | 2] : null);
        log.add("headings", { location: locationLabel(p), before, after: fix.style, note: `"${shorten(p.text, 80)}" set to ${fix.style}: ${fix.reason}` });
      }
      refreshHeadingContext(model);
    } else if (headingResult.status === "rejected") {
      log.add("headings", { note: `AI heading review failed: ${(headingResult.reason as Error).message}`, level: "attention" });
    }
    if (options.headings) log.touch("headings");

    // Proofreading edits.
    if (options.proofreading) {
      const byId = new Map(proofGroups.map((g) => [g.item.id, g]));
      if (proofResults.status === "fulfilled") {
        for (const edit of proofResults.value.flat()) {
          const group = byId.get(edit.id);
          if (!group) continue;
          aiEditCount += applyEditToGroup(group.paragraphs, edit, log, "proofreading");
        }
      } else {
        log.add("proofreading", { note: `Proofreading failed: ${(proofResults.reason as Error).message}`, level: "attention" });
      }
      log.touch("proofreading");
    }

    // Overall Assessment.
    if (doOa) {
      if (oaResult.status === "fulfilled" && oaResult.value) {
        for (const edit of oaResult.value.edits) {
          const idx = Number(edit.id.slice(1)) - 1;
          const p = oaParas[idx];
          if (p) aiEditCount += applyEditToGroup([p], edit, log, "overallAssessment");
        }
        const recs = oaResult.value.recommendations;
        for (const rec of recs) log.add("overallAssessment", { note: rec, level: "attention" });
        const heading = model.paragraphs.find((p) => p.headingLevel === 1 && normalizeHeading(p.text) === "overall assessment");
        if (recs.length && heading) {
          comments.add(heading.el, `Overall Assessment recommendations:\n${recs.map((r) => `• ${r}`).join("\n")}`);
        }
        if (!oaResult.value.edits.length && !recs.length) log.add("overallAssessment", { note: "The Overall Assessment already reads well -- no changes." });
      } else if (oaResult.status === "rejected") {
        log.add("overallAssessment", { note: `Overall Assessment review failed: ${(oaResult.reason as Error).message}`, level: "attention" });
      }
    }

    // Cross-check findings -> comments.
    if (doCross) {
      if (crossResult.status === "fulfilled" && crossResult.value) {
        const anchors = new Map<string, Anchored>([...obsRows, ...photoEntries].map((a) => [a.id, a]));
        const labels: Record<string, string> = {
          missing_photo: "No supporting photo",
          missing_observation: "Missing from Observations table",
          count_mismatch: "Count mismatch",
          location_mismatch: "Location mismatch",
          condition_mismatch: "Condition mismatch",
          other: "Check",
        };
        for (const f of crossResult.value) {
          const anchor = anchors.get(f.anchorId);
          const label = labels[f.type] ?? "Check";
          log.add("crossCheck", {
            location: anchor ? anchor.line.split(" | ")[1] || undefined : undefined,
            before: anchor ? shorten(anchor.line.split(" | ").slice(2).join(" | "), 200) : undefined,
            note: `${label}: ${f.message}`,
            level: "attention",
          });
          if (anchor) comments.add(anchor.anchor, `${label}: ${f.message}`);
        }
        if (!crossResult.value.length) {
          log.add("crossCheck", { note: `All ${photoEntries.length} photo captions and ${obsRows.length} Observations-table rows line up.` });
        }
      } else if (crossResult.status === "rejected") {
        log.add("crossCheck", { note: `Photo vs. observation check failed: ${(crossResult.reason as Error).message}`, level: "attention" });
      }
    }
  }

  // --- Serialize ---
  zip.file("word/document.xml", serializeXml(documentXml));
  if (stylesXml) zip.file("word/styles.xml", serializeXml(stylesXml));
  await comments.save();
  let buffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

  // --- Table of Contents ---
  if (options.toc && hasToc) {
    progress("Updating the Table of Contents in Microsoft Word...");
    const result = await updateTocWithWord(buffer);
    if (result.ok && result.buffer) {
      buffer = result.buffer;
      log.add("toc", { note: "Table of Contents rebuilt by Microsoft Word with current headings and page numbers." });
    } else {
      const fallback = await JSZip.loadAsync(buffer);
      const settings = fallback.file("word/settings.xml");
      if (settings) fallback.file("word/settings.xml", flagUpdateFieldsOnOpen(await settings.async("text")));
      buffer = await fallback.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
      log.add("toc", {
        note: `Word couldn't be used to rebuild the Table of Contents here (${result.error ?? "unknown error"}). The report is set to update it when opened -- click Yes when Word asks to update fields.`,
        level: "attention",
      });
    }
  }

  // --- Photos at 150 ppi ---
  if (options.photoResolution) {
    progress("Setting photos to 150 ppi...");
    const finalZip = await JSZip.loadAsync(buffer);
    await resamplePhotos(finalZip, log);
    buffer = await finalZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  }

  return { buffer, categories: log.toArray(), aiEditCount, commentCount: comments.count };
}

function applyEditToGroup(paragraphs: ParaInfo[], edit: ProofEdit, log: ReviewLog, key: "proofreading" | "overallAssessment"): number {
  let applied = 0;
  let failure: string | undefined;
  const before = readParagraph(paragraphs[0].el).text;
  for (const p of paragraphs) {
    const outcome = applyRedline(p.el, edit.text);
    if (outcome.applied) applied++;
    else if (outcome.reason !== "no change") failure = outcome.reason;
  }
  const where = paragraphs.length > 1 ? `${locationLabel(paragraphs[0])} (and ${paragraphs.length - 1} more place${paragraphs.length === 2 ? "" : "s"})` : locationLabel(paragraphs[0]);
  if (applied) {
    log.add(key, { location: where, before: shorten(before, 400), after: shorten(edit.text, 400), note: edit.change });
  } else if (failure) {
    log.add(key, {
      location: where,
      before: shorten(before, 400),
      after: shorten(edit.text, 400),
      note: `Suggested but not applied (${failure}): ${edit.change}`,
      level: "attention",
    });
  }
  return applied ? 1 : 0;
}

