/**
 * Cross-report comparison (Brett, 2026-10-13): one unit usually gets several inspections the same
 * day -- borescope, PT, ECI, UT -- each with its own report. The reports differ, but some things
 * must match across all of them:
 *  - Inspection Details: Utility, Site Name, Unit Number, Serial Number, Type of Fuel, Site
 *    Contact, Operational Data (fired hours / starts / trips);
 *  - visual observations for the components more than one report covers (Inlet, VIGV, the first
 *    compressor stages);
 *  - significant findings / recommendations for those same components.
 *
 * Each report is reduced to three positional lists (detail rows, Observations rows, Significant
 * Observations paragraphs). Claude compares them and points at items by id; the ids map back to
 * list positions, which are recomputed on each *reviewed* file to place Word comments -- the
 * review never adds or removes rows or recommendation paragraphs, so positions stay valid.
 */
import JSZip from "jszip";
import { CommentWriter } from "../docx/comments";
import { childrenThroughWrappers, descendants, parseXml, serializeXml, textOf } from "../docx/ooxml";
import { callStructured } from "./ai";
import { buildModel, DocModel, normalizeHeading } from "./model";
import type { ReviewItem } from "./types";

type Kind = "D" | "O" | "S";

interface Extracted {
  /** Inspection Details rows. */
  D: { text: string; anchor: Element }[];
  /** Inlet / compressor Observations-table rows. */
  O: { text: string; anchor: Element }[];
  /** Significant Observations and Recommendations paragraphs. */
  S: { text: string; anchor: Element }[];
  /** Cover-page lines (report title, utility, site/unit) -- context for the AI only. */
  cover: string;
}

function clean(s: string): string {
  return s.replace(/\s+/g, " ").trim();
}

function firstTextParagraph(el: Element): Element | null {
  return descendants(el, "w:p").find((p) => textOf(p).trim()) ?? descendants(el, "w:p")[0] ?? null;
}

export function extractForComparison(model: DocModel): Extracted {
  const out: Extracted = { D: [], O: [], S: [], cover: "" };
  out.cover = model.paragraphs
    .filter((p) => p.zone === "cover" && p.text.trim())
    .map((p) => clean(p.text))
    .join(" / ")
    .slice(0, 400);

  // Inspection Details: the main table right under the heading (before Inspection Areas). One-off
  // reports without that heading: any table that has the tell-tale labels.
  const detailTables: Element[] = [];
  for (const p of model.paragraphs) {
    if (!p.table || p.table.kind === "photo" || p.table.nested || detailTables.includes(p.table.topTbl)) continue;
    const underDetails = p.h1 && normalizeHeading(p.h1) === "inspection details" && !p.h2;
    if (underDetails || /serial number|site contact|type of fuel/i.test(textOf(p.table.topTbl))) detailTables.push(p.table.topTbl);
  }
  for (const tbl of detailTables) {
    for (const row of childrenThroughWrappers(tbl, "w:tr")) {
      const cells = childrenThroughWrappers(row, "w:tc");
      const text = cells.map((c) => clean(textOf(c))).join(" | ");
      const anchor = firstTextParagraph(row);
      if (text.replace(/[|\s]/g, "") && anchor) out.D.push({ text, anchor });
    }
  }

  // Observations rows for the components reports overlap on: Inlet and Compressor sections (or a
  // table with no subsection heading, as NDE reports often have). Combustion / Turbine / Exhaust
  // never overlap with the NDE inspections.
  const seenRows = new Set<Element>();
  for (const p of model.paragraphs) {
    if (p.table?.kind !== "observation" || p.table.nested || seenRows.has(p.table.row)) continue;
    seenRows.add(p.table.row);
    if (p.h2 && !/inlet|compressor|igv|guide vane/i.test(p.h2)) continue;
    const cells = childrenThroughWrappers(p.table.row, "w:tc");
    const text = cells.map((c) => clean(textOf(c))).join(" | ");
    const anchorCell = cells[cells.length > 1 ? 1 : 0];
    const anchor = anchorCell ? firstTextParagraph(anchorCell) : null;
    if (text.replace(/[|\s]/g, "") && anchor) out.O.push({ text: `${p.h2 ?? p.h1 ?? ""} | ${text}`, anchor });
  }

  // Significant Observations and Recommendations (NDE reports may say "Significant Findings").
  for (const p of model.paragraphs) {
    if (p.zone !== "body" || p.headingLevel || !p.h1 || !/significant/i.test(p.h1)) continue;
    if (!/[A-Za-z]{3,}/.test(p.text)) continue;
    out.S.push({ text: clean(p.text).slice(0, 700), anchor: p.el });
  }
  return out;
}

export async function extractFromDocx(buffer: Buffer): Promise<Extracted> {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = parseXml(await zip.file("word/document.xml")!.async("text"));
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? parseXml(await stylesEntry.async("text")) : null;
  return extractForComparison(buildModel(documentXml, stylesXml));
}

// --- The AI comparison ---

export interface CompareFinding {
  /** The field or component, e.g. "Serial Number", "R1 observations". */
  item: string;
  /** Ids like "A.D2", "B.O5", "C.S1" -- every report involved. */
  refs: string[];
  message: string;
}

const COMPARE_SCHEMA = {
  type: "object",
  properties: {
    findings: {
      type: "array",
      items: {
        type: "object",
        properties: {
          item: { type: "string", description: "The field or component that doesn't match, e.g. \"Serial Number\" or \"R1 observations\"." },
          refs: {
            type: "array",
            items: { type: "string" },
            description: "Ids of the rows/paragraphs involved, one per report that has the item (e.g. [\"A.D2\", \"B.D2\"]).",
          },
          message: {
            type: "string",
            description:
              "One or two sentences naming each report by its file name and what it says, e.g. 'Serial Number is 248914 in the Borescope report but 248941 in the PT report.'",
          },
        },
        required: ["item", "refs", "message"],
        additionalProperties: false,
      },
    },
  },
  required: ["findings"],
  additionalProperties: false,
};

export interface NamedExtract {
  letter: string;
  fileName: string;
  data: Extracted;
}

function describe(r: NamedExtract): string {
  const section = (kind: Kind, title: string) =>
    `${title}:\n${r.data[kind].length ? r.data[kind].map((x, i) => `${r.letter}.${kind}${i + 1} | ${x.text}`).join("\n") : "(none found)"}`;
  return [
    `=== REPORT ${r.letter}: "${r.fileName}" ===`,
    `Cover page: ${r.data.cover || "(none)"}`,
    section("D", "INSPECTION DETAILS (rows)"),
    section("O", "OBSERVATIONS - Inlet / Compressor (section | cells)"),
    section("S", "SIGNIFICANT OBSERVATIONS AND RECOMMENDATIONS"),
  ].join("\n");
}

export async function compareReports(reports: NamedExtract[]): Promise<CompareFinding[]> {
  const system = `You are checking a set of inspection reports written for Advanced Turbine Support (ATS) about the SAME turbine unit, from inspections done around the same time -- typically a borescope (visual) inspection plus NDE inspections such as PT (dye penetrant), ECI (eddy current) and UT (ultrasonic). Each report is different, but ATS requires certain things to match across all of them.

Report a finding wherever these do NOT match between reports:
1. Inspection Details: Utility / Client, Site Name, Unit Number, Serial Number, Type of Fuel, Site Contact, and the Operational Data (total fired hours, starts, fired starts, manual starts, unit trips). Flag any difference in value, even small ones (a typo, a different abbreviation, a transposed digit). Inspection Type, Inspectors and Inspection Date are expected to differ -- don't flag them.
2. Visual observations for the same component in more than one report -- Inlet, VIGV / IGV, and the first compressor stages (R0-R5, S0-S5 or as present). Compare only components that appear in at least two reports. NDE results (indications found by PT/ECI/UT) are not visual observations and are not expected in the borescope report -- don't flag their absence. Do flag visual findings that contradict each other (different conditions, counts, or locations for the same component), or a visual condition that appears in one report's visual observations but is missing from another report that also lists visual observations for that component.
3. Significant findings / recommendations for those same components: the same condition should get the same finding and recommendation in every report that covers it. Flag contradictions or a significant finding present in one report but missing from another report that covers that component.

Treat wording that means the same thing as a match ("Impact damage at the leading edge" vs "Leading edge impact damage"). Each finding must list the ids (like A.D2, B.D2) of the rows or paragraphs involved -- one per report that has the item; when a report is missing the item entirely, mention that in the message instead. Name reports by their file names in the message. Return an empty list if everything that should match does.`;
  const user = reports.map(describe).join("\n\n");
  const result = await callStructured<{ findings: CompareFinding[] }>(system, user, COMPARE_SCHEMA, "high");
  return result.findings;
}

// --- Writing findings into the reviewed files ---

/** Adds each finding involving `letter` as a Word comment on that report's matching row or
 *  paragraph (or the report's first heading when the report lacks the item). */
export async function addComparisonComments(
  buffer: Buffer,
  letter: string,
  findings: CompareFinding[],
  others: Map<string, string>,
): Promise<{ buffer: Buffer; added: number }> {
  const mine = findings.filter((f) => f.refs.some((r) => r.startsWith(`${letter}.`)) || f.refs.length === 0);
  if (!mine.length) return { buffer, added: 0 };

  const zip = await JSZip.loadAsync(buffer);
  const documentXml = parseXml(await zip.file("word/document.xml")!.async("text"));
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? parseXml(await stylesEntry.async("text")) : null;
  const model = buildModel(documentXml, stylesXml);
  const data = extractForComparison(model);
  const fallback = model.paragraphs.find((p) => p.headingLevel === 1 && p.zone === "body")?.el ?? model.paragraphs.find((p) => p.text.trim())?.el;

  const comments = new CommentWriter(zip);
  await comments.init();
  for (const f of mine) {
    const ref = f.refs.find((r) => r.startsWith(`${letter}.`));
    let anchor: Element | null | undefined = null;
    const m = ref ? /^[A-Z]\.([DOS])(\d+)$/.exec(ref) : null;
    if (m) anchor = data[m[1] as Kind][Number(m[2]) - 1]?.anchor;
    anchor = anchor ?? fallback;
    if (!anchor) continue;
    const otherNames = [...new Set(f.refs.map((r) => r.split(".")[0]).filter((l) => l !== letter))]
      .map((l) => others.get(l))
      .filter(Boolean);
    const withText = otherNames.length ? ` (compared with ${otherNames.join(", ")})` : "";
    comments.add(anchor, `Cross-report check${withText} -- ${f.item}: ${f.message}`);
  }
  await comments.save();
  zip.file("word/document.xml", serializeXml(documentXml));
  return { buffer: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), added: comments.count };
}

export function findingsToItems(findings: CompareFinding[], names: Map<string, string>): ReviewItem[] {
  return findings.map((f) => ({
    location: f.item,
    note: f.message,
    before: [...new Set(f.refs.map((r) => names.get(r.split(".")[0])).filter(Boolean))].join(" · ") || undefined,
    level: "attention" as const,
  }));
}
