/**
 * Photo date stamps vs. the inspection date (Brett, 2026-11-10): inspectors sometimes start a new
 * report from an old one instead of the template, and photos from the previous inspection get
 * left in -- their date stamps are months out.
 *
 * Each inspection photo's date comes from the camera's EXIF data when the file still has it
 * (about 1 in 10 pasted borescope photos do), otherwise from the date/time stamp burned into the
 * image itself: the top and bottom strips of every photo are tiled into a few contact-sheet
 * images and Claude reads the stamps off them. Photos dated more than two weeks before the
 * inspection date (or after it) get a Word comment.
 */
import type JSZip from "jszip";
import sharp, { type Metadata, type OverlayOptions } from "sharp";
import { childrenThroughWrappers, descendants, firstChild, textOf } from "../docx/ooxml";
import { callStructured, hasApiKey, mapLimit } from "./ai";
import type { DocModel } from "./model";

const EMU_PER_INCH = 914400;
const STRIP_WIDTH = 420;
const ROWS_PER_SHEET = 30;
/** Multi-day inspections: photos up to two weeks before the (completion) date are normal. */
const DAYS_BEFORE_OK = 14;
const DAYS_AFTER_OK = 2;

export interface PhotoDateFinding {
  anchor: Element;
  where: string;
  message: string;
}

export interface PhotoDateResult {
  findings: PhotoDateFinding[];
  checked: number;
  dated: number;
  inspectionDate: string | null;
  skippedReason?: string;
}

interface Photo {
  anchor: Element;
  where: string;
  bytes: Buffer;
  date: Date | null;
}

// --- Inspection date ---

function parseDate(text: string): Date | null {
  const m =
    /\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b/i.exec(text) ??
    /\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/.exec(text);
  if (!m) return null;
  const d = new Date(m[0]);
  return Number.isNaN(d.getTime()) ? null : d;
}

/** The value under the "Inspection Date" header in the Inspection Details table, else a date on
 *  the cover page. */
export function findInspectionDate(model: DocModel): { date: Date; text: string } | null {
  const seen = new Set<Element>();
  for (const p of model.paragraphs) {
    if (!p.table || p.table.kind === "photo" || seen.has(p.table.topTbl)) continue;
    seen.add(p.table.topTbl);
    const rows = childrenThroughWrappers(p.table.topTbl, "w:tr");
    for (let r = 0; r < rows.length - 1; r++) {
      const cells = childrenThroughWrappers(rows[r], "w:tc");
      const col = cells.findIndex((c) => /^inspection date/i.test(textOf(c).trim()));
      if (col < 0) continue;
      const valueCell = childrenThroughWrappers(rows[r + 1], "w:tc")[col];
      const text = valueCell ? textOf(valueCell).trim() : "";
      const date = parseDate(text);
      if (date) return { date, text };
    }
  }
  for (const p of model.paragraphs) {
    if (p.zone !== "cover") continue;
    const date = parseDate(p.text);
    if (date) return { date, text: p.text.trim() };
  }
  return null;
}

// --- Photos ---

function parseRels(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
    if (id && target && !/TargetMode="External"/.test(m[0])) map.set(id, target.startsWith("/") ? target.slice(1) : `word/${target}`);
  }
  return map;
}

function exifDate(meta: Metadata): Date | null {
  if (!meta.exif) return null;
  const dates = [...meta.exif.toString("latin1").matchAll(/(\d{4}):(\d\d):(\d\d) (\d\d):(\d\d):(\d\d)/g)]
    .map((m) => new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6])))
    .filter((d) => !Number.isNaN(d.getTime()) && d.getFullYear() > 2000);
  // DateTime can be a later edit; the earliest is when the photo was taken.
  return dates.length ? new Date(Math.min(...dates.map((d) => d.getTime()))) : null;
}

async function collectPhotos(zip: JSZip, model: DocModel): Promise<Photo[]> {
  const rels = parseRels((await zip.file("word/_rels/document.xml.rels")?.async("text")) ?? "");
  const photos: Photo[] = [];
  for (const table of model.photoTables) {
    if (table.isOperationalData || table.isDataPlate) continue;
    for (const row of childrenThroughWrappers(table.topTbl, "w:tr")) {
      // The row's photo: its largest inline picture (the ATS watermark logo is tiny).
      let best: { blipId: string; area: number; el: Element } | null = null;
      for (const inline of descendants(row, "wp:inline")) {
        const extent = firstChild(inline, "wp:extent");
        const cx = Number(extent?.getAttribute("cx") ?? 0);
        const cy = Number(extent?.getAttribute("cy") ?? 0);
        const blipId = descendants(inline, "a:blip")[0]?.getAttribute("r:embed");
        if (!blipId || cx < 1.5 * EMU_PER_INCH) continue;
        if (!best || cx * cy > best.area) best = { blipId, area: cx * cy, el: inline };
      }
      if (!best) continue;
      const path = rels.get(best.blipId);
      const bytes = path ? await zip.file(path)?.async("nodebuffer") : undefined;
      if (!bytes) continue;

      // Comment anchor: the photo's Observation value, else its first caption text.
      let anchor: Element | null = null;
      const fields: Record<string, string> = {};
      for (const capRow of descendants(row, "w:tr")) {
        const cells = childrenThroughWrappers(capRow, "w:tc");
        if (cells.length < 2) continue;
        const label = textOf(cells[0]).trim().toLowerCase().replace(/:$/, "");
        fields[label] = textOf(cells[cells.length - 1]).trim();
        if (/^(observation|classification)$/.test(label)) anchor = descendants(cells[cells.length - 1], "w:p")[0] ?? anchor;
      }
      anchor = anchor ?? descendants(row, "w:p").find((q) => textOf(q).trim()) ?? null;
      if (!anchor) continue;
      const where = [fields.component, fields.location].filter(Boolean).join(", ") || table.h2 || "Photo";

      let date: Date | null = null;
      try {
        date = exifDate(await sharp(bytes).metadata());
      } catch {
        continue; // not an image sharp can read (EMF etc.)
      }
      photos.push({ anchor, where, bytes, date });
    }
  }
  return photos;
}

// --- Reading burned-in stamps ---

export async function stripsFor(bytes: Buffer): Promise<{ top: Buffer; bottom: Buffer; height: number } | null> {
  const img = sharp(bytes);
  const meta = await img.metadata();
  if (!meta.width || !meta.height || meta.width < 200) return null;
  const band = Math.max(12, Math.round(meta.height * 0.11));
  const top = await sharp(bytes).extract({ left: 0, top: 0, width: meta.width, height: band }).resize({ width: STRIP_WIDTH }).png().toBuffer();
  const bottom = await sharp(bytes)
    .extract({ left: 0, top: meta.height - band, width: meta.width, height: band })
    .resize({ width: STRIP_WIDTH })
    .png()
    .toBuffer();
  const height = (await sharp(top).metadata()).height ?? 40;
  return { top, bottom, height };
}

/** One contact sheet: each row is "<n> | top strip | bottom strip". */
export async function contactSheet(rows: { n: number; top: Buffer; bottom: Buffer; height: number }[]): Promise<Buffer> {
  const labelW = 56;
  const gap = 8;
  const width = labelW + STRIP_WIDTH * 2 + gap * 3;
  const rowH = (r: { height: number }) => r.height + gap;
  const height = rows.reduce((h, r) => h + rowH(r), gap);
  const composites: OverlayOptions[] = [];
  let y = gap;
  for (const r of rows) {
    const label = Buffer.from(
      `<svg width="${labelW}" height="${r.height}" xmlns="http://www.w3.org/2000/svg"><rect width="100%" height="100%" fill="#ffffff"/><text x="6" y="${Math.min(r.height - 6, 28)}" font-family="Arial" font-size="22" font-weight="bold" fill="#c00000">${r.n}</text></svg>`,
    );
    composites.push({ input: label, left: 0, top: y });
    composites.push({ input: r.top, left: labelW + gap, top: y });
    composites.push({ input: r.bottom, left: labelW + gap * 2 + STRIP_WIDTH, top: y });
    y += rowH(r);
  }
  return sharp({ create: { width, height, channels: 3, background: "#ffffff" } }).composite(composites).png().toBuffer();
}

const STAMP_SCHEMA = {
  type: "object",
  properties: {
    stamps: {
      type: "array",
      items: {
        type: "object",
        properties: {
          row: { type: "integer" },
          date: { type: "string", description: "The stamp's date as YYYY-MM-DD, or an empty string if no date stamp is visible." },
        },
        required: ["row", "date"],
        additionalProperties: false,
      },
    },
  },
  required: ["stamps"],
  additionalProperties: false,
};

async function readStamps(photos: Photo[]): Promise<void> {
  const undated = photos.map((p, i) => ({ p, n: i + 1 })).filter((x) => !x.p.date);
  const prepared: { n: number; top: Buffer; bottom: Buffer; height: number }[] = [];
  for (const { p, n } of undated) {
    try {
      const strips = await stripsFor(p.bytes);
      if (strips) prepared.push({ n, ...strips });
    } catch {
      // unreadable image -- leave undated
    }
  }
  const sheets: (typeof prepared)[] = [];
  for (let i = 0; i < prepared.length; i += ROWS_PER_SHEET) sheets.push(prepared.slice(i, i + ROWS_PER_SHEET));

  await mapLimit(sheets, 3, async (rows) => {
    const sheet = await contactSheet(rows);
    const result = await callStructured<{ stamps: { row: number; date: string }[] }>(
      "You read the date/time stamps burned into borescope inspection photos.",
      [
        { type: "image", source: { type: "base64", media_type: "image/png", data: sheet.toString("base64") } },
        {
          type: "text",
          text: `Each numbered row of this image shows two strips cut from one borescope photo: its top edge (left) and its bottom edge (right). Most photos have a date/time stamp printed in one of those strips, usually top-right, like "10/9/2026 11:08 AM" (US month/day/year). For every row number (${rows
            .map((r) => r.n)
            .join(", ")}), return the stamp's date as YYYY-MM-DD, or an empty string when no date stamp is visible in that row. Read only what is printed -- never guess.`,
        },
      ],
      STAMP_SCHEMA,
      "low",
    );
    for (const s of result.stamps) {
      const photo = photos[s.row - 1];
      if (!photo || photo.date || !/^\d{4}-\d{2}-\d{2}$/.test(s.date)) continue;
      const [y, m, d] = s.date.split("-").map(Number);
      const date = new Date(y, m - 1, d);
      if (!Number.isNaN(date.getTime())) photo.date = date;
    }
  });
}

// --- Main ---

function formatDate(d: Date): string {
  return d.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
}

function describeGap(days: number): string {
  const abs = Math.abs(days);
  if (abs >= 60) return `about ${Math.round(abs / 30.4)} months`;
  return `${abs} days`;
}

export async function checkPhotoDates(zip: JSZip, model: DocModel): Promise<PhotoDateResult> {
  const inspection = findInspectionDate(model);
  if (!inspection) {
    return { findings: [], checked: 0, dated: 0, inspectionDate: null, skippedReason: "No inspection date was found in the Inspection Details table or on the cover page." };
  }
  const photos = await collectPhotos(zip, model);
  if (hasApiKey()) await readStamps(photos);

  const findings: PhotoDateFinding[] = [];
  const day = 86400000;
  const inspDay = new Date(inspection.date.getFullYear(), inspection.date.getMonth(), inspection.date.getDate()).getTime();
  for (const photo of photos) {
    if (!photo.date) continue;
    const photoDay = new Date(photo.date.getFullYear(), photo.date.getMonth(), photo.date.getDate()).getTime();
    const days = Math.round((photoDay - inspDay) / day);
    if (days >= -DAYS_BEFORE_OK && days <= DAYS_AFTER_OK) continue;
    findings.push({
      anchor: photo.anchor,
      where: photo.where,
      message:
        days < 0
          ? `Photo dated ${formatDate(photo.date)}, ${describeGap(days)} before the inspection date (${inspection.text}). It may be left over from a previous inspection or report.`
          : `Photo dated ${formatDate(photo.date)}, ${describeGap(days)} after the inspection date (${inspection.text}). Check the photo or the inspection date.`,
    });
  }
  return {
    findings,
    checked: photos.length,
    dated: photos.filter((p) => p.date).length,
    inspectionDate: inspection.text,
    skippedReason: !hasApiKey() ? "No Anthropic API key, so only photos with a camera date in the file could be checked." : undefined,
  };
}
