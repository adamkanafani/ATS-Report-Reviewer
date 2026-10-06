/**
 * Photo resolution: the equivalent of Word's File > Options > Advanced > Image Size and Quality
 * > "Default resolution: 150 ppi", minus the operational data photo(s).
 *
 * Word's setting is document-wide -- it would also crush the control-room HMI screenshot that
 * has to stay readable -- so instead of relying on it, every other photo is actually resampled
 * here to 150 pixels per inch at the size it's displayed in the report (accounting for any
 * cropping), and the operational data photo's bytes are never touched. The document's default
 * resolution setting is also set to 150 ppi so photos an inspector adds later follow suit.
 */
import type JSZip from "jszip";
import sharp from "sharp";
import { descendants, firstChild, parseXml } from "../docx/ooxml";
import { buildModel } from "./model";
import type { ReviewLog } from "./report";

const TARGET_PPI = 150;
const EMU_PER_INCH = 914400;

interface MediaNeed {
  widthPx: number;
  heightPx: number;
  operationalData: boolean;
}

function parseRels(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]+)"/.exec(m[0])?.[1];
    const external = /TargetMode="External"/.test(m[0]);
    if (!id || !target || external) continue;
    map.set(id, target.startsWith("/") ? target.slice(1) : `word/${target}`.replace(/\/\.\//g, "/"));
  }
  return map;
}

export async function setDefaultImageDpi(zip: JSZip) {
  const entry = zip.file("word/settings.xml");
  if (!entry) return;
  let xml = await entry.async("text");
  const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
  if (/<w14:defaultImageDpi\b/.test(xml)) {
    xml = xml.replace(/<w14:defaultImageDpi\b[^>]*\/>/, `<w14:defaultImageDpi w14:val="${TARGET_PPI}"/>`);
  } else {
    if (!/xmlns:w14=/.test(xml)) xml = xml.replace(/<w:settings\b/, `<w:settings xmlns:w14="${W14}"`);
    if (/mc:Ignorable="([^"]*)"/.test(xml)) {
      xml = xml.replace(/mc:Ignorable="([^"]*)"/, (m, list: string) => (list.split(" ").includes("w14") ? m : `mc:Ignorable="${list} w14"`));
    }
    const tag = `<w14:defaultImageDpi w14:val="${TARGET_PPI}"/>`;
    // Word writes its w14 settings at the tail, after w:listSeparator and before any w15/w16 ones.
    const later = /<(w15|w16[a-z]*):[A-Za-z]+\b/.exec(xml);
    if (later) xml = xml.slice(0, later.index) + tag + xml.slice(later.index);
    else xml = xml.replace("</w:settings>", `${tag}</w:settings>`);
  }
  zip.file("word/settings.xml", xml);
}

export async function resamplePhotos(zip: JSZip, log: ReviewLog) {
  const documentXml = parseXml(await zip.file("word/document.xml")!.async("text"));
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? parseXml(await stylesEntry.async("text")) : null;
  const rels = parseRels((await zip.file("word/_rels/document.xml.rels")?.async("text")) ?? "");
  const model = buildModel(documentXml, stylesXml);

  // Pictures belonging to operational data: inside an "Operational Data" photo table, or a
  // free-standing picture whose own paragraph or a neighbor (caption) mentions it.
  const opDataDrawings = new Set<Element>();
  for (const table of model.photoTables) {
    if (!table.isOperationalData) continue;
    for (const d of descendants(table.topTbl, "w:drawing")) opDataDrawings.add(d);
  }
  model.paragraphs.forEach((p, i) => {
    if (p.table) return;
    const drawings = descendants(p.el, "w:drawing");
    if (!drawings.length) return;
    const nearby = [model.paragraphs[i - 1], p, model.paragraphs[i + 1], model.paragraphs[i + 2]]
      .filter(Boolean)
      .map((q) => q.text)
      .join(" ");
    if (/operational\s+data/i.test(nearby) || /operational\s+data/i.test(p.h2 ?? "")) drawings.forEach((d) => opDataDrawings.add(d));
  });

  const needs = new Map<string, MediaNeed>();
  for (const drawing of descendants(documentXml.documentElement, "w:drawing")) {
    const container = firstChild(drawing, "wp:inline") ?? firstChild(drawing, "wp:anchor");
    const extent = container ? firstChild(container, "wp:extent") : null;
    const blip = descendants(drawing, "a:blip")[0];
    const relId = blip?.getAttribute("r:embed");
    const path = relId ? rels.get(relId) : undefined;
    if (!extent || !path) continue;
    const cx = Number(extent.getAttribute("cx"));
    const cy = Number(extent.getAttribute("cy"));
    if (!cx || !cy) continue;

    const srcRect = descendants(drawing, "a:srcRect")[0];
    const crop = (attr: string) => Math.max(0, Number(srcRect?.getAttribute(attr) || 0)) / 100000;
    const keepW = Math.max(0.05, 1 - crop("l") - crop("r"));
    const keepH = Math.max(0.05, 1 - crop("t") - crop("b"));

    const widthPx = ((cx / EMU_PER_INCH) * TARGET_PPI) / keepW;
    const heightPx = ((cy / EMU_PER_INCH) * TARGET_PPI) / keepH;
    const prev = needs.get(path);
    needs.set(path, {
      widthPx: Math.max(prev?.widthPx ?? 0, widthPx),
      heightPx: Math.max(prev?.heightPx ?? 0, heightPx),
      operationalData: (prev?.operationalData ?? false) || opDataDrawings.has(drawing),
    });
  }

  let resampled = 0;
  let skippedOpData = 0;
  let bytesBefore = 0;
  let bytesAfter = 0;
  let failed = 0;

  for (const [path, need] of needs) {
    if (need.operationalData) {
      skippedOpData++;
      continue;
    }
    const ext = path.split(".").pop()?.toLowerCase();
    if (ext !== "jpg" && ext !== "jpeg" && ext !== "png") continue;
    const entry = zip.file(path);
    if (!entry) continue;
    const bytes = await entry.async("nodebuffer");
    try {
      const meta = await sharp(bytes).metadata();
      if (!meta.width || !meta.height) continue;
      // EXIF orientations 5-8 store the pixels rotated 90 degrees from how they display.
      const rotated = (meta.orientation ?? 1) >= 5;
      const pxW = rotated ? meta.height : meta.width;
      const pxH = rotated ? meta.width : meta.height;
      const scale = Math.max(need.widthPx / pxW, need.heightPx / pxH);
      if (scale >= 0.95) continue; // already at (or below) 150 ppi for its displayed size
      const targetW = Math.max(1, Math.round(meta.width * scale));
      const targetH = Math.max(1, Math.round(meta.height * scale));
      let pipeline = sharp(bytes).resize(targetW, targetH, { fit: "fill" }).withMetadata({ density: TARGET_PPI });
      pipeline = ext === "png" ? pipeline.png({ compressionLevel: 9 }) : pipeline.jpeg({ quality: 90 });
      const out = await pipeline.toBuffer();
      if (out.length >= bytes.length) continue;
      zip.file(path, out);
      resampled++;
      bytesBefore += bytes.length;
      bytesAfter += out.length;
    } catch {
      failed++;
    }
  }

  await setDefaultImageDpi(zip);

  const mb = (n: number) => (n / 1024 / 1024).toFixed(1);
  log.add("photos", {
    note: resampled
      ? `Resampled ${resampled} photo${resampled === 1 ? "" : "s"} to 150 ppi (${mb(bytesBefore)} MB -> ${mb(bytesAfter)} MB).`
      : "No photos were above 150 ppi -- nothing to resample.",
  });
  log.add("photos", { note: "Set the document's default image resolution (Word > Options > Advanced) to 150 ppi." });
  if (skippedOpData) {
    log.add("photos", {
      note: `Left ${skippedOpData} operational data photo${skippedOpData === 1 ? "" : "s"} at full resolution.`,
    });
  } else {
    log.add("photos", {
      note: "No operational data photo was found (looked for a photo labeled \"Operational Data\"), so every photo was eligible for 150 ppi.",
      level: "attention",
    });
  }
  if (failed) log.add("photos", { note: `${failed} photo${failed === 1 ? "" : "s"} couldn't be read and were left unchanged.`, level: "attention" });
}
