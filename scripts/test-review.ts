/**
 * Runs the full review pipeline outside the web app.
 *
 *   npm run test:review -- <input.docx> [output.docx] [--no-ai] [--seed-errors] [--mock-ai]
 *
 * --seed-errors first damages a copy of the input the way a rushed one-off report would be
 * (typos, a colored run, Calibri text, a heading left as Normal, a Normal paragraph styled as a
 * heading) so there's something for every check to find on an already-approved report.
 *
 * --mock-ai swaps the Claude calls for canned responses (fixing the seeded typos, one
 * cross-check finding, one Overall Assessment edit + recommendation) so the redline/comment
 * application path can be verified -- and the output opened in Word -- without an API key.
 */
import { promises as fs } from "fs";
import path from "path";
import JSZip from "jszip";
import { DEFAULT_OPTIONS } from "../src/lib/review/types";
import { runReview } from "../src/lib/review/pipeline";
import * as ai from "../src/lib/review/ai";

async function loadEnvLocal() {
  try {
    const text = await fs.readFile(path.join(__dirname, "..", ".env.local"), "utf8");
    for (const line of text.split(/\r?\n/)) {
      const m = /^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/.exec(line);
      if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^"(.*)"$/, "$1");
    }
  } catch {
    // no .env.local -- fine
  }
}

async function seedErrors(input: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(input);
  let xml = await zip.file("word/document.xml")!.async("text");
  const swaps: [RegExp, string][] = [
    [/>The purpose of this borescope examination was/, ">The purpse of this borescope examinaton was"],
    [/>Impact damage at the leading edge/, ">impact Damage at the Leading edge"],
    [/We recommend performing/, "We recomend performing"],
    [/appears to be in good condition/, "appear to be in good conditon"],
    [/>Typical</, ">Typcal<"],
  ];
  for (const [re, to] of swaps) xml = xml.replace(re, to);
  xml = xml.replace(/(<w:pStyle w:val="Heading2"\/>)/, '<w:pStyle w:val="Normal"/>');
  let colored = 0;
  xml = xml.replace(/<w:rPr><w:sz w:val="22"\/>/g, (m) => (colored++ < 5 ? '<w:rPr><w:color w:val="1F4E79"/><w:sz w:val="22"/>' : m));
  zip.file("word/document.xml", xml);
  return zip.generateAsync({ type: "nodebuffer" });
}

function installMockAi() {
  const fixes: [RegExp, string][] = [
    [/purpse/, "purpose"],
    [/examinaton/, "examination"],
    [/^impact Damage at the Leading edge/, "Impact damage at the leading edge"],
    [/recomend/, "recommend"],
    [/appear to be in good conditon/, "appears to be in good condition"],
    [/^Typcal$/, "Typical"],
  ];
  const mock = ai as unknown as Record<string, unknown>;
  mock.hasApiKey = () => true;
  mock.proofreadBatch = async (items: ai.ProofItem[]) =>
    items.flatMap((item) => {
      let text = item.text;
      for (const [re, to] of fixes) text = text.replace(re, to);
      return text === item.text ? [] : [{ id: item.id, text, change: "Fixed spelling (mock)" }];
    });
  mock.reviewOverallAssessment = async (items: ai.ProofItem[]) => ({
    edits: items
      .filter((i) => /good conditon|good condition/.test(i.text))
      .slice(0, 1)
      .map((i) => ({ id: i.id, text: i.text.replace(/appear to be in good conditon/, "appears to be in good condition"), change: "Grammar (mock)" })),
    recommendations: ["Mock recommendation: mention the R-1 mold replica follow-up in the Compressor assessment."],
  });
  mock.crossCheckObservations = async () => [
    { anchorId: "O3", type: "count_mismatch", message: "Mock finding: table lists impact damage (6) but only 5 R1 photos show it." },
    { anchorId: "P5", type: "missing_observation", message: "Mock finding: photo shows a condition not in the Observations table." },
  ];
  mock.reviewHeadings = async () => [];
}

async function main() {
  await loadEnvLocal();
  const args = process.argv.slice(2);
  const flags = new Set(args.filter((a) => a.startsWith("--")));
  const [input, outputArg] = args.filter((a) => !a.startsWith("--"));
  if (!input) throw new Error("usage: test-review <input.docx> [output.docx] [--no-ai] [--seed-errors]");
  let buffer: Buffer = await fs.readFile(input);
  if (flags.has("--seed-errors")) buffer = await seedErrors(buffer);
  if (flags.has("--no-ai")) delete process.env.ANTHROPIC_API_KEY;
  if (flags.has("--mock-ai")) installMockAi();

  const output = outputArg ?? path.join("test-output", path.basename(input).replace(/\.docx$/i, " - Reviewed.docx"));
  await fs.mkdir(path.dirname(output), { recursive: true });

  const started = Date.now();
  const result = await runReview(buffer, DEFAULT_OPTIONS, (msg) => console.log(`[${((Date.now() - started) / 1000).toFixed(1)}s] ${msg}`));
  await fs.writeFile(output, result.buffer);

  for (const cat of result.categories) {
    console.log(`\n== ${cat.title} (${cat.items.length})${cat.skippedReason ? ` -- SKIPPED: ${cat.skippedReason}` : ""}`);
    for (const item of cat.items.slice(0, 25)) {
      console.log(`  - ${item.location ? `[${item.location}] ` : ""}${item.note}`);
      if (item.before) console.log(`      before: ${item.before.slice(0, 160)}`);
      if (item.after) console.log(`      after:  ${item.after.slice(0, 160)}`);
    }
    if (cat.items.length > 25) console.log(`  ... ${cat.items.length - 25} more`);
  }
  console.log(`\nAI edits: ${result.aiEditCount}, comments: ${result.commentCount}`);
  console.log(`Wrote ${output} (${(result.buffer.length / 1024 / 1024).toFixed(1)} MB, input ${(buffer.length / 1024 / 1024).toFixed(1)} MB) in ${((Date.now() - started) / 1000).toFixed(1)}s`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
