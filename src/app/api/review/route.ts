import { NextRequest } from "next/server";
import { promises as fs } from "fs";
import path from "path";
import { runReview } from "@/lib/review/pipeline";
import { hasApiKey, mapLimit } from "@/lib/review/ai";
import { addComparisonComments, compareReports, CompareFinding, extractFromDocx, findingsToItems } from "@/lib/review/compare";
import { DEFAULT_OPTIONS, ProgressEvent, ReviewCategory, ReviewOptions, ReviewResult } from "@/lib/review/types";
import { saveJob } from "@/lib/jobs";

// A full review (AI passes + Word TOC rebuild) takes a minute or two per report.
export const maxDuration = 1800;

/** Two reports at a time: their AI calls overlap nicely, and the Word step is queued anyway. */
const REVIEW_CONCURRENCY = 2;

interface Input {
  name: string;
  buffer: Buffer;
}

/** Takes one or more inspector reports (uploaded files and/or paths picked in the in-app file
 *  browser) and streams progress back as newline-delimited JSON, ending in a "done" event that
 *  carries each report's change log and download id -- plus, for several reports of the same
 *  unit, the cross-report comparison. */
export async function POST(request: NextRequest) {
  const form = await request.formData();
  const optionsRaw = form.get("options");

  const inputs: Input[] = [];
  try {
    for (const file of form.getAll("file")) {
      if (file instanceof File) inputs.push({ name: file.name, buffer: Buffer.from(await file.arrayBuffer()) });
    }
    for (const p of form.getAll("path")) {
      if (typeof p === "string" && p) inputs.push({ name: path.basename(p), buffer: await fs.readFile(p) });
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "Could not read the report";
    return Response.json({ error: /EBUSY|EPERM/.test(message) ? "A report is open in Word or locked -- close it and try again." : message }, { status: 400 });
  }
  if (!inputs.length) return Response.json({ error: "No report was provided." }, { status: 400 });
  const notDocx = inputs.find((i) => !/\.docx$/i.test(i.name));
  if (notDocx) {
    return Response.json({ error: `"${notDocx.name}" isn't a Word .docx file (save older .doc files as .docx first).` }, { status: 400 });
  }

  const options: ReviewOptions = { ...DEFAULT_OPTIONS, ...(typeof optionsRaw === "string" ? JSON.parse(optionsRaw) : {}) };
  const multi = inputs.length > 1;
  const letters = inputs.map((_, i) => String.fromCharCode(65 + i));
  const names = new Map(letters.map((l, i) => [l, inputs[i].name]));

  const encoder = new TextEncoder();
  const stream = new ReadableStream({
    async start(controller) {
      const send = (event: ProgressEvent) => controller.enqueue(encoder.encode(JSON.stringify(event) + "\n"));
      try {
        // Cross-report comparison runs alongside the reviews -- it only needs the original files.
        let comparison: ReviewCategory | null = null;
        let comparePromise: Promise<CompareFinding[]> | null = null;
        if (multi && options.compareReports) {
          if (!hasApiKey()) {
            comparison = { key: "crossReport", title: "Report Comparison", summary: "", items: [], skippedReason: "No Anthropic API key is set up on this computer, so the reports couldn't be compared." };
          } else {
            send({ type: "step", message: `Comparing the ${inputs.length} reports with each other...` });
            comparePromise = Promise.all(inputs.map((i) => extractFromDocx(i.buffer))).then((data) =>
              compareReports(data.map((d, i) => ({ letter: letters[i], fileName: inputs[i].name, data: d }))),
            );
            comparePromise.then(() => send({ type: "step", message: "Report comparison done." })).catch(() => {});
          }
        }

        const outputs = await mapLimit(inputs, REVIEW_CONCURRENCY, async (input) => {
          const prefix = multi ? `[${input.name}] ` : "";
          try {
            return await runReview(input.buffer, options, (message) => send({ type: "step", message: prefix + message }));
          } catch (err) {
            console.error(`review failed for ${input.name}:`, err);
            send({ type: "step", message: `${prefix}Review failed: ${err instanceof Error ? err.message : err}` });
            return err instanceof Error ? err : new Error(String(err));
          }
        });

        let findings: CompareFinding[] = [];
        if (comparePromise) {
          try {
            findings = await comparePromise;
            comparison = {
              key: "crossReport",
              title: "Report Comparison",
              summary: `Compared ${inputs.length} reports for the details, observations and significant findings that must match. Each mismatch is also a Word comment in every report involved.`,
              items: findings.length
                ? findingsToItems(findings, names)
                : [{ note: `All ${inputs.length} reports agree on the Inspection Details, the shared observations and the significant findings.` }],
            };
          } catch (err) {
            comparison = {
              key: "crossReport",
              title: "Report Comparison",
              summary: "",
              items: [],
              skippedReason: `The report comparison failed: ${err instanceof Error ? err.message : err}`,
            };
          }
        }

        const results: ReviewResult[] = [];
        for (let i = 0; i < inputs.length; i++) {
          const out = outputs[i];
          if (out instanceof Error) {
            results.push({ jobId: "", outputFileName: inputs[i].name, error: out.message, categories: [], aiEditCount: 0, commentCount: 0 });
            continue;
          }
          let buffer = out.buffer;
          let commentCount = out.commentCount;
          const categories = out.categories;
          if (findings.length) {
            const added = await addComparisonComments(buffer, letters[i], findings, names);
            buffer = added.buffer;
            commentCount += added.added;
            const mine = findings.filter((f) => f.refs.some((r) => r.startsWith(`${letters[i]}.`)));
            if (mine.length) {
              categories.unshift({
                key: "crossReport",
                title: "Cross-Report Check",
                summary: "Details and observations that must match the other reports for this unit but don't. Each one is also a Word comment in this report.",
                items: findingsToItems(mine, names),
              });
            }
          }
          results.push({
            jobId: saveJob(buffer, inputs[i].name),
            // Same name as the inspector's file, so it's ready to file without renaming (Brett's
            // request). It downloads to the browser's Downloads folder, so the original is never
            // overwritten.
            outputFileName: inputs[i].name,
            categories,
            aiEditCount: out.aiEditCount,
            commentCount,
          });
        }
        send({ type: "done", results, comparison });
      } catch (err) {
        console.error("review failed:", err);
        send({ type: "error", message: err instanceof Error ? err.message : "The review failed." });
      } finally {
        controller.close();
      }
    },
  });

  return new Response(stream, {
    headers: { "Content-Type": "application/x-ndjson; charset=utf-8", "Cache-Control": "no-cache" },
  });
}
