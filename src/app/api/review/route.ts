import { NextRequest } from "next/server";
import { promises as fs } from "fs";
import path from "path";
import { runReview } from "@/lib/review/pipeline";
import { DEFAULT_OPTIONS, ProgressEvent, ReviewOptions } from "@/lib/review/types";
import { saveJob } from "@/lib/jobs";

// A full review (AI passes + Word TOC rebuild) takes a few minutes on a large report.
export const maxDuration = 900;

/** Takes the inspector's report (an uploaded file, or a path picked in the in-app file browser)
 *  and streams progress back as newline-delimited JSON, ending in a "done" event that carries
 *  the change log and a job id to download the reviewed file with. */
export async function POST(request: NextRequest) {
  const form = await request.formData();
  const file = form.get("file");
  const filePath = form.get("path");
  const optionsRaw = form.get("options");

  let input: Buffer;
  let originalName: string;
  try {
    if (file instanceof File) {
      input = Buffer.from(await file.arrayBuffer());
      originalName = file.name;
    } else if (typeof filePath === "string" && filePath) {
      input = await fs.readFile(filePath);
      originalName = path.basename(filePath);
    } else {
      return Response.json({ error: "No report was provided." }, { status: 400 });
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "Could not read the report";
    return Response.json({ error: /EBUSY|EPERM/.test(message) ? "The report is open in Word or locked -- close it and try again." : message }, { status: 400 });
  }
  if (!/\.docx$/i.test(originalName)) {
    return Response.json({ error: "Only Word .docx files can be reviewed (save older .doc files as .docx first)." }, { status: 400 });
  }

  const options: ReviewOptions = { ...DEFAULT_OPTIONS, ...(typeof optionsRaw === "string" ? JSON.parse(optionsRaw) : {}) };
  const outputFileName = originalName.replace(/\.docx$/i, "") + " - Reviewed.docx";

  const encoder = new TextEncoder();
  const stream = new ReadableStream({
    async start(controller) {
      const send = (event: ProgressEvent) => controller.enqueue(encoder.encode(JSON.stringify(event) + "\n"));
      try {
        const output = await runReview(input, options, (message) => send({ type: "step", message }));
        const jobId = saveJob(output.buffer, outputFileName);
        send({
          type: "done",
          result: {
            jobId,
            outputFileName,
            categories: output.categories,
            aiEditCount: output.aiEditCount,
            commentCount: output.commentCount,
          },
        });
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
