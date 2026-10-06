import { NextRequest } from "next/server";
import { getJob } from "@/lib/jobs";

export async function GET(request: NextRequest) {
  const jobId = request.nextUrl.searchParams.get("jobId");
  const job = jobId ? getJob(jobId) : undefined;
  if (!job) {
    return Response.json({ error: "This reviewed report is no longer available -- please run the review again." }, { status: 410 });
  }
  return new Response(new Uint8Array(job.buffer), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="reviewed-report.docx"; filename*=UTF-8''${encodeURIComponent(job.fileName)}`,
    },
  });
}
