import { randomUUID } from "crypto";

/** Finished reviews, held in memory just long enough to download -- the reviewed file is the
 *  only thing worth keeping, and it lands in the user's Downloads folder. */
interface Job {
  buffer: Buffer;
  fileName: string;
  createdAt: number;
}

const JOB_TTL_MS = 2 * 60 * 60 * 1000;
const jobs = new Map<string, Job>();

export function saveJob(buffer: Buffer, fileName: string): string {
  const id = randomUUID();
  jobs.set(id, { buffer, fileName, createdAt: Date.now() });
  return id;
}

export function getJob(id: string): Job | undefined {
  return jobs.get(id);
}

declare global {
  var __atsReviewerJobSweepStarted: boolean | undefined;
}

if (!globalThis.__atsReviewerJobSweepStarted) {
  globalThis.__atsReviewerJobSweepStarted = true;
  setInterval(() => {
    const cutoff = Date.now() - JOB_TTL_MS;
    for (const [id, job] of jobs) if (job.createdAt < cutoff) jobs.delete(id);
  }, 10 * 60 * 1000).unref?.();
}
