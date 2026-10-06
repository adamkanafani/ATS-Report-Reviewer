import { hasApiKey } from "@/lib/review/ai";

/** Lets the start screen warn up front when the AI checks can't run on this computer. */
export async function GET() {
  return Response.json({ hasApiKey: hasApiKey() });
}
