/**
 * Every Claude call the reviewer makes. Each is a single structured-output request (no agent
 * loop): the engine decides what text to send and applies whatever comes back itself, so the
 * model can only ever propose replacement text for paragraphs it was shown -- never touch the
 * document structure directly.
 */
import Anthropic from "@anthropic-ai/sdk";

export const REVIEW_MODEL = "claude-opus-5-5";

let client: Anthropic | null = null;

export function hasApiKey(): boolean {
  return !!process.env.ANTHROPIC_API_KEY?.trim();
}

function getClient(): Anthropic {
  if (!client) client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY, maxRetries: 3 });
  return client;
}

type Effort = "low" | "medium" | "high";

async function finalMessage(system: string, user: string, schema: Record<string, unknown>, effort: Effort) {
  const params = {
    model: REVIEW_MODEL,
    max_tokens: 64000,
    thinking: { type: "adaptive" as const },
    output_config: { effort, format: { type: "json_schema" as const, schema } },
    system,
    messages: [{ role: "user" as const, content: user }],
  };
  try {
    // Server-side refusal fallback: if a safety classifier misfires on turbine-damage
    // language, the API reroutes the request instead of returning nothing.
    return await getClient()
      .beta.messages.stream({ ...params, betas: ["server-side-fallback-2026-07-01"], fallbacks: "default" })
      .finalMessage();
  } catch (err) {
    // If this account/API version rejects the fallback option itself, run without it rather
    // than failing the whole review.
    if (err instanceof Anthropic.BadRequestError && /fallback/i.test(err.message)) {
      return await getClient().beta.messages.stream(params).finalMessage();
    }
    throw err;
  }
}

async function callJson<T>(system: string, user: string, schema: Record<string, unknown>, effort: Effort): Promise<T> {
  let message;
  try {
    message = await finalMessage(system, user, schema, effort);
  } catch (err) {
    if (err instanceof Anthropic.AuthenticationError) throw new Error("The Anthropic API key was rejected -- check ANTHROPIC_API_KEY in .env.local.");
    if (err instanceof Anthropic.RateLimitError) throw new Error("The Anthropic API rate limit was hit -- wait a minute and run the review again.");
    if (err instanceof Anthropic.APIConnectionError) throw new Error("Couldn't reach the Anthropic API -- check this computer's internet connection.");
    throw err;
  }
  if (message.stop_reason === "refusal") throw new Error("The AI declined to review this part of the report.");
  if (message.stop_reason === "max_tokens") throw new Error("The AI response was cut off (report section too long).");
  const text = message.content.map((b) => (b.type === "text" ? b.text : "")).join("");
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error("The AI returned a response that could not be read.");
  }
}

/** Runs async jobs with at most `limit` in flight. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

// --- Shared house-style rules (from ATS's Report Composition standards) ---

const ATS_CONTEXT = `You are reviewing a borescope inspection report written by a field inspector for Advanced Turbine Support (ATS), which inspects gas and steam turbines (GE, Siemens, Mitsubishi, etc.). The report goes to ATS's report reviewer, then to the customer.`;

const HOUSE_STYLE = `ATS house style (Report Composition standards):
- Observations (Observations-table Condition cells, and photo-table Observation / Comments values) are brief statements, not full sentences. The first word is capitalized; everything else is lowercase except proper nouns and abbreviations (LE, TE, IGV, VIGV, EGV, R-1, S-17, TIL, CDC, GE, DLN). No random capitalization ("Leading edge Impact Damage" -> "Leading edge impact damage").
- Multiple observations on one component are separated by commas, e.g. "Leading edge impact damage (1), Trailing edge impact damage (2), Rub marks at 6 o'clock".
- Previously identified conditions read "Previously identified <condition>, appears unchanged" (or ", has increased by ...").
- Spell out the word "Number"; never use "#".
- Measurements use a leading zero before a decimal (0.023 inch, never .023 inch) and the correct singular/plural ("1 inch", "0.5 inch", "2 inches").
- Dates never zero-pad the day: "May 7, 2017", not "May 07, 2017".
- Narrative sections (Purpose, Notes, Overall Assessment, Inspection Access, Significant Observations and Recommendations) use complete sentences with proper grammar. A recommendation for work already being done during the inspection is worded "Continue to inspect".`;

const EDIT_RULES = `Editing rules:
- Make the minimum change needed. Do not rephrase text that is already correct, and do not change meaning.
- Never add, remove, or change findings, quantities, stage/position/blade numbers, measurement values, part numbers, model numbers, TIL numbers, names, dates, or units -- only how they are written.
- Do not expand or "correct" turbine terminology or abbreviations (bucket, nozzle, shroud block, transition piece, crossfire tube, flex seal, honeycomb, tip curl, rolled metal, rub marks, R1/S1, etc.).
- Tokens like ⟦1⟧ stand for fields, links, or dropdowns. Keep every one exactly as written, in the same order.
- Keep tabs and the existing spacing between sentences as they are.
- Leave "Typical", "No defects identified", and template labels (Component, Location, Observation, Comments) unchanged.
- Leave placeholder text such as X, XX, 20XX, or "Choose an item." unchanged -- the reviewer handles those.`;

// --- 1. Proofreading ---

export interface ProofItem {
  id: string;
  context: string;
  text: string;
}
export interface ProofEdit {
  id: string;
  text: string;
  change: string;
}

const PROOF_SCHEMA = {
  type: "object",
  properties: {
    edits: {
      type: "array",
      items: {
        type: "object",
        properties: {
          id: { type: "string" },
          text: { type: "string", description: "The full corrected text of the item." },
          change: { type: "string", description: "A few words describing what was corrected." },
        },
        required: ["id", "text", "change"],
        additionalProperties: false,
      },
    },
  },
  required: ["edits"],
  additionalProperties: false,
};

export async function proofreadBatch(items: ProofItem[]): Promise<ProofEdit[]> {
  const system = `${ATS_CONTEXT}

Your job: proofread each item for spelling, grammar, punctuation, capitalization, and the house style below. Return an edit ONLY for items that need a correction, giving the item's full corrected text. Items that are fine must be left out.

${HOUSE_STYLE}

${EDIT_RULES}`;
  const user = `Proofread these report items. Each has an id, where it appears in the report, and its text.\n\n${JSON.stringify(items, null, 1)}`;
  const result = await callJson<{ edits: ProofEdit[] }>(system, user, PROOF_SCHEMA, "medium");
  const valid = new Set(items.map((i) => i.id));
  return result.edits.filter((e) => valid.has(e.id));
}

// --- 2. Overall Assessment readability ---

export interface AssessmentResult {
  edits: ProofEdit[];
  recommendations: string[];
}

const ASSESSMENT_SCHEMA = {
  type: "object",
  properties: {
    edits: PROOF_SCHEMA.properties.edits,
    recommendations: {
      type: "array",
      items: { type: "string" },
      description: "Readability or content recommendations that can't be fixed by editing text alone.",
    },
  },
  required: ["edits", "recommendations"],
  additionalProperties: false,
};

export async function reviewOverallAssessment(items: ProofItem[], supportingContext: string): Promise<AssessmentResult> {
  const system = `${ATS_CONTEXT}

Your job: review the report's Overall Assessment section and improve its readability. The Overall Assessment summarizes the condition of each inspected section (Inlet, Compressor, Combustion, Turbine, Exhaust, ...) in complete sentences with proper grammar, and lists any required follow-on maintenance.

Improve each item so it reads clearly and professionally: fix grammar and spelling, split run-on sentences, remove redundancy, use consistent sentence structure across sections, and make sure every statement is a complete sentence. Keep each item's facts exactly as they are -- do not add, drop, or soften any finding, and do not invent conditions that are not stated. Return an edit only for items you changed, with the full revised text.

Also give up to 5 short recommendations for anything that can't be fixed by rewording alone, for example: a section that was inspected (per the Observations tables) but has no assessment, an assessment that contradicts the Observations tables, or follow-on maintenance from the recommendations that isn't mentioned. Return an empty list if there is nothing worth raising.

${HOUSE_STYLE}

${EDIT_RULES}`;
  const user = `Overall Assessment items:\n${JSON.stringify(items, null, 1)}\n\nFor reference only (do not edit), the rest of the report's findings:\n${supportingContext}`;
  const result = await callJson<AssessmentResult>(system, user, ASSESSMENT_SCHEMA, "high");
  const valid = new Set(items.map((i) => i.id));
  return { edits: result.edits.filter((e) => valid.has(e.id)), recommendations: result.recommendations.slice(0, 8) };
}

// --- 3. Photo tables vs. Observations tables ---

export interface CrossCheckFinding {
  anchorId: string;
  type: "missing_photo" | "missing_observation" | "count_mismatch" | "location_mismatch" | "condition_mismatch" | "other";
  message: string;
}

const CROSSCHECK_SCHEMA = {
  type: "object",
  properties: {
    findings: {
      type: "array",
      items: {
        type: "object",
        properties: {
          anchorId: {
            type: "string",
            description: "The id of the Observations-table row (O...) or photo (P...) this finding should be attached to.",
          },
          type: {
            type: "string",
            enum: ["missing_photo", "missing_observation", "count_mismatch", "location_mismatch", "condition_mismatch", "other"],
          },
          message: { type: "string", description: "One or two sentences the reviewer can act on, naming the component and the photo(s)/row involved." },
        },
        required: ["anchorId", "type", "message"],
        additionalProperties: false,
      },
    },
  },
  required: ["findings"],
  additionalProperties: false,
};

export async function crossCheckObservations(observationRows: string, photoEntries: string): Promise<CrossCheckFinding[]> {
  const system = `${ATS_CONTEXT}

Your job: confirm that the observations listed in the report's photo tables match the report's Observations tables. ATS rules: every observation listed in an Observations table needs a photo supporting it, and every non-typical condition shown in the photos should appear in the Observations table for that component.

Compare the two lists and report real discrepancies only:
- missing_photo: an Observations-table condition (other than "No defects identified" / typical) with no photo showing it.
- missing_observation: a photo showing a defect/condition that the Observations table for that component doesn't mention (or says "No defects identified").
- count_mismatch: quantities disagree (e.g. table says impact damage (6) but photos show 4 blades).
- location_mismatch: the component, stage, position, or location (LE/TE/tip/platform, o'clock) disagrees.
- condition_mismatch: the condition itself is described differently in a way that changes meaning.
Match components sensibly: R1 / R-1 / "Stage 1 Rotor Blade" are the same; S1 / "Stage 1 Stator Vane" are the same; IGV / VIGV / "Variable Inlet Guide Vanes" are the same; combustion "Liner 3" / "Position 3" / "Combustion Liner 3" are the same. Treat wording differences that mean the same thing ("Impact damage at the leading edge" vs "Leading edge impact damage") as matches -- do not report them. Photos labeled Operational Data or Data Plate are not inspection findings; ignore them.

Attach each finding to the most relevant id: the Observations-table row (O...) when a row is involved, otherwise the photo (P...). Return an empty list if everything lines up.

The message is read by the report reviewer as a Word comment, where the O/P ids don't exist. Never write an id in the message -- refer to rows by their label ("the R4 row") and to photos by their caption ("the Stage 4 Rotor Blade, Leading Edge Pressure Side photo").`;
  const user = `OBSERVATIONS TABLES (one line per row: id | section | cells):\n${observationRows}\n\nPHOTO TABLES (one line per photo: id | section | caption fields):\n${photoEntries}`;
  const result = await callJson<{ findings: CrossCheckFinding[] }>(system, user, CROSSCHECK_SCHEMA, "high");
  return result.findings;
}

// --- 4. Heading style review ---

export interface HeadingCandidate {
  id: string;
  style: string;
  text: string;
}
export interface HeadingFix {
  id: string;
  style: "Heading 1" | "Heading 2" | "Normal";
  reason: string;
}

const HEADING_SCHEMA = {
  type: "object",
  properties: {
    fixes: {
      type: "array",
      items: {
        type: "object",
        properties: {
          id: { type: "string" },
          style: { type: "string", enum: ["Heading 1", "Heading 2", "Normal"] },
          reason: { type: "string" },
        },
        required: ["id", "style", "reason"],
        additionalProperties: false,
      },
    },
  },
  required: ["fixes"],
  additionalProperties: false,
};

export async function reviewHeadings(outline: HeadingCandidate[]): Promise<HeadingFix[]> {
  const system = `${ATS_CONTEXT}

Your job: check that every heading in the report has the correct Word Style selected, so the Table of Contents is right. ATS rules:
- Heading 1: the main report sections -- Documentation & Photographs, Overall Assessment, Inspection Details, Significant Observations and Recommendations, Observations, Photos, and any other top-level section a one-off report adds at that same level.
- Heading 2: subsections -- Purpose, Inspection Areas, Applicable TIL's, Inspection Access, and each section name under Observations and under Photos (Inlet Section, Compressor Section, Combustion Section, Turbine Section, Exhaust Section, Generator, ...).
- Normal: Notes, body text, labels, and a section name repeated at the top of a later page as a "continued" label (only the first occurrence of each subsection name is a heading).
You're given the report outline: short paragraphs outside tables with their current style. Return a fix ONLY where the current style is clearly wrong. One-off reports may have their own section names -- judge by the paragraph's role in the outline, not just its wording.`;
  const user = `Report outline (id | current style | text), in document order:\n${outline.map((o) => `${o.id} | ${o.style} | ${o.text}`).join("\n")}`;
  const result = await callJson<{ fixes: HeadingFix[] }>(system, user, HEADING_SCHEMA, "medium");
  const valid = new Set(outline.map((o) => o.id));
  return result.fixes.filter((f) => valid.has(f.id));
}
