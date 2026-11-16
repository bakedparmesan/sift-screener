// Sift: identify which spreadsheet columns hold the title, abstract and screening decisions.
// The Claude API key lives only in this function's environment (Supabase secret ANTHROPIC_API_KEY).
// It does one fixed job, so it can't be used as a general-purpose Claude proxy.
import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import Anthropic from "npm:@anthropic-ai/sdk@1.3.0";
import { jsonSchemaOutputFormat } from "npm:@anthropic-ai/sdk@1.3.0/helpers/json-schema";

const ALLOWED_ORIGINS = new Set([
  "https://bakedparmesan.github.io",
  "http://127.0.0.1:8765",
  "http://localhost:8765",
]);
const MAX_COLUMNS = 80;
const MAX_SAMPLES = 4;
const MAX_NAME = 200;
const MAX_SAMPLE = 240;
const MAX_BODY_BYTES = 250_000; // older clients may still send every sheet

// Light per-IP rate limit (per function instance): 15 requests a minute.
const hits = new Map<string, number[]>();
function rateLimited(ip: string): boolean {
  const now = Date.now();
  const recent = (hits.get(ip) || []).filter((t) => now - t < 60_000);
  recent.push(now);
  hits.set(ip, recent);
  if (hits.size > 5000) hits.clear();
  return recent.length > 15;
}

function cors(origin: string | null): Record<string, string> {
  return {
    "Access-Control-Allow-Origin": origin && ALLOWED_ORIGINS.has(origin) ? origin : "https://bakedparmesan.github.io",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
    "Vary": "Origin",
  };
}
// Never let a key reach the logs, whatever an error message contains.
const redact = (v: unknown) => String(v).replace(/sk-ant-[A-Za-z0-9_\-\s]+/g, "sk-ant-[redacted]").slice(0, 300);
const apiKey = () => (Deno.env.get("ANTHROPIC_API_KEY") || "").replace(/\s+/g, "");

function reply(status: number, body: unknown, origin: string | null) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...cors(origin), "Content-Type": "application/json" },
  });
}

const SYSTEM = `You help researchers set up title and abstract screening for a systematic review.
You get the columns of one sheet of search results: each column's position, its header, and a few sample values.
Pick the column positions for:
- title: the article or record title.
- abstract: the abstract or summary text.
- title_decision: an existing column that records title screening decisions (include/exclude, 1/0, yes/no). Use -1 if there isn't one.
- abstract_decision: an existing column that records abstract screening decisions. Use -1 if there isn't one.
Judge by both the header and the sample values: titles are one sentence long, abstracts are paragraphs.
Never pick the same column twice. Use -1 for title or abstract only if no column fits.
In note, write one short plain sentence about anything the user should double-check, or an empty string. Refer to columns by their headers, never by position numbers.`;

const OUTPUT = jsonSchemaOutputFormat({
  type: "object",
  properties: {
    title: { type: "integer" },
    abstract: { type: "integer" },
    title_decision: { type: "integer" },
    abstract_decision: { type: "integer" },
    note: { type: "string" },
  },
  required: ["title", "abstract", "title_decision", "abstract_decision", "note"],
  additionalProperties: false,
} as const);

type Column = { name: string; samples: string[] };

function cleanColumnList(raw: unknown): Column[] | null {
  if (!Array.isArray(raw) || raw.length > MAX_COLUMNS) return null;
  const out: Column[] = [];
  for (const c of raw) {
    if (!c || typeof c !== "object") return null;
    const name = String((c as { name?: unknown }).name ?? "").slice(0, MAX_NAME);
    const samples = (c as { samples?: unknown }).samples;
    out.push({ name, samples: Array.isArray(samples) ? samples.slice(0, MAX_SAMPLES).map((v) => String(v ?? "").slice(0, MAX_SAMPLE)) : [] });
  }
  return out;
}

// Accepts { columns } (one sheet). Older clients sent { sheets: [...] }; use the first sheet then.
function cleanInput(input: unknown): Column[] | null {
  if (!input || typeof input !== "object") return null;
  const body = input as { columns?: unknown; sheets?: unknown };
  const raw = Array.isArray(body.columns) ? body.columns
    : Array.isArray(body.sheets) && body.sheets[0] && typeof body.sheets[0] === "object" ? (body.sheets[0] as { columns?: unknown }).columns
    : null;
  const columns = cleanColumnList(raw);
  return columns && columns.length ? columns : null;
}

Deno.serve(async (req: Request) => {
  const origin = req.headers.get("origin");
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors(origin) });
  if (req.method !== "POST") return reply(405, { error: "Use POST." }, origin);
  if (!origin || !ALLOWED_ORIGINS.has(origin)) return reply(403, { error: "This service only works from Sift." }, origin);

  const ip = (req.headers.get("x-forwarded-for") || "").split(",")[0].trim() || "unknown";
  if (rateLimited(ip)) return reply(429, { error: "Too many requests. Wait a minute and try again." }, origin);

  const text = await req.text();
  if (text.length > MAX_BODY_BYTES) return reply(413, { error: "Too much data sent." }, origin);
  let columns: Column[] | null = null;
  try { columns = cleanInput(JSON.parse(text)); } catch { columns = null; }
  if (!columns) return reply(400, { error: "Send the spreadsheet's columns." }, origin);

  if (!apiKey()) {
    console.error("ANTHROPIC_API_KEY secret is not set");
    return reply(503, { error: "The column finder isn’t set up yet. Pick the columns yourself for now." }, origin);
  }

  const listing = columns.map((c, i) => ({ position: i, header: c.name || "(no header)", samples: c.samples }));

  try {
    const client = new Anthropic({ apiKey: apiKey() }); // key comes from the function's secrets, whitespace removed
    const response = await client.messages.parse({
      model: "claude-haiku-5-5",
      max_tokens: 4000,
      output_config: { effort: "low", format: OUTPUT },
      system: SYSTEM,
      messages: [{ role: "user", content: JSON.stringify(listing) }],
    });
    if (response.stop_reason === "refusal") return reply(502, { error: "Claude couldn’t identify the columns for this file." }, origin);
    const out = response.parsed_output;
    if (!out) return reply(502, { error: "Claude’s answer couldn’t be read. Try again." }, origin);
    const ok = (n: number) => Number.isInteger(n) && n >= -1 && n < columns!.length;
    if (![out.title, out.abstract, out.title_decision, out.abstract_decision].every(ok)) {
      return reply(502, { error: "Claude picked a column that doesn’t exist. Try again." }, origin);
    }
    // Users know columns by their headers, not by Claude's position numbers.
    const note = out.note.replace(/\b(column|position)\s+(\d+)\b/gi, (m: string, _w: string, n: string, at: number, all: string) => {
      const c = columns![Number(n)];
      if (!c || !c.name) return m;
      const startOfSentence = at === 0 || /[.!?]\s*$/.test(all.slice(0, at));
      return `${startOfSentence ? "The" : "the"} “${c.name}” column`;
    });
    return reply(200, { ...out, note }, origin);
  } catch (e) {
    if (e instanceof Anthropic.AuthenticationError || e instanceof Anthropic.PermissionDeniedError) {
      console.error("Claude key rejected", e.status);
      return reply(500, { error: "The column finder isn’t set up correctly. Pick the columns yourself for now." }, origin);
    }
    if (e instanceof Anthropic.RateLimitError) return reply(429, { error: "Claude is busy. Try again in a moment." }, origin);
    if (e instanceof Anthropic.APIConnectionError) return reply(502, { error: "Couldn’t reach Claude. Try again." }, origin);
    if (e instanceof Anthropic.APIError) {
      console.error("Claude API error", e.status, redact(e.message));
      return reply(502, { error: "Claude couldn’t answer just now. Try again." }, origin);
    }
    console.error("Unexpected error", e instanceof Error ? e.name : typeof e, redact(e instanceof Error ? e.message : e));
    return reply(500, { error: "Something went wrong. Pick the columns yourself for now." }, origin);
  }
});
