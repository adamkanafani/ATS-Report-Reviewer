"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { ProgressEvent, ReviewCategory, ReviewCategoryKey, ReviewItem, ReviewOptions, ReviewResult } from "@/lib/review/types";
import { DEFAULT_OPTIONS } from "@/lib/review/types";

interface BrowseResponse {
  dir: string | null;
  parent: string | null;
  folders: string[];
  files: string[];
  isDriveList: boolean;
  error?: string;
}

/** The report to review: either a file picked in the in-app browser (read from disk by the
 *  server) or one uploaded straight from the browser's own file dialog / drag-and-drop. */
type ReportSource = { kind: "path"; path: string; name: string } | { kind: "upload"; file: File; name: string };

const DRIVES_ROOT = "__DRIVES__";
const ONEDRIVE_ROOT = "__ONEDRIVE__";

/** Brett's request, item by item -- the checklist on the start screen. */
const CHECKS: { key: keyof ReviewOptions; label: string; ai?: boolean }[] = [
  { key: "formatting", label: "Correct formatting (fonts, sizes, 7\" tables, all text black)" },
  { key: "proofreading", label: "Spell check and grammar review of all text", ai: true },
  { key: "crossCheck", label: "Confirm photo-table observations match the Observations tables", ai: true },
  { key: "overallAssessment", label: "Review the Overall Assessment for readability", ai: true },
  { key: "headings", label: "Check every heading's Style selection" },
  { key: "photoTables", label: "Set all photo-table text to Times New Roman 12" },
  { key: "toc", label: "Update the Table of Contents" },
  { key: "photoResolution", label: "Set photos to 150 ppi (operational data photos excluded)" },
];

function FolderIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 20 16" fill="none" aria-hidden="true">
      <path
        d="M1 2.5C1 1.67 1.67 1 2.5 1H7.2C7.68 1 8.13 1.23 8.42 1.62L9.4 3H17.5C18.33 3 19 3.67 19 4.5V13.5C19 14.33 18.33 15 17.5 15H2.5C1.67 15 1 14.33 1 13.5V2.5Z"
        fill="#ffffff"
      />
    </svg>
  );
}

export default function Home() {
  const [sources, setSources] = useState<ReportSource[]>([]);
  // True while the picker is open to add more reports to an existing selection.
  const [adding, setAdding] = useState(false);
  // Remembers where the file picker was browsing, so reopening it resumes there.
  const [browseDir, setBrowseDir] = useState<string | null>(null);
  const [options, setOptions] = useState<ReviewOptions>(DEFAULT_OPTIONS);
  const [running, setRunning] = useState(false);
  const [steps, setSteps] = useState<string[]>([]);
  const [outcome, setOutcome] = useState<{ results: ReviewResult[]; comparison: ReviewCategory | null } | null>(null);
  const [error, setError] = useState<string | null>(null);

  function reset() {
    setSources([]);
    setAdding(false);
    setRunning(false);
    setSteps([]);
    setOutcome(null);
    setError(null);
  }

  function addSources(more: ReportSource[]) {
    setSources((prev) => {
      const key = (s: ReportSource) => (s.kind === "path" ? s.path : `upload:${s.name}:${s.file.size}`);
      const seen = new Set(prev.map(key));
      return [...prev, ...more.filter((s) => !seen.has(key(s)))];
    });
    setAdding(false);
  }

  async function runReview() {
    if (!sources.length) return;
    setRunning(true);
    setSteps([]);
    setError(null);
    setOutcome(null);
    try {
      const form = new FormData();
      for (const source of sources) {
        if (source.kind === "upload") form.append("file", source.file);
        else form.append("path", source.path);
      }
      form.set("options", JSON.stringify(options));
      const res = await fetch("/api/review", { method: "POST", body: form });
      if (!res.ok || !res.body) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || "The review failed to start.");
      }
      const reader = res.body.getReader();
      const decoder = new TextDecoder();
      let buffered = "";
      let finished = false;
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        buffered += decoder.decode(value, { stream: true });
        const lines = buffered.split("\n");
        buffered = lines.pop() ?? "";
        for (const line of lines) {
          if (!line.trim()) continue;
          const event = JSON.parse(line) as ProgressEvent;
          if (event.type === "step") setSteps((prev) => [...prev, event.message]);
          else if (event.type === "done") {
            finished = true;
            setOutcome({ results: event.results, comparison: event.comparison });
          } else if (event.type === "error") throw new Error(event.message);
        }
      }
      if (!finished) throw new Error("The review stopped before finishing -- is the app's server window still open?");
    } catch (err) {
      setError(err instanceof Error ? err.message : "The review failed.");
    } finally {
      setRunning(false);
    }
  }

  if (outcome && sources.length) {
    return <ResultsScreen sources={sources} results={outcome.results} comparison={outcome.comparison} onReviewAnother={reset} />;
  }

  if (running || (error && steps.length > 0)) {
    return <ProgressScreen sources={sources} steps={steps} error={error} onRetry={runReview} onChangeFile={reset} running={running} />;
  }

  if (!sources.length || adding) {
    return (
      <FilePicker
        title={adding ? "Add more reports" : "Select the inspector's report"}
        instructions={
          adding
            ? "Tick the other reports for this unit (borescope, PT, ECI, UT...), or drop them here."
            : "Pick the Word report (.docx) exactly as it came from the inspector, or drop it here. To review several reports for the same unit together -- and check they match -- tick each one. The app makes corrected copies; the originals are never changed."
        }
        extensions="docx"
        onChosen={(paths) => addSources(paths.map((p) => ({ kind: "path" as const, path: p, name: p.split(/[\\/]/).pop() ?? p })))}
        onUpload={(files) => addSources(files.map((file) => ({ kind: "upload" as const, file, name: file.name })))}
        onBack={adding ? () => setAdding(false) : undefined}
        initialDir={browseDir}
        onDirChange={setBrowseDir}
      />
    );
  }

  return (
    <ReadyScreen
      sources={sources}
      options={options}
      onToggle={(key) => setOptions((prev) => ({ ...prev, [key]: !prev[key] }))}
      error={error}
      onReview={runReview}
      onAddMore={() => setAdding(true)}
      onRemove={(i) => setSources((prev) => prev.filter((_, j) => j !== i))}
      onChangeFile={reset}
    />
  );
}

function ReadyScreen({
  sources,
  options,
  onToggle,
  error,
  onReview,
  onAddMore,
  onRemove,
  onChangeFile,
}: {
  sources: ReportSource[];
  options: ReviewOptions;
  onToggle: (key: keyof ReviewOptions) => void;
  error: string | null;
  onReview: () => void;
  onAddMore: () => void;
  onRemove: (index: number) => void;
  onChangeFile: () => void;
}) {
  const [hasApiKey, setHasApiKey] = useState<boolean | null>(null);
  useEffect(() => {
    fetch("/api/status")
      .then((res) => res.json())
      .then((json) => setHasApiKey(!!json.hasApiKey))
      .catch(() => setHasApiKey(null));
  }, []);
  const anyChecked = CHECKS.some((c) => options[c.key]);
  const multi = sources.length > 1;

  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>Ready to review</h1>
        </div>
        <p className="app-version">Version {process.env.NEXT_PUBLIC_APP_VERSION}</p>
        <ul className="source-list">
          {sources.map((source, i) => (
            <li key={i}>
              <span className="job-path">{source.kind === "path" ? source.path : source.name}</span>
              {multi && (
                <button type="button" className="unit-row-delete" title="Remove this report" aria-label="Remove this report" onClick={() => onRemove(i)}>
                  ×
                </button>
              )}
            </li>
          ))}
        </ul>
        <button type="button" className="secondary add-more" onClick={onAddMore}>
          + Add another report for this unit
        </button>
        <p className="folder-picker-subtitle" style={{ marginTop: 14 }}>
          {multi ? "Each reviewed copy downloads under its own file name" : "The reviewed copy downloads under the same file name"}, into your
          Downloads folder. Everything the AI changed is in <span className="red-sample">red</span>; everything else is black. Findings that
          need a decision (mismatched photos, Overall Assessment suggestions{multi ? ", reports that don't match each other" : ""}) are left as
          Word comments.
        </p>
        <div className="modal-checklist">
          {multi && (
            <label className="modal-checkbox-row compare-row">
              <input type="checkbox" checked={options.compareReports} onChange={() => onToggle("compareReports")} />
              Compare these {sources.length} reports with each other (Inspection Details, shared observations, significant findings)
              <span className="ai-tag">AI</span>
            </label>
          )}
          {CHECKS.map((check) => (
            <label key={check.key} className="modal-checkbox-row">
              <input type="checkbox" checked={options[check.key]} onChange={() => onToggle(check.key)} />
              {check.label}
              {check.ai && <span className="ai-tag">AI</span>}
            </label>
          ))}
        </div>
        {hasApiKey === false && (
          <p className="folder-picker-warning">
            No Anthropic API key is set up on this computer, so the checks marked AI will be skipped. See the setup guide to add one.
          </p>
        )}
        {error && <p className="folder-picker-error">{error}</p>}
        <div className="modal-actions" style={{ justifyContent: "space-between" }}>
          <button className="secondary" onClick={onChangeFile}>
            {multi ? "Start Over" : "Change File"}
          </button>
          <button className="generate-report" disabled={!anyChecked} onClick={onReview}>
            {multi ? `Review ${sources.length} Reports` : "Review Report"}
          </button>
        </div>
      </div>
    </div>
  );
}

function ProgressScreen({
  sources,
  steps,
  error,
  running,
  onRetry,
  onChangeFile,
}: {
  sources: ReportSource[];
  steps: string[];
  error: string | null;
  running: boolean;
  onRetry: () => void;
  onChangeFile: () => void;
}) {
  const multi = sources.length > 1;
  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>{error ? "Review failed" : multi ? `Reviewing ${sources.length} reports...` : "Reviewing..."}</h1>
        </div>
        {sources.map((s, i) => (
          <p key={i} className="job-path">
            {s.name}
          </p>
        ))}
        <p className="folder-picker-subtitle" style={{ marginTop: 14 }}>
          {error
            ? error
            : multi
              ? "This takes a minute or two per report. Leave this window open -- the reviewed files download from the next screen."
              : "A large report takes a few minutes. Leave this window open -- the reviewed file downloads from the next screen."}
        </p>
        <ol className="progress-list">
          {steps.map((step, i) => (
            <li key={i} className={i === steps.length - 1 && running ? "current" : "done"}>
              {step}
            </li>
          ))}
        </ol>
        {!running && error && (
          <div className="modal-actions" style={{ justifyContent: "space-between" }}>
            <button className="secondary" onClick={onChangeFile}>
              Choose different files
            </button>
            <button onClick={onRetry}>Try again</button>
          </div>
        )}
      </div>
    </div>
  );
}

function downloadJob(result: ReviewResult) {
  const a = document.createElement("a");
  a.href = `/api/download?jobId=${encodeURIComponent(result.jobId)}`;
  a.download = result.outputFileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function ResultsScreen({
  sources,
  results,
  comparison,
  onReviewAnother,
}: {
  sources: ReportSource[];
  results: ReviewResult[];
  comparison: ReviewCategory | null;
  onReviewAnother: () => void;
}) {
  // Tab per report, plus the comparison when several reports were checked against each other.
  // Opens on the comparison, since that's the reason for reviewing them together.
  const [tab, setTab] = useState<number | "comparison">(comparison ? "comparison" : 0);
  const ok = results.filter((r) => !r.error);

  // Start the downloads right away -- the files are what Brett came for; the change log is
  // extra. Spaced out so the browser doesn't drop any (it may ask once to allow multiple
  // downloads from this site).
  const started = useRef(false);
  useEffect(() => {
    if (started.current) return;
    started.current = true;
    ok.forEach((r, i) => setTimeout(() => downloadJob(r), i * 700));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const current = tab === "comparison" ? null : results[tab];
  const headerPath =
    results.length === 1 ? (sources[0]?.kind === "path" ? sources[0].path : sources[0]?.name) : `${results.length} reports reviewed together`;

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1 style={{ fontSize: 16, fontWeight: 700, color: "#ffffff" }}>ATS Report Reviewer</h1>
        <span className="app-version">v {process.env.NEXT_PUBLIC_APP_VERSION}</span>
        <span className="job-path">{headerPath}</span>
        <button className="secondary" onClick={onReviewAnother}>
          Review Another Report
        </button>
      </header>
      {(results.length > 1 || comparison) && (
        <nav className="report-tabs">
          {comparison && (
            <button type="button" className={`report-tab ${tab === "comparison" ? "active" : ""}`} onClick={() => setTab("comparison")}>
              Report Comparison
              {comparison.items.some((i) => i.level === "attention") && <span className="status-badge needs-attention">Review</span>}
            </button>
          )}
          {results.map((r, i) => (
            <button key={i} type="button" className={`report-tab ${tab === i ? "active" : ""}`} onClick={() => setTab(i)}>
              {r.outputFileName}
              {r.error && <span className="status-badge missing">Failed</span>}
            </button>
          ))}
        </nav>
      )}
      {current ? (
        current.error ? (
          <div className="app-body">
            <main className="main-panel">
              <h2 style={{ marginBottom: 8 }}>{current.outputFileName}</h2>
              <p className="section-reason">This report couldn&apos;t be reviewed: {current.error}</p>
            </main>
          </div>
        ) : (
          <ReportView key={tab} result={current} />
        )
      ) : (
        comparison && (
          <div className="app-body">
            <main className="main-panel">
              <CategoryPanel category={comparison} />
            </main>
          </div>
        )
      )}
      <footer className="app-header">
        {current && !current.error ? (
          <>
            <span className="footer-stats">
              {current.aiEditCount} AI edit{current.aiEditCount === 1 ? "" : "s"} in red · {current.commentCount} comment
              {current.commentCount === 1 ? "" : "s"} for review
            </span>
            {/* The report already downloaded automatically when this screen opened -- confirm that
                instead of offering a second big download button. The link is only for a download
                the browser blocked or that got lost. */}
            <span className="download-status">
              <span>
                ✓ Saved to your Downloads folder: <strong>{current.outputFileName}</strong>
              </span>
              <button type="button" className="download-again" onClick={() => downloadJob(current)}>
                Download again
              </button>
            </span>
          </>
        ) : (
          <span className="download-status">
            <span>
              ✓ {ok.length} reviewed report{ok.length === 1 ? "" : "s"} saved to your Downloads folder
            </span>
            {ok.length > 0 && (
              <button type="button" className="download-again" onClick={() => ok.forEach((r, i) => setTimeout(() => downloadJob(r), i * 700))}>
                Download all again
              </button>
            )}
          </span>
        )}
      </footer>
    </div>
  );
}

/** One report's change log: category sidebar + item list. */
function ReportView({ result }: { result: ReviewResult }) {
  // Open on whatever most needs Brett's attention, rather than on a step that was skipped.
  const [selected, setSelected] = useState<ReviewCategoryKey>(
    () =>
      (
        result.categories.find((c) => c.items.some((i) => i.level === "attention")) ??
        result.categories.find((c) => !c.skippedReason) ??
        result.categories[0]
      )?.key ?? "formatting",
  );
  const category = result.categories.find((c) => c.key === selected) ?? result.categories[0];

  return (
    <div className="app-body">
      <nav className="sidebar">
        <div className="sidebar-scroll">
          {result.categories.map((cat) => {
            const attention = cat.items.filter((i) => i.level === "attention").length;
            return (
              <button key={cat.key} className={`section-item ${selected === cat.key ? "active" : ""}`} onClick={() => setSelected(cat.key)}>
                <div className="title-row">
                  <span>{cat.title}</span>
                  {cat.skippedReason ? (
                    <span className="status-badge missing">Skipped</span>
                  ) : attention > 0 ? (
                    <span className="status-badge needs-attention">Review</span>
                  ) : (
                    <span className="status-badge ready">Done</span>
                  )}
                </div>
                <div className="confidence-tag">{categoryCountLabel(cat)}</div>
              </button>
            );
          })}
        </div>
      </nav>
      <main className="main-panel">{category && <CategoryPanel category={category} />}</main>
    </div>
  );
}

function categoryCountLabel(cat: ReviewCategory): string {
  if (cat.skippedReason && cat.items.length === 0) return "Not run";
  const n = cat.items.length;
  switch (cat.key) {
    case "proofreading":
    case "overallAssessment":
      return `${n} change${n === 1 ? "" : "s"}`;
    case "crossReport":
    case "crossCheck":
    case "placeholders":
      return `${cat.items.filter((i) => i.level === "attention").length} to check`;
    case "headings":
      return `${n} style${n === 1 ? "" : "s"} fixed`;
    default:
      return `${n} item${n === 1 ? "" : "s"}`;
  }
}

function CategoryPanel({ category }: { category: ReviewCategory }) {
  return (
    <>
      <h2 style={{ marginBottom: 8 }}>{category.title}</h2>
      <p className="section-reason">{category.skippedReason ?? category.summary}</p>
      {category.items.length === 0 ? (
        !category.skippedReason && <p className="empty-note">Nothing needed changing here.</p>
      ) : (
        <ul className="change-list">
          {category.items.map((item, i) => (
            <ChangeCard key={i} item={item} />
          ))}
        </ul>
      )}
    </>
  );
}

function ChangeCard({ item }: { item: ReviewItem }) {
  const showDiff = item.before !== undefined && item.after !== undefined;
  return (
    <li className={`change-card ${item.level === "attention" ? "attention" : ""}`}>
      {item.location && <div className="change-location">{item.location}</div>}
      <div className="change-note">{item.note}</div>
      {showDiff ? (
        <WordDiff before={item.before!} after={item.after!} />
      ) : (
        item.before && <div className="change-quote">{item.before}</div>
      )}
    </li>
  );
}

/** Word-level before/after view: removed words struck through, added words in red -- the same
 *  red Brett will see in the Word file. */
function WordDiff({ before, after }: { before: string; after: string }) {
  const ops = useMemo(() => diffWords(before, after), [before, after]);
  if (!/[a-z]/i.test(before + after) || before === after) {
    return (
      <div className="change-diff">
        <span>{before}</span> → <span className="ins">{after}</span>
      </div>
    );
  }
  return (
    <div className="change-diff">
      {ops.map((op, i) =>
        op.op === "eq" ? (
          <span key={i}>{op.tok}</span>
        ) : op.op === "del" ? (
          <span key={i} className="del">
            {op.tok}
          </span>
        ) : (
          <span key={i} className="ins">
            {op.tok}
          </span>
        ),
      )}
    </div>
  );
}

function diffWords(a: string, b: string): { op: "eq" | "del" | "ins"; tok: string }[] {
  const ta = a.match(/\S+|\s+/g) ?? [];
  const tb = b.match(/\S+|\s+/g) ?? [];
  const n = ta.length;
  const m = tb.length;
  if (n * m > 250_000) return [{ op: "del", tok: a }, { op: "ins", tok: ` ${b}` }];
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) dp[i][j] = ta[i] === tb[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const out: { op: "eq" | "del" | "ins"; tok: string }[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (ta[i] === tb[j]) {
      out.push({ op: "eq", tok: ta[i++] });
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) out.push({ op: "del", tok: ta[i++] });
    else out.push({ op: "ins", tok: tb[j++] });
  }
  while (i < n) out.push({ op: "del", tok: ta[i++] });
  while (j < m) out.push({ op: "ins", tok: tb[j++] });
  return out;
}

function FilePicker({
  title,
  instructions,
  extensions,
  onChosen,
  onUpload,
  onBack,
  initialDir,
  onDirChange,
}: {
  title: string;
  instructions: string;
  extensions: string;
  /** One file clicked, or several ticked (possibly across folders). */
  onChosen: (paths: string[]) => void;
  /** Files from the browser's own file dialog or a drag-and-drop, instead of the in-app browser. */
  onUpload: (files: File[]) => void;
  onBack?: () => void;
  initialDir?: string | null;
  onDirChange?: (dir: string | null) => void;
}) {
  // Directory-visit history (like a browser's back/forward), separate from `data.parent`
  // ("Up one level", which walks toward the filesystem root, not visit order).
  const [history, setHistory] = useState<(string | null)[]>([initialDir ?? null]);
  const [historyIndex, setHistoryIndex] = useState(0);
  const dir = history[historyIndex];
  const [data, setData] = useState<BrowseResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  // Ticked files (full paths) -- kept while browsing other folders, so a unit's reports can be
  // picked from wherever they live.
  const [ticked, setTicked] = useState<string[]>([]);
  const uploadInputRef = useRef<HTMLInputElement>(null);

  function toggleTicked(full: string) {
    setTicked((prev) => (prev.includes(full) ? prev.filter((p) => p !== full) : [...prev, full]));
  }

  function navigateTo(next: string | null) {
    const truncated = history.slice(0, historyIndex + 1);
    setHistory([...truncated, next]);
    setHistoryIndex(truncated.length);
  }

  function acceptFiles(list: FileList | null | undefined) {
    const files = Array.from(list ?? []);
    if (!files.length) return;
    const docx = files.filter((f) => /\.docx$/i.test(f.name));
    if (!docx.length) {
      setError("Only Word .docx files can be reviewed.");
      return;
    }
    onUpload(docx);
  }


  useEffect(() => {
    onDirChange?.(dir);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dir]);

  useEffect(() => {
    let cancelled = false;
    // Resetting loading/error synchronously here (React's own documented pattern for
    // effect-based data fetching) so the picker shows a loading state immediately when `dir`
    // changes, rather than briefly showing the previous directory's stale listing.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setLoading(true);
    setError(null);
    const params = new URLSearchParams({ ext: extensions });
    if (dir) params.set("dir", dir);
    fetch(`/api/browse?${params.toString()}`)
      .then((res) => res.json())
      .then((json: BrowseResponse) => {
        if (cancelled) return;
        if (json.error) setError(json.error);
        else setData(json);
      })
      .catch((err) => {
        if (!cancelled) setError(err instanceof Error ? err.message : "Failed to browse");
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [dir, extensions]);

  const joinPath = (base: string, name: string) => (base.endsWith("\\") ? `${base}${name}` : `${base}\\${name}`);

  return (
    <div
      className="folder-picker"
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={(e) => {
        if (e.currentTarget === e.target) setDragging(false);
      }}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        acceptFiles(e.dataTransfer.files);
      }}
    >
      <div className={`folder-picker-card ${dragging ? "dragging" : ""}`}>
        <div className="app-logo">
          <h1>{title}</h1>
        </div>
        <p className="folder-picker-subtitle">{instructions}</p>
        <p className="app-version">Version {process.env.NEXT_PUBLIC_APP_VERSION}</p>
        {data && !data.isDriveList && data.dir && <p className="job-path">{data.dir}</p>}
        <div className="toolbar">
          {onBack && (
            <button className="secondary" onClick={onBack}>
              Back
            </button>
          )}
          <button className="secondary" disabled={historyIndex === 0} onClick={() => setHistoryIndex((i) => i - 1)} title="Previous folder" aria-label="Previous folder">
            ◀
          </button>
          <button
            className="secondary"
            disabled={historyIndex === history.length - 1}
            onClick={() => setHistoryIndex((i) => i + 1)}
            title="Next folder"
            aria-label="Next folder"
          >
            ▶
          </button>
          <button className="secondary" onClick={() => navigateTo(DRIVES_ROOT)}>
            This PC
          </button>
          <button className="secondary" onClick={() => navigateTo(ONEDRIVE_ROOT)}>
            OneDrive
          </button>
          {data?.parent && (
            <button className="secondary" onClick={() => navigateTo(data.parent)}>
              Up one level
            </button>
          )}
          <input
            ref={uploadInputRef}
            type="file"
            accept=".docx"
            multiple
            hidden
            onChange={(e) => {
              acceptFiles(e.target.files);
              e.target.value = "";
            }}
          />
          <button onClick={() => uploadInputRef.current?.click()}>Upload files...</button>
        </div>
        {error && <p className="folder-picker-error">{error}</p>}
        {loading ? (
          <div className="folder-list-loading">Loading...</div>
        ) : (
          <div className="folder-list folder-list-enter">
            {data && data.folders.length === 0 && data.files.length === 0 && <div className="folder-list-empty">Nothing here.</div>}
            {data?.folders.map((name) => {
              const full = data.isDriveList ? name : joinPath(data.dir as string, name);
              return (
                <button key={full} onClick={() => navigateTo(full)}>
                  <FolderIcon /> {name}
                </button>
              );
            })}
            {data?.files
              .filter((name) => !name.startsWith("~$"))
              .map((name) => {
                const full = joinPath(data.dir as string, name);
                const isTicked = ticked.includes(full);
                return (
                  <div key={full} className={`file-row ${isTicked ? "ticked" : ""}`}>
                    <input
                      type="checkbox"
                      checked={isTicked}
                      onChange={() => toggleTicked(full)}
                      title="Tick to review several reports together"
                      aria-label={`Select ${name}`}
                    />
                    {/* With nothing ticked, clicking a report reviews just that one (the common
                        case); once anything is ticked, clicking toggles it instead. */}
                    <button onClick={() => (ticked.length ? toggleTicked(full) : onChosen([full]))}>📄 {name}</button>
                  </div>
                );
              })}
          </div>
        )}
        {ticked.length > 0 && (
          <div className="ticked-bar">
            <span>
              {ticked.length} report{ticked.length === 1 ? "" : "s"} selected
              <button type="button" className="link-button" onClick={() => setTicked([])}>
                Clear
              </button>
            </span>
            <button className="generate-report" onClick={() => onChosen(ticked)}>
              {ticked.length === 1 ? "Continue with 1 report" : `Review ${ticked.length} reports together`}
            </button>
          </div>
        )}
        {dragging && <div className="drop-hint">Drop the report(s) to review them</div>}
      </div>
    </div>
  );
}
