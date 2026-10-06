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
  const [source, setSource] = useState<ReportSource | null>(null);
  // Remembers where the file picker was browsing, so "Change File" reopens it there.
  const [browseDir, setBrowseDir] = useState<string | null>(null);
  const [options, setOptions] = useState<ReviewOptions>(DEFAULT_OPTIONS);
  const [running, setRunning] = useState(false);
  const [steps, setSteps] = useState<string[]>([]);
  const [result, setResult] = useState<ReviewResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  function reset() {
    setSource(null);
    setRunning(false);
    setSteps([]);
    setResult(null);
    setError(null);
  }

  async function runReview() {
    if (!source) return;
    setRunning(true);
    setSteps([]);
    setError(null);
    setResult(null);
    try {
      const form = new FormData();
      if (source.kind === "upload") form.set("file", source.file);
      else form.set("path", source.path);
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
            setResult(event.result);
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

  if (result && source) {
    return <ResultsScreen source={source} result={result} onReviewAnother={reset} />;
  }

  if (running || (error && steps.length > 0)) {
    return <ProgressScreen source={source} steps={steps} error={error} onRetry={runReview} onChangeFile={reset} running={running} />;
  }

  if (!source) {
    return (
      <FilePicker
        title="Select the inspector's report"
        instructions="Pick the Word report (.docx) exactly as it came from the inspector, or drop it here. The app makes a corrected copy -- the original file is never changed."
        extensions="docx"
        onChosen={(p) => setSource({ kind: "path", path: p, name: p.split(/[\\/]/).pop() ?? p })}
        onUpload={(file) => setSource({ kind: "upload", file, name: file.name })}
        initialDir={browseDir}
        onDirChange={setBrowseDir}
      />
    );
  }

  return (
    <ReadyScreen
      source={source}
      options={options}
      onToggle={(key) => setOptions((prev) => ({ ...prev, [key]: !prev[key] }))}
      error={error}
      onReview={runReview}
      onChangeFile={reset}
    />
  );
}

function ReadyScreen({
  source,
  options,
  onToggle,
  error,
  onReview,
  onChangeFile,
}: {
  source: ReportSource;
  options: ReviewOptions;
  onToggle: (key: keyof ReviewOptions) => void;
  error: string | null;
  onReview: () => void;
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

  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>Ready to review</h1>
        </div>
        <p className="job-path">{source.kind === "path" ? source.path : source.name}</p>
        <p className="folder-picker-subtitle" style={{ marginTop: 14 }}>
          The reviewed copy comes back as a new Word file. Everything the AI changed is in <span className="red-sample">red</span>;
          everything else is black. Findings that need a decision (mismatched photos, Overall Assessment suggestions) are left as Word
          comments.
        </p>
        <div className="modal-checklist">
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
            Change File
          </button>
          <button className="generate-report" disabled={!anyChecked} onClick={onReview}>
            Review Report
          </button>
        </div>
      </div>
    </div>
  );
}

function ProgressScreen({
  source,
  steps,
  error,
  running,
  onRetry,
  onChangeFile,
}: {
  source: ReportSource | null;
  steps: string[];
  error: string | null;
  running: boolean;
  onRetry: () => void;
  onChangeFile: () => void;
}) {
  return (
    <div className="folder-picker">
      <div className="folder-picker-card">
        <div className="app-logo">
          <h1>{error ? "Review failed" : "Reviewing..."}</h1>
        </div>
        {source && <p className="job-path">{source.name}</p>}
        <p className="folder-picker-subtitle" style={{ marginTop: 14 }}>
          {error ? error : "A large report takes a few minutes. Leave this window open -- the reviewed file downloads from the next screen."}
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
              Choose a different file
            </button>
            <button onClick={onRetry}>Try again</button>
          </div>
        )}
      </div>
    </div>
  );
}

function ResultsScreen({ source, result, onReviewAnother }: { source: ReportSource; result: ReviewResult; onReviewAnother: () => void }) {
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

  function download() {
    const a = document.createElement("a");
    a.href = `/api/download?jobId=${encodeURIComponent(result.jobId)}`;
    a.download = result.outputFileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  // Start the download right away -- the file is what Brett came for; the change log is extra.
  const started = useRef(false);
  useEffect(() => {
    if (started.current) return;
    started.current = true;
    download();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div className="app-shell">
      <header className="app-header">
        <h1 style={{ fontSize: 16, fontWeight: 700, color: "#ffffff" }}>ATS Report Reviewer</h1>
        <span className="job-path">{source.kind === "path" ? source.path : source.name}</span>
        <button className="secondary" onClick={onReviewAnother}>
          Review Another Report
        </button>
      </header>
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
      <footer className="app-header">
        <span className="footer-stats">
          {result.aiEditCount} AI edit{result.aiEditCount === 1 ? "" : "s"} in red · {result.commentCount} comment
          {result.commentCount === 1 ? "" : "s"} for review
        </span>
        <button className="generate-report" onClick={download}>
          Download Reviewed Report
        </button>
      </footer>
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
  initialDir,
  onDirChange,
}: {
  title: string;
  instructions: string;
  extensions: string;
  onChosen: (path: string) => void;
  /** A file from the browser's own file dialog or a drag-and-drop, instead of the in-app browser. */
  onUpload: (file: File) => void;
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
  const uploadInputRef = useRef<HTMLInputElement>(null);

  function navigateTo(next: string | null) {
    const truncated = history.slice(0, historyIndex + 1);
    setHistory([...truncated, next]);
    setHistoryIndex(truncated.length);
  }

  function acceptFile(file: File | undefined) {
    if (!file) return;
    if (!/\.docx$/i.test(file.name)) {
      setError("Only Word .docx files can be reviewed.");
      return;
    }
    onUpload(file);
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
        acceptFile(e.dataTransfer.files?.[0]);
      }}
    >
      <div className={`folder-picker-card ${dragging ? "dragging" : ""}`}>
        <div className="app-logo">
          <h1>{title}</h1>
        </div>
        <p className="folder-picker-subtitle">{instructions}</p>
        {data && !data.isDriveList && data.dir && <p className="job-path">{data.dir}</p>}
        <div className="toolbar">
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
            hidden
            onChange={(e) => {
              acceptFile(e.target.files?.[0]);
              e.target.value = "";
            }}
          />
          <button onClick={() => uploadInputRef.current?.click()}>Upload a file...</button>
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
                return (
                  <button key={full} onClick={() => onChosen(full)}>
                    📄 {name}
                  </button>
                );
              })}
          </div>
        )}
        {dragging && <div className="drop-hint">Drop the report to review it</div>}
      </div>
    </div>
  );
}
