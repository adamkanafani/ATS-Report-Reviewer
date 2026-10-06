/**
 * Table of Contents update through the user's own installed Microsoft Word (COM automation via
 * PowerShell). Page numbers only exist once Word lays the document out, so there's no way to
 * compute an accurate TOC from the XML alone -- Word itself has to rebuild it.
 *
 * Careful about the user's own Word session: `New-Object -ComObject Word.Application` attaches to
 * an already-running Word if there is one, so the script opens the report invisibly, closes only
 * that document, and quits Word only if it was the one that started it.
 *
 * When Word isn't available (or fails), the caller falls back to flagging the document so Word
 * offers to update its fields the next time it's opened.
 */
import { execFile } from "child_process";
import { promises as fs } from "fs";
import os from "os";
import path from "path";
import { randomUUID } from "crypto";

const SCRIPT = String.raw`
param([string]$Path)
$ErrorActionPreference = 'Stop'
$word = $null
$doc = $null
$startedWord = $false
try {
  $running = $null
  try { $running = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { }
  if ($running) { $word = $running } else { $word = New-Object -ComObject Word.Application; $startedWord = $true }
  $word.DisplayAlerts = 0
  $missing = [Type]::Missing
  $doc = $word.Documents.Open($Path, $false, $false, $false, $missing, $missing, $missing, $missing, $missing, $missing, $missing, $false)
  $count = $doc.TablesOfContents.Count
  for ($i = 1; $i -le $count; $i++) { $doc.TablesOfContents.Item($i).Update() }
  # Second pass: rebuilding the TOC can change its own length and shift every page after it.
  for ($i = 1; $i -le $count; $i++) { $doc.TablesOfContents.Item($i).UpdatePageNumbers() }
  $doc.Save()
  Write-Output "TOC_COUNT=$count"
} finally {
  # Word's Close/Quit take their SaveChanges argument by reference -- a plain 0 throws in
  # PowerShell, which would silently leave a hidden Word running.
  $noSave = 0
  if ($doc) { try { $doc.Close([ref]$noSave) } catch { } }
  if ($word -and $startedWord) {
    try { $word.Quit([ref]$noSave) } catch { try { $word.Quit() } catch { } }
  }
  if ($word) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
`;

export interface WordTocResult {
  ok: boolean;
  tocCount: number;
  buffer?: Buffer;
  error?: string;
}

export async function updateTocWithWord(input: Buffer, timeoutMs = 5 * 60 * 1000): Promise<WordTocResult> {
  if (process.platform !== "win32") return { ok: false, tocCount: 0, error: "Microsoft Word automation is only available on Windows." };
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ats-reviewer-"));
  const docPath = path.join(dir, `${randomUUID()}.docx`);
  const scriptPath = path.join(dir, "update-toc.ps1");
  try {
    await fs.writeFile(docPath, input);
    await fs.writeFile(scriptPath, SCRIPT, "utf8");
    const stdout = await new Promise<string>((resolve, reject) => {
      execFile(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-Path", docPath],
        { timeout: timeoutMs, windowsHide: true },
        (err, out, errOut) => {
          if (err) reject(new Error((errOut || err.message).trim().split(/\r?\n/).slice(0, 3).join(" ")));
          else resolve(out);
        },
      );
    });
    const tocCount = Number(/TOC_COUNT=(\d+)/.exec(stdout)?.[1] ?? 0);
    return { ok: true, tocCount, buffer: await fs.readFile(docPath) };
  } catch (err) {
    return { ok: false, tocCount: 0, error: err instanceof Error ? err.message : String(err) };
  } finally {
    fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

/** Fallback: ask Word to refresh fields (including the TOC) when the file is next opened. */
export function flagUpdateFieldsOnOpen(settingsXml: string): string {
  if (/<w:updateFields\b/.test(settingsXml)) return settingsXml.replace(/<w:updateFields\b[^>]*\/>/, '<w:updateFields w:val="true"/>');
  // updateFields sits just before these in the CT_Settings sequence.
  const after = /<(w:hdrShapeDefaults|w:footnotePr|w:endnotePr|w:compat|w:docVars|w:rsids|m:mathPr|w:themeFontLang|w:clrSchemeMapping|w:decimalSymbol|w:listSeparator)\b/.exec(
    settingsXml,
  );
  const tag = '<w:updateFields w:val="true"/>';
  if (after) return settingsXml.slice(0, after.index) + tag + settingsXml.slice(after.index);
  return settingsXml.replace("</w:settings>", `${tag}</w:settings>`);
}
