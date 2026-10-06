import { NextRequest, NextResponse } from "next/server";
import { promises as fs } from "fs";
import path from "path";
import os from "os";
import { exec } from "child_process";
import { promisify } from "util";

const execAsync = promisify(exec);

/** Matches Windows Explorer's "natural" name sort (numeric-aware). */
function naturalCompare(a: string, b: string): number {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}

/** Sentinel meaning "the drives list" (Explorer's "This PC") -- see Tech-Rep-Report-App's
 *  api/browse/route.ts, which this file is adapted from. */
const DRIVES_ROOT = "__DRIVES__";

/** Sentinel meaning "jump to OneDrive" -- OneDrive syncs to a plain folder under the user's
 *  profile, not its own drive letter, so it never shows up in the "This PC" drives list (the
 *  same as real Windows Explorer). Resolved here rather than hardcoded client-side since the
 *  exact folder name is tenant-specific. */
const ONEDRIVE_ROOT = "__ONEDRIVE__";

async function resolveOneDriveDir(): Promise<string | null> {
  const home = os.homedir();
  const preferred = path.join(home, "OneDrive - Allied Power Group");
  try {
    if ((await fs.stat(preferred)).isDirectory()) return preferred;
  } catch {
    // fall through to a generic search below
  }
  try {
    const entries = await fs.readdir(home, { withFileTypes: true });
    const match = entries.find((e) => e.isDirectory() && e.name.toLowerCase().startsWith("onedrive"));
    return match ? path.join(home, match.name) : null;
  } catch {
    return null;
  }
}

async function listDrives(): Promise<string[]> {
  try {
    const { stdout } = await execAsync('powershell -NoProfile -Command "(Get-PSDrive -PSProvider FileSystem).Root"');
    return stdout
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .sort(naturalCompare);
  } catch {
    return [];
  }
}

async function mapWithConcurrency<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const i = next++;
      results[i] = await fn(items[i]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

const FS_CONCURRENCY = 16;

interface DirListing {
  folders: string[];
  files: string[];
}

/** Lists `dir`'s direct children, split into folder names and file names (filtered to
 *  `extensions`). Unlike the INA app's folder-only version, this app's file picker needs to
 *  show and select specific files (the raw MDI doc, the ATS template), not just navigate
 *  toward a job folder. */
async function listDir(dir: string, extensions: string[]): Promise<DirListing> {
  let entries;
  try {
    entries = await fs.readdir(dir, { withFileTypes: true });
  } catch {
    return { folders: [], files: [] };
  }

  const classified = await mapWithConcurrency(entries, FS_CONCURRENCY, async (e) => {
    if (e.name.startsWith(".")) return null;
    if (e.isDirectory()) return { kind: "folder" as const, name: e.name };
    if (e.isFile()) return matchesExtension(e.name, extensions) ? { kind: "file" as const, name: e.name } : null;
    // OneDrive placeholder dirents sometimes report as symlinks -- stat() to find out what they really are.
    try {
      const stat = await fs.stat(path.join(dir, e.name));
      if (stat.isDirectory()) return { kind: "folder" as const, name: e.name };
      if (stat.isFile() && matchesExtension(e.name, extensions)) return { kind: "file" as const, name: e.name };
      return null;
    } catch {
      return null;
    }
  });

  const folders = classified.filter((c) => c?.kind === "folder").map((c) => c!.name);
  const files = classified.filter((c) => c?.kind === "file").map((c) => c!.name);
  return { folders: folders.sort(naturalCompare), files: files.sort(naturalCompare) };
}

function matchesExtension(name: string, extensions: string[]): boolean {
  if (extensions.length === 0) return true;
  const ext = path.extname(name).toLowerCase().replace(/^\./, "");
  return extensions.includes(ext);
}

export async function GET(request: NextRequest) {
  const dirParam = request.nextUrl.searchParams.get("dir");
  const extParam = request.nextUrl.searchParams.get("ext");
  const extensions = extParam ? extParam.split(",").map((e) => e.trim().toLowerCase()) : [];

  if (dirParam === DRIVES_ROOT) {
    const folders = await listDrives();
    return NextResponse.json({ dir: null, parent: null, folders, files: [], isDriveList: true });
  }

  let dir = dirParam || os.homedir();
  if (dirParam === ONEDRIVE_ROOT) {
    const resolved = await resolveOneDriveDir();
    if (!resolved) {
      return NextResponse.json({ error: "Could not find a OneDrive folder under your user profile." }, { status: 400 });
    }
    dir = resolved;
  }

  try {
    const { folders, files } = await listDir(dir, extensions);
    const dirnameOfDir = path.dirname(dir);
    const parent = dirnameOfDir === dir ? DRIVES_ROOT : dirnameOfDir;
    return NextResponse.json({ dir, parent, folders, files, isDriveList: false });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : "Could not read directory" }, { status: 400 });
  }
}
