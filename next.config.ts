import type { NextConfig } from "next";
import { execSync } from "child_process";

/** "1cf4d81 · Oct 12" -- shown in the app so it's obvious whether Update.bat actually took. */
function appVersion(): string {
  try {
    return execSync('git log -1 --format="%h · %cd" --date=format:"%b %d"', { encoding: "utf8" }).trim();
  } catch {
    return "unknown version";
  }
}

const nextConfig: NextConfig = {
  env: {
    NEXT_PUBLIC_APP_VERSION: appVersion(),
  },
};

export default nextConfig;
