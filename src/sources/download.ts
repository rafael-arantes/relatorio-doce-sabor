import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import type { Config } from "../config.js";
import { debugLog } from "../config.js";

/** Turns a share link into a direct-download URL OneDrive respects. */
export function downloadUrl(shareUrl: string): string {
  const url = new URL(shareUrl);
  url.searchParams.set("download", "1");
  return url.toString();
}

/**
 * Best-effort download of a OneDrive/SharePoint share link.
 *
 * Works for links shared as "Anyone with the link". Files restricted to
 * specific people need the Graph backend instead (`EXCEL_SOURCE=graph`).
 */
export async function downloadShareLink(
  config: Config,
  shareUrl: string,
  destination: string,
): Promise<string> {
  const candidates = [downloadUrl(shareUrl), shareUrl];

  for (const candidate of candidates) {
    debugLog(config, "downloading", candidate);
    const res = await fetch(candidate, {
      redirect: "follow",
      headers: { "user-agent": "excel-mcp/0.1", accept: "*/*" },
    });
    if (!res.ok) {
      debugLog(config, `download failed with ${res.status}`);
      continue;
    }

    const contentType = res.headers.get("content-type") ?? "";
    const bytes = Buffer.from(await res.arrayBuffer());
    // xlsx is a zip archive: "PK\x03\x04"
    const isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
    if (!isZip) {
      debugLog(config, `not a zip payload (${contentType}), trying next candidate`);
      continue;
    }

    mkdirSync(dirname(destination), { recursive: true });
    writeFileSync(destination, bytes);
    return destination;
  }

  throw new Error(
    "Could not download the workbook from EXCEL_SHARE_URL as a raw .xlsx. " +
      "The link is probably restricted to specific people or serves an HTML page. " +
      "Either switch to EXCEL_SOURCE=graph and run `npm run login`, or download the " +
      "file manually and point EXCEL_LOCAL_FILE at it.",
  );
}

export function fileSize(path: string): number | undefined {
  try {
    return existsSync(path) ? readFileSync(path).byteLength : undefined;
  } catch {
    return undefined;
  }
}
