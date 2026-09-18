import type { Config } from "../config.js";
import { debugLog } from "../config.js";
import type { TokenProvider } from "../auth.js";
import type { WorkbookSource } from "../types.js";
import { resolveDriveItem, type DriveItemResponse } from "./graph-api.js";
import { LocalWorkbookSource } from "./local.js";

/**
 * 64 MB is far beyond any real financial workbook; the guard exists so a wrong
 * file id can't make the process swallow a video.
 */
const MAX_BYTES = 64 * 1024 * 1024;

interface ItemWithDownloadUrl extends DriveItemResponse {
  "@microsoft.graph.downloadUrl"?: string;
}

/**
 * Reads Excel Online **without** the Graph workbook API: it resolves the file
 * through Graph, downloads the raw `.xlsx` bytes, and parses them locally.
 *
 * This matters because the workbook API is documented as unsupported for
 * personal Microsoft accounts and for application (app-only) permissions, while
 * `driveItem/content` supports both (delegated `Files.Read` and application
 * `Files.Read.All`). It is also immune to the workbook API's .xls/IRM limits.
 *
 * The trade-off is that a workbook generated from formulas is read as last
 * *saved*, so a file open in Excel with unsaved edits may look stale.
 */
export class GraphDownloadSource implements WorkbookSource {
  private constructor(private readonly inner: LocalWorkbookSource) {}

  static async open(config: Config, tokens: TokenProvider): Promise<GraphDownloadSource> {
    const item = await resolveDriveItem(config, tokens);
    const bytes = await downloadItemBytes(config, tokens, item);

    const inner = await LocalWorkbookSource.fromBuffer(config, bytes, {
      name: item.name,
      source: "graph",
      location: `drive=${item.parentReference?.driveId} item=${item.id} (baixado)`,
      sizeBytes: item.size ?? bytes.byteLength,
      lastModified: item.lastModifiedDateTime,
      webUrl: item.webUrl,
    });
    return new GraphDownloadSource(inner);
  }

  describe(): ReturnType<WorkbookSource["describe"]> {
    return this.inner.describe();
  }

  listSheets(): ReturnType<WorkbookSource["listSheets"]> {
    return this.inner.listSheets();
  }

  readRange(...args: Parameters<WorkbookSource["readRange"]>): ReturnType<WorkbookSource["readRange"]> {
    return this.inner.readRange(...args);
  }

  async close(): Promise<void> {
    await this.inner.close();
  }
}

/** Fetches the raw bytes of a drive item, preferring the preauthenticated URL. */
export async function downloadItemBytes(
  config: Config,
  tokens: TokenProvider,
  item: DriveItemResponse,
): Promise<Buffer> {
  const driveId = item.parentReference?.driveId;
  const token = await tokens.getToken();

  // 1) @microsoft.graph.downloadUrl — preauthenticated, no Authorization header
  //    needed, and no 302 to follow. Valid only for a few minutes, so it is used
  //    immediately and never cached across processes.
  const metaUrl = driveId
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}?$select=id,@microsoft.graph.downloadUrl`
    : `https://graph.microsoft.com/v1.0/me/drive/items/${item.id}?$select=id,@microsoft.graph.downloadUrl`;

  const candidates: Array<{ url: string; headers: Record<string, string> }> = [];

  try {
    const meta = (await fetch(metaUrl, {
      headers: { authorization: `Bearer ${token}`, accept: "application/json" },
    }).then(async (response) => {
      if (!response.ok) return null;
      return (await response.json()) as ItemWithDownloadUrl;
    })) as ItemWithDownloadUrl | null;

    const direct = meta?.["@microsoft.graph.downloadUrl"];
    if (direct) candidates.push({ url: direct, headers: {} });
  } catch (error) {
    debugLog(config, "downloadUrl lookup failed:", (error as Error).message);
  }

  // 2) /content — redirects to the same preauthenticated URL; fetch follows it.
  if (driveId) {
    candidates.push({
      url: `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${item.id}/content`,
      headers: { authorization: `Bearer ${token}` },
    });
  }
  candidates.push({
    url: `https://graph.microsoft.com/v1.0/me/drive/items/${item.id}/content`,
    headers: { authorization: `Bearer ${token}` },
  });

  let lastStatus = "";
  for (const candidate of candidates) {
    const response = await fetch(candidate.url, {
      redirect: "follow",
      headers: { accept: "application/octet-stream", ...candidate.headers },
    });
    if (!response.ok) {
      lastStatus = `${response.status} ${response.statusText}`;
      debugLog(config, `download attempt failed: ${lastStatus}`);
      continue;
    }

    const declared = Number(response.headers.get("content-length") ?? "0");
    if (declared > MAX_BYTES) {
      throw new Error(
        `The file "${item.name}" is ${Math.round(declared / 1024 / 1024)} MB, over the ` +
          `${Math.round(MAX_BYTES / 1024 / 1024)} MB guard. Point EXCEL_DRIVE_ID/EXCEL_ITEM_ID at the workbook itself.`,
      );
    }

    const bytes = Buffer.from(await response.arrayBuffer());
    if (bytes.byteLength > MAX_BYTES) {
      throw new Error(`The downloaded file exceeds the ${MAX_BYTES / 1024 / 1024} MB guard.`);
    }

    // xlsx is a zip archive: "PK\x03\x04". Anything else means we got an HTML
    // error page or a different document type.
    const isZip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
    if (!isZip) {
      lastStatus = `payload is not an .xlsx archive (content-type ${response.headers.get("content-type") ?? "?"} )`;
      debugLog(config, lastStatus);
      continue;
    }

    return bytes;
  }

  throw new Error(
    `Could not download "${item.name}" from OneDrive/SharePoint (${lastStatus || "no attempt succeeded"}).`,
  );
}
