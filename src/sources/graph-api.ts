import type { Config } from "../config.js";
import { debugLog } from "../config.js";
import { AuthError, GRAPH, TokenProvider } from "../auth.js";

const API = `${GRAPH}/v1.0`;

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly graphCode: string,
    readonly hint?: string,
  ) {
    super(message);
    this.name = "GraphError";
  }
}

export interface DriveItemRef {
  driveId: string;
  itemId: string;
}

export interface DriveItemResponse {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  parentReference?: { driveId?: string; driveType?: string; path?: string };
}

interface GraphErrorBody {
  error?: { code?: string; message?: string };
}

export function encodeDrivePath(path: string): string {
  const normalized = path.startsWith("/") ? path : `/${path}`;
  return normalized
    .split("/")
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

/** Microsoft accepts sharing links base64-encoded (url-safe, `u!` prefixed). */
export function encodeSharingUrl(url: string): string {
  const base64 = Buffer.from(url, "utf8").toString("base64");
  const urlSafe = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  return `u!${urlSafe}`;
}

export function hintFor(status: number, code: string): string | undefined {
  if (status === 401 || code === "InvalidAuthenticationToken") {
    return "The access token was rejected. Run `npm run login` again to refresh credentials.";
  }
  if (status === 403 || code === "accessDenied") {
    return (
      "Permission denied. The signed-in account must be able to read the file " +
      "(Files.Read.All / Sites.Read.All consent), or the file must be shared with it. " +
      "For a share link specifically, Microsoft documents Files.ReadWrite as the " +
      "least-privileged delegated permission — add it to EXCEL_SCOPES and re-run `npm run login`."
    );
  }
  if (status === 404 || code === "itemNotFound") {
    return (
      "File or worksheet not found. Check EXCEL_WORKBOOK_PATH (drive-relative, e.g. " +
      "/Documentos/PLANILHA.xlsx), EXCEL_DRIVE_ID + EXCEL_ITEM_ID, or EXCEL_SHARE_URL. " +
      "Also confirm the file is saved as .xlsx — the Graph workbook API cannot read legacy .xls."
    );
  }
  if (code === "invalidRequest") {
    return (
      "Graph rejected the workbook request. Legacy .xls files and workbooks protected by " +
      "Information Rights Management are not supported — save/convert to .xlsx."
    );
  }
  return undefined;
}

export async function graphFetch<T>(
  config: Config,
  tokens: TokenProvider,
  url: string,
  attempts = 4,
): Promise<T> {
  const token = await tokens.getToken();
  const res = await fetch(url, {
    headers: { authorization: `Bearer ${token}`, accept: "application/json" },
  });

  if ((res.status === 429 || res.status >= 500) && attempts > 1) {
    const retryAfter = Number(res.headers.get("retry-after") ?? "1");
    const waitMs = Math.min(Math.max(retryAfter, 1), 10) * 1000;
    debugLog(config, `retrying ${res.status} in ${waitMs}ms`);
    await new Promise((r) => setTimeout(r, waitMs));
    return graphFetch<T>(config, tokens, url, attempts - 1);
  }

  const bodyText = await res.text();
  let parsed: unknown = {};
  try {
    parsed = bodyText ? JSON.parse(bodyText) : {};
  } catch {
    parsed = {};
  }

  if (!res.ok) {
    const err = parsed as GraphErrorBody;
    const code = err.error?.code ?? String(res.status);
    throw new GraphError(
      `Graph ${res.status} ${code}: ${err.error?.message ?? res.statusText}`,
      res.status,
      code,
      hintFor(res.status, code),
    );
  }
  return parsed as T;
}

export async function resolveDriveItem(
  config: Config,
  tokens: TokenProvider,
): Promise<DriveItemResponse> {
  const select = "$select=id,name,size,webUrl,lastModifiedDateTime,file,parentReference";

  let url: string;
  let locationHint: string;

  if (config.driveId && config.itemId) {
    url = `${API}/drives/${config.driveId}/items/${config.itemId}?${select}`;
    locationHint = "EXCEL_DRIVE_ID + EXCEL_ITEM_ID";
  } else if (config.workbookPath) {
    url = `${API}/me/drive/root:${encodeDrivePath(config.workbookPath)}?${select}`;
    locationHint = `EXCEL_WORKBOOK_PATH=${config.workbookPath}`;
  } else if (config.shareUrl) {
    url = `${API}/shares/${encodeSharingUrl(config.shareUrl)}/driveItem?${select}`;
    locationHint = "EXCEL_SHARE_URL";
  } else {
    throw new AuthError(
      "No workbook configured.",
      "Set one of EXCEL_WORKBOOK_PATH, EXCEL_DRIVE_ID + EXCEL_ITEM_ID, or EXCEL_SHARE_URL in .env.",
    );
  }

  try {
    return await graphFetch<DriveItemResponse>(config, tokens, url);
  } catch (error) {
    if (error instanceof GraphError) {
      throw new GraphError(
        `Could not resolve workbook via ${locationHint}: ${error.message}`,
        error.status,
        error.graphCode,
        error.hint,
      );
    }
    throw error;
  }
}
