import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";

/**
 * Minimal .env loader (no dotenv dependency).
 * Real environment variables always win over the file, so a wrapper process
 * (Claude Code / Claude Desktop) can override anything.
 */
function loadDotEnv(file: string): void {
  if (!existsSync(file)) return;
  let raw: string;
  try {
    raw = readFileSync(file, "utf8");
  } catch {
    return;
  }
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq <= 0) continue;
    const key = trimmed.slice(0, eq).trim();
    let value = trimmed.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (process.env[key] === undefined) process.env[key] = value;
  }
}

const startDir = import.meta.dirname;

/**
 * Walks up from this module until it finds the package root, so `.env` and the
 * token cache resolve identically whether we run from `src/` (tsx) or from
 * `dist/src/` (compiled).
 */
export function findProjectRoot(): string {
  let current = resolve(startDir);
  for (let depth = 0; depth < 6; depth++) {
    if (existsSync(resolve(current, "package.json"))) return current;
    const parent = resolve(current, "..");
    if (parent === current) break;
    current = parent;
  }
  return resolve(startDir, "..");
}

const projectRoot = findProjectRoot();
loadDotEnv(resolve(projectRoot, ".env"));
loadDotEnv(resolve(projectRoot, ".env.local"));

function str(name: string, fallback = ""): string {
  const v = process.env[name];
  return v === undefined || v === "" ? fallback : v.trim();
}

function int(name: string, fallback: number): number {
  const v = Number.parseInt(str(name, ""), 10);
  return Number.isFinite(v) && v > 0 ? v : fallback;
}

function bool(name: string): boolean {
  return ["1", "true", "yes", "on"].includes(str(name, "").toLowerCase());
}

export type SourceMode = "graph" | "graph_download" | "local";
export type AuthMode = "device_code" | "client_secret";

/**
 * Microsoft's first-party "Microsoft Graph Command Line Tools" public client.
 * Lets the device-code flow work with zero Azure app registration; the user
 * still has to consent interactively. Override with EXCEL_CLIENT_ID if you
 * prefer your own registration.
 */
export const GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

export interface Config {
  projectRoot: string;
  source: SourceMode;
  auth: AuthMode;
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scopes: string[];
  tokenCache: string;
  workbookPath: string;
  driveId: string;
  itemId: string;
  shareUrl: string;
  localFile: string;
  maxRows: number;
  maxCells: number;
  /** Retry a failed workbook-API call by downloading the file instead. */
  graphFallback: boolean;
  /** Refresh token via env (server deploys without .tokens/graph.json). */
  refreshTokenEnv: string;
  debug: boolean;
}

export function loadConfig(): Config {
  const shareUrl = str("EXCEL_SHARE_URL");
  const localFile = str("EXCEL_LOCAL_FILE", str("EXCEL_XLSX_PATH"));

  let source = str("EXCEL_SOURCE", "auto").toLowerCase();
  let resolved: SourceMode;
  if (source === "local" || source === "graph" || source === "graph_download") {
    resolved = source as SourceMode;
  } else {
    resolved = localFile ? "local" : "graph";
  }

  const auth = str("EXCEL_AUTH", "device_code").toLowerCase();
  const authMode: AuthMode =
    auth === "client_secret" || auth === "app_only" || auth === "client_credentials"
      ? "client_secret"
      : "device_code";

  return {
    projectRoot,
    source: resolved,
    auth: authMode,
    tenantId: str("EXCEL_TENANT_ID", "common"),
    clientId: str("EXCEL_CLIENT_ID", GRAPH_CLI_CLIENT_ID),
    clientSecret: str("EXCEL_CLIENT_SECRET"),
    scopes: str("EXCEL_SCOPES", "offline_access Files.Read.All Sites.Read.All")
      .split(/[\s,]+/)
      .filter(Boolean),
    tokenCache: resolve(projectRoot, str("EXCEL_TOKEN_CACHE", ".tokens/graph.json")),
    workbookPath: str("EXCEL_WORKBOOK_PATH"),
    driveId: str("EXCEL_DRIVE_ID"),
    itemId: str("EXCEL_ITEM_ID"),
    shareUrl,
    localFile,
    maxRows: int("EXCEL_MAX_ROWS", 5000),
    maxCells: int("EXCEL_MAX_CELLS", 200_000),
    graphFallback: !["0", "false", "no", "off"].includes(str("EXCEL_GRAPH_FALLBACK", "1").toLowerCase()),
    refreshTokenEnv: str("EXCEL_REFRESH_TOKEN"),
    debug: bool("EXCEL_DEBUG"),
  };
}

export function debugLog(config: Config, ...args: unknown[]): void {
  if (config.debug) console.error("[excel-mcp]", ...args);
}

/** Returns human-readable setup problems; empty array means the config is usable. */
export function validateConfig(config: Config): string[] {
  const problems: string[] = [];
  if (config.source === "graph" || config.source === "graph_download") {
    if (!config.shareUrl && !config.localFile) {
      const located = (config.driveId && config.itemId) || config.workbookPath;
      if (!located) {
        problems.push(
          "No workbook configured. Set exactly one of EXCEL_WORKBOOK_PATH, " +
            "EXCEL_DRIVE_ID + EXCEL_ITEM_ID, or EXCEL_SHARE_URL.",
        );
      }
    }
    if (config.auth === "client_secret") {
      if (!config.clientSecret) problems.push("EXCEL_AUTH=client_secret requires EXCEL_CLIENT_SECRET.");
      if (!config.clientId) problems.push("EXCEL_AUTH=client_secret requires EXCEL_CLIENT_ID.");
      if (config.tenantId === "common")
        problems.push("EXCEL_AUTH=client_secret requires a real EXCEL_TENANT_ID (not 'common').");
      if (config.source === "graph") {
        problems.push(
          "EXCEL_AUTH=client_secret (app-only) is not supported by the Graph workbook API. " +
            "Set EXCEL_SOURCE=graph_download to read the file through driveItem/content instead.",
        );
      }
    } else if (!config.clientId) {
      problems.push("EXCEL_AUTH=device_code requires EXCEL_CLIENT_ID.");
    }
  } else if (!config.localFile && !config.shareUrl) {
    problems.push("EXCEL_SOURCE=local requires EXCEL_LOCAL_FILE (or EXCEL_SHARE_URL to download from).");
  }
  return problems;
}
