import { LocalWorkbookSource } from "./sources/local.js";
import { GraphWorkbookSource } from "./sources/graph.js";
import { GraphDownloadSource } from "./sources/graph-download.js";
import { TokenProvider } from "./auth.js";
import type { Config } from "./config.js";
import type { WorkbookSource } from "./types.js";

/**
 * Opens the configured workbook through the right backend.
 *
 * - `local`           — a .xlsx on disk (or downloaded from a share link)
 * - `graph`           — the Graph *workbook* API (live cell values, best fidelity)
 * - `graph_download`  — resolve via Graph, download the .xlsx, parse locally
 *                       (works for personal Microsoft accounts and app-only tokens)
 * - `auto`            — local when a local file is configured, otherwise graph
 */
export async function openWorkbook(
  config: Config,
  options: { forceSource?: "graph" | "graph_download" | "local"; localFile?: string } = {},
): Promise<WorkbookSource> {
  const source = options.forceSource ?? config.source;

  if (source === "local") return LocalWorkbookSource.open(config, options.localFile);

  const tokens = new TokenProvider(config);
  if (source === "graph_download") return GraphDownloadSource.open(config, tokens);
  return GraphWorkbookSource.connect(config, tokens);
}

