import type { Config } from "../config.js";
import { debugLog } from "../config.js";
import type { TokenProvider } from "../auth.js";
import { makeGrid, mergeMatrices, trimEmptyEdges } from "../cells.js";
import type { Grid, SheetInfo, WorkbookDescription, WorkbookSource } from "../types.js";
import {
  graphFetch,
  GraphError,
  resolveDriveItem,
  type DriveItemRef,
  type DriveItemResponse,
} from "./graph-api.js";
import { GraphDownloadSource } from "./graph-download.js";

const API = "https://graph.microsoft.com/v1.0";

interface RangeResponse {
  address?: string;
  rowCount?: number;
  columnCount?: number;
  values?: unknown[][];
  text?: unknown[][];
}

/**
 * Errors meaning "the workbook API is not available for this item/account", as
 * opposed to "you asked for the wrong sheet". Only the former triggers the
 * download fallback.
 */
export function isWorkbookApiUnavailable(error: unknown): boolean {
  if (!(error instanceof GraphError)) return false;
  if (error.status === 401 || error.status === 403 || error.status >= 500) return true;
  return [
    "invalidRequest",
    "notSupported",
    "generalException",
    "accessDenied",
    "unknownError",
    "ErrorInvalidFileType",
    "FileNotXlsx",
  ].includes(error.graphCode);
}

/**
 * Reads a workbook that lives in OneDrive / SharePoint ("Excel Online") through
 * the Microsoft Graph workbook API. Read-only: only GET requests are issued.
 *
 * If the workbook API proves unavailable for the account or file, this degrades
 * transparently to `GraphDownloadSource` (resolve via Graph, download the .xlsx,
 * parse locally) — the documented workaround for personal Microsoft accounts and
 * for app-only (client credentials) tokens.
 */
export class GraphWorkbookSource implements WorkbookSource {
  private fallback: GraphDownloadSource | null = null;
  private fallbackReason: string | null = null;
  private usingFallback = false;

  private constructor(
    private readonly config: Config,
    private readonly tokens: TokenProvider,
    private readonly ref: DriveItemRef,
    private readonly item: DriveItemResponse,
  ) {}

  static async connect(config: Config, tokens: TokenProvider): Promise<GraphWorkbookSource> {
    const item = await resolveDriveItem(config, tokens);
    const driveId = item.parentReference?.driveId;
    if (!driveId) {
      throw new Error(
        "Graph did not return the drive id for this item. Pass EXCEL_DRIVE_ID + EXCEL_ITEM_ID explicitly.",
      );
    }
    return new GraphWorkbookSource(config, tokens, { driveId, itemId: item.id }, item);
  }

  /**
   * Runs a workbook-API operation; on a "not available" error it opens the
   * downloaded copy once and retries the same operation against it.
   */
  private async run<T>(
    api: () => Promise<T>,
    viaFallback: (source: WorkbookSource) => Promise<T>,
  ): Promise<T> {
    if (this.usingFallback && this.fallback) return viaFallback(this.fallback);

    try {
      return await api();
    } catch (error) {
      if (!this.config.graphFallback || !isWorkbookApiUnavailable(error)) throw error;

      this.fallbackReason = error instanceof Error ? error.message : String(error);
      console.error(
        `[excel-mcp] Graph recusou a API de workbook (${this.fallbackReason}); ` +
          "baixando o arquivo e lendo localmente.",
      );
      this.fallback ??= await GraphDownloadSource.open(this.config, this.tokens);
      this.usingFallback = true;
      return viaFallback(this.fallback);
    }
  }

  /** True once the server had to fall back to a downloaded copy. */
  get degraded(): boolean {
    return this.usingFallback;
  }

  get fallbackInfo(): string | null {
    return this.fallbackReason;
  }

  private workbookUrl(suffix: string): string {
    return `${API}/drives/${this.ref.driveId}/items/${this.ref.itemId}/workbook${suffix}`;
  }

  private worksheetUrl(sheet: string): string {
    return this.workbookUrl(`/worksheets/${encodeURIComponent(sheet)}`);
  }

  async describe(): Promise<WorkbookDescription> {
    return this.run(
      async () => {
        const sheets = await this.listSheetsViaApi();
        return {
          name: this.item.name,
          source: "graph" as const,
          location: `drive=${this.ref.driveId} item=${this.ref.itemId}`,
          sizeBytes: this.item.size,
          lastModified: this.item.lastModifiedDateTime,
          webUrl: this.item.webUrl,
          sheets,
        };
      },
      (source) => source.describe(),
    );
  }

  async listSheets(): Promise<SheetInfo[]> {
    return this.run(
      () => this.listSheetsViaApi(),
      (source) => source.listSheets(),
    );
  }

  private async listSheetsViaApi(): Promise<SheetInfo[]> {
    const response = await graphFetch<{
      value?: Array<{ id: string; name: string; position?: number; visibility?: string }>;
    }>(this.config, this.tokens, this.workbookUrl("/worksheets?$select=id,name,position,visibility"));

    return (response.value ?? []).map((sheet) => ({
      name: sheet.name,
      id: sheet.id,
      position: sheet.position,
      visibility: sheet.visibility,
    }));
  }

  async readRange(sheet: string, address?: string): Promise<Grid> {
    return this.run(
      () => this.readRangeViaApi(sheet, address),
      (source) => source.readRange(sheet, address),
    );
  }

  private async readRangeViaApi(sheet: string, address?: string): Promise<Grid> {
    if (address) {
      const url = `${this.worksheetUrl(sheet)}/range(address='${encodeURIComponent(address)}')?$select=address,rowCount,columnCount,values,text`;
      return this.toGrid(sheet, await graphFetch<RangeResponse>(this.config, this.tokens, url));
    }

    // `?$select=text` is not honoured by every tenant/drive type, so try the
    // richest shape first and degrade instead of failing outright.
    const attempts = [
      `${this.worksheetUrl(sheet)}/usedRange(valuesOnly=true)?$select=address,rowCount,columnCount,values,text`,
      `${this.worksheetUrl(sheet)}/usedRange?$select=address,rowCount,columnCount,values,text`,
      `${this.worksheetUrl(sheet)}/usedRange(valuesOnly=true)`,
    ];

    let lastError: unknown;
    for (const url of attempts) {
      try {
        return this.toGrid(sheet, await graphFetch<RangeResponse>(this.config, this.tokens, url));
      } catch (error) {
        lastError = error;
        debugLog(this.config, "usedRange attempt failed:", (error as Error).message);
      }
    }
    throw lastError instanceof Error ? lastError : new Error(String(lastError));
  }

  private toGrid(sheet: string, range: RangeResponse): Grid {
    const grid = makeGrid(
      sheet,
      mergeMatrices((range.values ?? []) as unknown[][], range.text as unknown[][] | undefined),
      false,
    );
    return trimEmptyEdges({
      ...grid,
      address: range.address ?? grid.address,
      rowCount: range.rowCount ?? grid.rowCount,
      columnCount: range.columnCount ?? grid.columnCount,
    });
  }

  async close(): Promise<void> {
    await this.fallback?.close();
  }
}

export { resolveDriveItem };
