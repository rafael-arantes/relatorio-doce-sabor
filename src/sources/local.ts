import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import type { Config } from "../config.js";
import { debugLog } from "../config.js";
import { normalizeCell, makeGrid, trimEmptyEdges } from "../cells.js";
import type { Cell, Grid, SheetInfo, WorkbookDescription, WorkbookSource } from "../types.js";
import { downloadShareLink, fileSize } from "./download.js";

interface ExcelJSWorkbook {
  worksheets: ExcelJSWorksheet[];
  xlsx: {
    readFile(path: string): Promise<unknown>;
    load(buffer: Buffer): Promise<unknown>;
  };
}
interface ExcelJSWorksheet {
  name: string;
  id: number;
  state?: string;
  rowCount: number;
  columnCount: number;
  actualRowCount: number;
  actualColumnCount: number;
  getCell(row: number, column: number): { value: unknown };
}
interface ExcelJSModule {
  Workbook: new () => ExcelJSWorkbook;
}

/** Where a workbook came from — carried through to `describe()`. */
export interface WorkbookMeta {
  name: string;
  source: "local" | "graph";
  location: string;
  sizeBytes?: number;
  lastModified?: string;
  webUrl?: string;
}


let excelJsCache: ExcelJSModule | null = null;
async function loadExcelJS(): Promise<ExcelJSModule> {
  if (!excelJsCache) {
    // exceljs is CJS; the default export holds the module namespace under ESM.
    const imported = (await import("exceljs")) as unknown as
      | ExcelJSModule
      | { default: ExcelJSModule };
    const mod = "Workbook" in imported ? imported : (imported as { default: ExcelJSModule }).default;
    if (!mod?.Workbook) throw new Error("Failed to load exceljs (Workbook class not found).");
    excelJsCache = mod;
  }
  return excelJsCache;
}

function readSheetGrid(
  sheet: ExcelJSWorksheet,
  maxRows: number,
  maxCells: number,
): { rows: Cell[][]; truncated: boolean } {
  // `actualColumnCount` stops at the first empty column, which would silently
  // hide blocks sitting further right (the classic per-supplier layout), so the
  // worksheet's real dimension is the source of truth.
  const rowCount = Math.min(Math.max(sheet.rowCount || 0, sheet.actualRowCount || 0), maxRows);
  const columnCount = Math.min(
    Math.max(sheet.columnCount || 0, sheet.actualColumnCount || 0),
    Math.min(maxCells, 300),
  );
  const rows: Cell[][] = [];
  let cells = 0;
  let truncated = false;

  for (let r = 1; r <= rowCount; r++) {
    const row: Cell[] = [];
    for (let c = 1; c <= columnCount; c++) {
      if (cells >= maxCells) {
        truncated = true;
        break;
      }
      row.push(normalizeCell(sheet.getCell(r, c).value));
      cells++;
    }
    if (truncated) break;
    rows.push(row);
  }

  return { rows, truncated };
}

export function parseA1(address: string): {
  start: { row: number; column: number };
  end: { row: number; column: number };
} {
  const cleaned = address.replace(/^[^!]*!/, "").replace(/\$/g, "");
  const [left, right] = cleaned.split(":");
  const from = parseCellRef(left ?? "A1");
  const to = right ? parseCellRef(right) : from;
  return {
    start: { row: Math.min(from.row, to.row), column: Math.min(from.column, to.column) },
    end: { row: Math.max(from.row, to.row), column: Math.max(from.column, to.column) },
  };
}

function parseCellRef(ref: string): { row: number; column: number } {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref.trim());
  if (!match) throw new Error(`Invalid cell reference "${ref}" (expected e.g. A1 or B2:C10).`);
  const letters = (match[1] ?? "").toUpperCase();
  let column = 0;
  for (const char of letters) column = column * 26 + (char.charCodeAt(0) - 64);
  return { row: Number(match[2]), column };
}

/**
 * Reads a .xlsx from the filesystem, from a downloaded share link, or from bytes
 * fetched through Graph (see `graph-download.ts`) — the parsing is identical, only
 * the origin differs, which is why this class takes a `meta` descriptor.
 */
export class LocalWorkbookSource implements WorkbookSource {
  private constructor(
    private readonly config: Config,
    private readonly workbook: ExcelJSWorkbook,
    private readonly meta: WorkbookMeta,
  ) {}

  static async open(config: Config, pathOverride?: string): Promise<LocalWorkbookSource> {
    let path = pathOverride ?? config.localFile;
    if (!path && config.shareUrl) {
      path = resolve(config.projectRoot, ".tokens", "workbook-download.xlsx");
      await downloadShareLink(config, config.shareUrl, path);
    }
    if (!path) throw new Error("EXCEL_SOURCE=local needs EXCEL_LOCAL_FILE or EXCEL_SHARE_URL.");

    let bytes: Buffer;
    try {
      bytes = readFileSync(path);
    } catch (error) {
      throw new Error(`Workbook not found at ${path}: ${(error as Error).message}`);
    }
    return LocalWorkbookSource.fromBuffer(config, bytes, {
      name: path.split("/").pop() ?? path,
      source: "local",
      location: path,
      sizeBytes: fileSize(path),
    });
  }

  /** Opens a workbook already held in memory (e.g. downloaded via Graph). */
  static async fromBuffer(
    config: Config,
    bytes: Buffer,
    meta: WorkbookMeta,
  ): Promise<LocalWorkbookSource> {
    const { Workbook } = await loadExcelJS();
    const workbook = new Workbook();
    try {
      await workbook.xlsx.load(bytes);
    } catch (error) {
      throw new Error(
        `Could not open "${meta.name}" from ${meta.location}: ${(error as Error).message}. ` +
          "Legacy .xls files must be re-saved as .xlsx first — the .xlsx format is a zip archive.",
      );
    }
    debugLog(config, `opened ${meta.name} (${meta.source}) with ${bytes.byteLength} bytes`);
    return new LocalWorkbookSource(config, workbook, meta);
  }

  async describe(): Promise<WorkbookDescription> {
    return { ...this.meta, sheets: await this.listSheets() };
  }

  async listSheets(): Promise<SheetInfo[]> {
    return this.workbook.worksheets.map((sheet, index) => ({
      name: sheet.name,
      id: String(sheet.id),
      position: index + 1,
      visibility: sheet.state ?? "visible",
      rowCount: sheet.actualRowCount || sheet.rowCount,
      columnCount: sheet.actualColumnCount || sheet.columnCount,
    }));
  }

  private findSheet(name: string): ExcelJSWorksheet {
    const wanted = name.trim().toLowerCase();
    const sheet =
      this.workbook.worksheets.find((s) => s.name.toLowerCase() === wanted) ??
      this.workbook.worksheets.find((s) => s.name.toLowerCase().startsWith(wanted)) ??
      (Number.isInteger(Number(name)) ? this.workbook.worksheets[Number(name) - 1] : undefined);

    if (!sheet) {
      const available = this.workbook.worksheets.map((s) => s.name).join(", ");
      throw new Error(`Worksheet "${name}" not found. Available sheets: ${available}`);
    }
    return sheet;
  }

  async readRange(sheet: string, address?: string): Promise<Grid> {
    const target = this.findSheet(sheet);
    const { rows, truncated } = readSheetGrid(target, this.config.maxRows, this.config.maxCells);

    if (!address) return trimEmptyEdges(makeGrid(target.name, rows, truncated));

    const { start, end } = parseA1(address);
    const sliced: Cell[][] = [];
    for (let r = start.row; r <= Math.min(end.row, rows.length); r++) {
      const sourceRow = rows[r - 1] ?? [];
      const row: Cell[] = [];
      for (let c = start.column; c <= end.column; c++) row.push(sourceRow[c - 1] ?? null);
      sliced.push(row);
    }
    return makeGrid(target.name, sliced, truncated, start.row, start.column - 1);
  }

  async close(): Promise<void> {
    /* nothing to release */
  }
}
