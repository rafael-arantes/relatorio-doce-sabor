export type Cell = string | number | boolean | null;

export interface SheetInfo {
  name: string;
  id?: string;
  position?: number;
  visibility?: string;
  /** e.g. "A1:H240" — only populated when the backend can report it cheaply. */
  usedRange?: string;
  rowCount?: number;
  columnCount?: number;
}

export interface WorkbookDescription {
  name: string;
  source: "graph" | "local";
  /** Path, drive+item id or local absolute path — whatever identifies it. */
  location: string;
  sizeBytes?: number;
  lastModified?: string;
  webUrl?: string;
  sheets: SheetInfo[];
}

export interface Grid {
  sheet: string;
  /** A1 address the grid actually covers, e.g. "A1:H240". */
  address: string;
  /** Fully rectangular: every row has the same length. */
  rows: Cell[][];
  rowCount: number;
  columnCount: number;
  /** True when the backend cut the grid short because of EXCEL_MAX_ROWS/CELLS. */
  truncated: boolean;
  /** First row number of the grid inside the sheet (1-based). */
  startRow: number;
  startColumn: number;
}

export interface WorkbookSource {
  describe(): Promise<WorkbookDescription>;
  listSheets(): Promise<SheetInfo[]>;
  /** address omitted => used range of the sheet. */
  readRange(sheet: string, address?: string): Promise<Grid>;
  close(): Promise<void>;
}

export function flatten(grid: Grid): Cell[] {
  const out: Cell[] = [];
  for (const row of grid.rows) for (const cell of row) out.push(cell);
  return out;
}

export function isBlank(cell: Cell | undefined): boolean {
  return (
    cell === null ||
    cell === undefined ||
    (typeof cell === "string" && cell.trim() === "")
  );
}

/** Converts 0-based column index to a spreadsheets column letter (0 -> A). */
export function columnLetter(index: number): string {
  let n = index + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

export function toAddress(
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
): string {
  const top = columnLetter(startColumn) + String(startRow);
  const bottom =
    columnLetter(startColumn + Math.max(columnCount, 1) - 1) +
    String(startRow + Math.max(rowCount, 1) - 1);
  return `${top}:${bottom}`;
}
