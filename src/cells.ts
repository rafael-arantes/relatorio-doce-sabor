import { toISODate } from "./parse.js";
import type { Cell, Grid } from "./types.js";
import { columnLetter, isBlank, toAddress } from "./types.js";

/** Normalizes a raw value coming from Graph (JSON) or exceljs into a flat Cell. */
export function normalizeCell(value: unknown): Cell {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value;
  if (typeof value === "string") return value.trim() === "" ? null : value;
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value.toISOString().slice(0, 10);
  }
  if (typeof value === "object") {
    const obj = value as Record<string, unknown>;
    // Graph error cells: {"error":"#DIV/0!"}
    if (typeof obj.error === "string") return obj.error;
    // Rich text / hyperlink cells from exceljs.
    if (typeof obj.text === "string") return obj.text;
    if (obj.result !== undefined) return normalizeCell(obj.result);
    if (Array.isArray(obj.richText)) {
      const joined = obj.richText
        .map((part) => (part && typeof part === "object" ? String((part as { text?: unknown }).text ?? "") : ""))
        .join("");
      return joined.trim() === "" ? null : joined;
    }
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }
  return String(value);
}

/**
 * Picks the most report-friendly representation of a cell.
 *
 * Graph returns `values` (raw) and `text` (as displayed). Raw numbers keep full
 * precision, but a date cell's raw form is an opaque serial number — so when the
 * displayed text is a date, the displayed text wins. Everything else prefers the
 * raw value, falling back to text when the raw value is empty.
 */
export function pickCell(raw: Cell | undefined, text: Cell | undefined): Cell {
  const rawBlank = raw === null || raw === undefined;
  const textBlank = text === null || text === undefined;

  if (!textBlank && typeof text === "string" && toISODate(text) !== null) return text;
  if (!rawBlank) return raw;
  if (!textBlank) return text;
  return null;
}

export function mergeMatrices(values: unknown[][], text?: unknown[][]): Cell[][] {
  const rowCount = Math.max(values.length, text?.length ?? 0);
  const out: Cell[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const valueRow = values[r] ?? [];
    const textRow = text?.[r] ?? [];
    const width = Math.max(valueRow.length, textRow.length);
    const row: Cell[] = [];
    for (let c = 0; c < width; c++) {
      row.push(pickCell(normalizeCell(valueRow[c]), normalizeCell(textRow[c])));
    }
    out.push(row);
  }
  return out;
}

/** Pads every row to the widest row so index math is predictable. */
export function rectangularize(rows: Cell[][]): Grid["rows"] {
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return rows.map((row) => {
    const copy = row.slice();
    while (copy.length < width) copy.push(null);
    return copy;
  });
}

export function makeGrid(
  sheet: string,
  rows: Cell[][],
  truncated: boolean,
  startRow = 1,
  startColumn = 0,
): Grid {
  const padded = rectangularize(rows);
  const rowCount = padded.length;
  const columnCount = padded[0]?.length ?? 0;
  return {
    sheet,
    address: toAddress(startRow, startColumn, Math.max(rowCount, 1), Math.max(columnCount, 1)),
    rows: padded,
    rowCount,
    columnCount,
    truncated,
    startRow,
    startColumn,
  };
}

/**
 * Drops leading/trailing blank rows and columns.
 *
 * Excel's "used range" is often much bigger than the real data (a stray format,
 * a deleted row), which both bloats responses and confuses block detection.
 * Addresses stay correct because `startRow`/`startColumn` are shifted.
 */
export function trimEmptyEdges(grid: Grid): Grid {
  const rows = grid.rows;
  const rowBlank = (row: Cell[] | undefined): boolean => (row ?? []).every((cell) => isBlank(cell));

  let top = 0;
  while (top < rows.length && rowBlank(rows[top])) top++;
  let bottom = rows.length - 1;
  while (bottom >= top && rowBlank(rows[bottom])) bottom--;
  if (bottom < top) {
    return { ...grid, rows: [], rowCount: 0, columnCount: 0 };
  }

  const slice = rows.slice(top, bottom + 1);
  const columnBlank = (column: number): boolean => slice.every((row) => isBlank(row[column]));

  let left = 0;
  let right = grid.columnCount - 1;
  while (left <= right && columnBlank(left)) left++;
  while (right >= left && columnBlank(right)) right--;

  const trimmed = slice.map((row) => row.slice(left, right + 1));
  return makeGrid(grid.sheet, trimmed, grid.truncated, grid.startRow + top, grid.startColumn + left);
}

/** Human-readable A1 address of a sub-rectangle inside a grid. */
export function gridAddress(
  grid: Grid,
  rowIndex: number,
  columnIndex: number,
  rowSpan = 1,
  columnSpan = 1,
): string {
  const startRow = grid.startRow + rowIndex;
  const startColumn = grid.startColumn + columnIndex;
  const endRow = startRow + Math.max(rowSpan, 1) - 1;
  const endColumn = startColumn + Math.max(columnSpan, 1) - 1;
  return `${columnLetter(startColumn)}${startRow}:${columnLetter(endColumn)}${endRow}`;
}
