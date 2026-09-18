import { looksLikeDate, toNumber } from "./parse.js";
import type { Cell, Grid } from "./types.js";
import { columnLetter, isBlank } from "./types.js";

export interface Table {
  headers: string[];
  /** Data rows, aligned with `headers`. */
  rows: Cell[][];
  /** -1 when positional headers were generated (no usable header row). */
  headerRowIndex: number;
}

/** A header row is textual and has at least `minLabels` labels. */
export function looksLikeHeaderRow(row: Cell[], minLabels = 2): boolean {
  const filled = row.filter((cell) => !isBlank(cell));
  if (filled.length < minLabels) return false;
  const textual = filled.filter((cell) => {
    if (typeof cell !== "string") return false;
    if (toNumber(cell) !== null) return false;
    if (looksLikeDate(cell)) return false;
    return cell.trim().length > 0;
  });
  return textual.length / filled.length >= 0.6;
}

export function dedupeHeaders(names: string[]): string[] {
  const seen = new Map<string, number>();
  return names.map((name, index) => {
    const base = name.trim() || `coluna_${columnLetter(index)}`;
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

/**
 * Treats row `headerRowIndex` as the header row and everything below as data.
 * When the requested header row doesn't look like one, positional names
 * (coluna_A, coluna_B, ...) are generated and no row is consumed — so a messy
 * sheet that starts straight with data isn't silently missing its first row.
 */
export function toTable(grid: Grid, headerRowIndex = 0): Table {
  const width = grid.columnCount;
  let headerIndex = headerRowIndex;

  if (headerRowIndex === 0 && !looksLikeHeaderRow(grid.rows[0] ?? [])) {
    headerIndex = -1;
  } else if (headerRowIndex >= 0 && headerRowIndex >= grid.rows.length) {
    headerIndex = -1;
  }

  const headers =
    headerIndex >= 0
      ? dedupeHeaders(
          Array.from({ length: width }, (_, c) => {
            const cell = grid.rows[headerIndex]?.[c];
            return isBlank(cell) ? "" : String(cell);
          }),
        )
      : Array.from({ length: width }, (_, c) => `coluna_${columnLetter(c)}`);

  const dataStart = headerIndex >= 0 ? headerIndex + 1 : 0;
  const rows: Cell[][] = [];
  for (let r = dataStart; r < grid.rows.length; r++) {
    const row = grid.rows[r] ?? [];
    if (row.every((cell) => isBlank(cell))) continue;
    const padded = row.slice(0, width);
    while (padded.length < width) padded.push(null);
    rows.push(padded);
  }

  return { headers, rows, headerRowIndex: headerIndex };
}

/** Table rows as objects keyed by header name. */
export function toRecords(table: Table): Array<Record<string, Cell>> {
  return table.rows.map((row) => {
    const record: Record<string, Cell> = {};
    table.headers.forEach((header, index) => {
      record[header] = row[index] ?? null;
    });
    return record;
  });
}

/** Finds a header by fuzzy match (accent/case insensitive, substring aware). */
export function findHeader(headers: string[], candidates: string[]): string | null {
  const normalized = headers.map((h) => ({ raw: h, key: h.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "") }));
  for (const candidate of candidates) {
    const exact = normalized.find((h) => h.key === candidate.toLowerCase());
    if (exact) return exact.raw;
  }
  for (const candidate of candidates) {
    const partial = normalized.find((h) => h.key.includes(candidate.toLowerCase()));
    if (partial) return partial.raw;
  }
  return null;
}
