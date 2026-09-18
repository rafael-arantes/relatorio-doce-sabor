import { gridAddress } from "./cells.js";
import { looksLikeDate, toNumber } from "./parse.js";
import type { Cell, Grid } from "./types.js";
import { columnLetter, isBlank } from "./types.js";

export interface Block {
  index: number;
  address: string;
  /** 0-based offsets inside the grid. */
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  cellCount: number;
  /** Title cell when the block starts with a lone label ("MERCADO"). */
  label: string | null;
  headerRow: Cell[];
  /** Offset of the header row inside the block (0, or 1 when a label exists). */
  headerRowOffset: number;
  dataRowCount: number;
}

export interface DetectBlocksOptions {
  /** Drop regions smaller than this many cells. Default 2. */
  minCells?: number;
  /** Merge blocks separated by up to N blank rows when they overlap horizontally. */
  mergeGapRows?: number;
  maxBlocks?: number;
}

interface Box {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
  cells: number;
}

/**
 * Finds the disconnected regions of a sheet — the "horizontal blocks" pattern
 * used by hand-built spreadsheets, where every supplier gets its own little
 * table with its own Data/Valor columns.
 *
 * Uses 4-connected flood fill over non-empty cells, so it copes with tables
 * stacked vertically, placed side by side, or both.
 */
export function detectBlocks(
  grid: Grid,
  options: DetectBlocksOptions = {},
): { blocks: Block[]; truncated: boolean } {
  const { minCells = 2, mergeGapRows = 0, maxBlocks = 200 } = options;
  const rows = grid.rows.length;
  const cols = grid.columnCount;

  const visited = new Uint8Array(rows * cols);
  const filled = (r: number, c: number): boolean => !isBlank(grid.rows[r]?.[c]);
  const boxes: Box[] = [];

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (visited[r * cols + c] || !filled(r, c)) continue;

      const box: Box = { r0: r, r1: r, c0: c, c1: c, cells: 0 };
      const stack: Array<[number, number]> = [[r, c]];
      visited[r * cols + c] = 1;

      while (stack.length > 0) {
        const [cr, cc] = stack.pop()!;
        box.cells++;
        if (cr < box.r0) box.r0 = cr;
        if (cr > box.r1) box.r1 = cr;
        if (cc < box.c0) box.c0 = cc;
        if (cc > box.c1) box.c1 = cc;

        const neighbours: Array<[number, number]> = [
          [cr - 1, cc],
          [cr + 1, cc],
          [cr, cc - 1],
          [cr, cc + 1],
        ];
        for (const [nr, nc] of neighbours) {
          if (nr < 0 || nc < 0 || nr >= rows || nc >= cols) continue;
          if (visited[nr * cols + nc] || !filled(nr, nc)) continue;
          visited[nr * cols + nc] = 1;
          stack.push([nr, nc]);
        }
      }

      if (box.cells >= minCells) boxes.push(box);
    }
  }

  // Optional vertical merge for one table broken up by spacer rows.
  const merged: Box[] = [];
  for (const box of boxes.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0)) {
    const host = merged.find(
      (m) => box.r0 > m.r1 && box.r0 - m.r1 - 1 <= mergeGapRows && !(box.c1 < m.c0 || box.c0 > m.c1),
    );
    if (host) {
      host.r1 = Math.max(host.r1, box.r1);
      host.c0 = Math.min(host.c0, box.c0);
      host.c1 = Math.max(host.c1, box.c1);
      host.cells += box.cells;
    } else {
      merged.push({ ...box });
    }
  }

  const truncated = merged.length > maxBlocks;
  const selected = merged
    .sort((a, b) => b.cells - a.cells)
    .slice(0, maxBlocks)
    .sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);

  const blocks = selected.map((box, index) => {
    const rowCount = box.r1 - box.r0 + 1;
    const columnCount = box.c1 - box.c0 + 1;
    const slice = grid.rows
      .slice(box.r0, box.r1 + 1)
      .map((row) => row.slice(box.c0, box.c1 + 1).map((cell) => cell ?? null));

    const firstRow = slice[0] ?? [];
    const firstRowFilled = firstRow.filter((cell) => !isBlank(cell));
    let label: string | null = null;
    let headerRowOffset = 0;

    // A lone title cell on its own row => block label; header row moves down.
    if (
      rowCount > 1 &&
      firstRowFilled.length === 1 &&
      typeof firstRowFilled[0] === "string" &&
      firstRowFilled[0].trim().split(/\s+/).length <= 4
    ) {
      label = String(firstRowFilled[0]).trim();
      headerRowOffset = 1;
    }

    return {
      index,
      address: gridAddress(grid, box.r0, box.c0, rowCount, columnCount),
      startRow: box.r0,
      startColumn: box.c0,
      rowCount,
      columnCount,
      cellCount: box.cells,
      label,
      headerRow: slice[headerRowOffset] ?? [],
      headerRowOffset,
      dataRowCount: Math.max(rowCount - headerRowOffset - (label ? 1 : 0), 0),
    };
  });

  return { blocks, truncated };
}

/** Column-level profile, used for normalization and for sheet discovery. */
export interface ColumnProfile {
  column: number;
  letter: string;
  filled: number;
  blanks: number;
  dateHits: number;
  numberHits: number;
  textHits: number;
  dateShare: number;
  numberShare: number;
  distinct: number;
  samples: Cell[];
}

export function profileColumns(grid: Grid, startRow = 0): ColumnProfile[] {
  const profiles: ColumnProfile[] = [];

  for (let c = 0; c < grid.columnCount; c++) {
    const samples: Cell[] = [];
    const distinct = new Set<string>();
    let filled = 0;
    let blanks = 0;
    let dateHits = 0;
    let numberHits = 0;
    let textHits = 0;

    for (let r = startRow; r < grid.rows.length; r++) {
      const cell = grid.rows[r]?.[c] ?? null;
      if (isBlank(cell)) {
        blanks++;
        continue;
      }
      filled++;
      if (looksLikeDate(cell)) dateHits++;
      if (toNumber(cell) !== null) numberHits++;
      if (typeof cell === "string" && cell.trim().length >= 3) textHits++;
      if (samples.length < 5) samples.push(cell);
      distinct.add(String(cell));
    }

    profiles.push({
      column: c,
      letter: columnLetter(c),
      filled,
      blanks,
      dateHits,
      numberHits,
      textHits,
      dateShare: filled ? dateHits / filled : 0,
      numberShare: filled ? numberHits / filled : 0,
      distinct: distinct.size,
      samples,
    });
  }

  return profiles;
}

