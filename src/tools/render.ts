import { toCsv } from "../report.js";
import { toRecords, toTable } from "../tables.js";
import type { Cell, Grid } from "../types.js";
import { isBlank } from "../types.js";

export type OutputFormat = "rows" | "records" | "csv" | "markdown";

function gridToMarkdown(grid: Grid, limit: number): string {
  const rows = grid.rows.slice(0, limit);
  if (rows.length === 0) return "_Nenhum dado nesta aba._";
  const cell = (value: Cell | undefined): string =>
    isBlank(value) ? "" : String(value).replace(/\|/g, "\\|");
  const header = `| ${(rows[0] ?? []).map(cell).join(" | ")} |`;
  const separator = `| ${Array.from({ length: grid.columnCount }, () => "---").join(" | ")} |`;
  const body = rows.slice(1).map((row) => `| ${row.map(cell).join(" | ")} |`);
  return [header, separator, ...body].join("\n");
}

/**
 * Renders a grid in whichever shape the caller asked for and keeps the payload
 * small: `limiteLinhas` caps what is returned while the totals still report the
 * real size of the sheet.
 */
export function renderGrid(
  grid: Grid,
  format: OutputFormat,
  limit: number,
  headerRow?: number,
): { payload: unknown; text?: string } {
  if (format === "markdown") {
    return { payload: { aba: grid.sheet, endereco: grid.address }, text: gridToMarkdown(grid, limit) };
  }

  if (format === "csv") {
    const table = toTable(grid, headerRow ?? 0);
    const records = toRecords(table);
    return {
      payload: { aba: grid.sheet, endereco: grid.address },
      text: `\`\`\`csv\n${toCsv(table.headers, records)}\`\`\``,
    };
  }

  if (format === "records") {
    const table = toTable(grid, headerRow ?? 0);
    const records = toRecords(table);
    return {
      payload: {
        aba: grid.sheet,
        endereco: grid.address,
        linhaCabecalho: table.headerRowIndex,
        colunas: table.headers,
        linhasTotais: records.length,
        linhas: records.slice(0, limit),
        truncado: grid.truncated || records.length > limit,
      },
    };
  }

  return {
    payload: {
      aba: grid.sheet,
      endereco: grid.address,
      linhasTotais: grid.rows.length,
      colunasTotais: grid.columnCount,
      linhas: grid.rows.slice(0, limit),
      truncado: grid.truncated || grid.rows.length > limit,
    },
  };
}
