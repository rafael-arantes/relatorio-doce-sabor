import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { profileColumns } from "../blocks.js";
import { guarded, markdown, ok, type Session } from "../session.js";
import type { Cell } from "../types.js";
import { columnLetter, isBlank } from "../types.js";
import { renderGrid } from "./render.js";

const sourceSchema = {
  fonte: z
    .enum(["graph", "graph_download", "local"])
    .optional()
    .describe("Override the configured source for this call"),
  arquivoLocal: z
    .string()
    .optional()
    .describe("Read a specific .xlsx path instead of the configured workbook"),
};

export function registerReadTools(server: McpServer, session: Session): void {
  server.registerTool(
    "excel_read_sheet",
    {
      title: "Read a worksheet",
      description:
        "Reads a worksheet's used range (or an explicit A1 range). format='records' gives a normal table " +
        "with headers, 'rows' a faithful raw grid, 'markdown' a rendered preview, 'csv' an export-ready " +
        "view. Defaults to 'records', capped by `limiteLinhas`.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name exactly as listed by excel_workbook_info"),
        intervalo: z
          .string()
          .optional()
          .describe("Optional A1 range such as 'A1:F50'. Defaults to the whole used range."),
        formato: z.enum(["records", "rows", "markdown", "csv"]).optional().describe("Default 'records'"),
        limiteLinhas: z.number().int().positive().optional().describe("Max rows returned (default 200)"),
        linhaCabecalho: z
          .number()
          .int()
          .optional()
          .describe("0-based header row index inside the range (default 0; -1 = no header)"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba, args.intervalo);
      const rendered = renderGrid(
        grid,
        args.formato ?? "records",
        args.limiteLinhas ?? 200,
        args.linhaCabecalho,
      );
      return rendered.text ? markdown(rendered.text, rendered.payload) : ok(rendered.payload);
    }),
  );

  server.registerTool(
    "excel_read_range",
    {
      title: "Read an A1 range",
      description:
        "Reads one explicit A1 range and returns the raw cell grid (plus a markdown preview when asked). " +
        "Use it after excel_detect_blocks to zoom into a specific block.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name"),
        intervalo: z.string().min(1).describe("A1 range, e.g. 'A1:H40'"),
        formato: z.enum(["rows", "records", "markdown", "csv"]).optional().describe("Default 'rows'"),
        limiteLinhas: z.number().int().positive().optional().describe("Max rows returned (default 200)"),
        linhaCabecalho: z.number().int().optional(),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba, args.intervalo);
      const rendered = renderGrid(grid, args.formato ?? "rows", args.limiteLinhas ?? 200, args.linhaCabecalho);
      return rendered.text ? markdown(rendered.text, rendered.payload) : ok(rendered.payload);
    }),
  );

  server.registerTool(
    "excel_column_profile",
    {
      title: "Profile a worksheet's columns",
      description:
        "Per-column statistics: filled/blank counts, and how many cells (and what share) parse as dates, " +
        "numbers or text, plus sample values. Use it to pick column mappings before calling " +
        "excel_build_report or excel_normalize_transactions.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1),
        linhaInicial: z
          .number()
          .int()
          .min(0)
          .optional()
          .describe("0-based row to start profiling from (skip title rows)"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const profiles = profileColumns(grid, args.linhaInicial ?? 0);
      return ok({
        aba: grid.sheet,
        endereco: grid.address,
        colunas: profiles.map((profile) => ({
          letra: profile.letter,
          indice: profile.column,
          preenchidas: profile.filled,
          vazias: profile.blanks,
          datas: `${profile.dateHits} (${Math.round(profile.dateShare * 100)}%)`,
          numeros: `${profile.numberHits} (${Math.round(profile.numberShare * 100)}%)`,
          textos: profile.textHits,
          valoresDistintos: profile.distinct,
          amostras: profile.samples,
        })),
      });
    }),
  );

  server.registerTool(
    "excel_search",
    {
      title: "Search cells",
      description:
        "Case- and accent-insensitive substring search across a worksheet. Returns each hit with its A1 " +
        "address and the full row it belongs to, for context. Handy to locate a supplier, a category or a " +
        "month heading inside a messy sheet.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1),
        termo: z.string().min(1),
        limite: z.number().int().positive().optional().describe("Max hits to return (default 50)"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const fold = (input: string) =>
        input
          .toLowerCase()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "");
      const needle = fold(args.termo);
      const limit = args.limite ?? 50;
      const hits: Array<{ celula: string; valor: Cell; linha: Cell[] }> = [];

      for (let r = 0; r < grid.rows.length && hits.length < limit; r++) {
        const row = grid.rows[r] ?? [];
        for (let c = 0; c < row.length && hits.length < limit; c++) {
          const cell = row[c];
          if (isBlank(cell)) continue;
          if (!fold(String(cell)).includes(needle)) continue;
          hits.push({
            celula: `${columnLetter(grid.startColumn + c)}${grid.startRow + r}`,
            valor: cell ?? null,
            linha: row,
          });
        }
      }

      return ok({ aba: grid.sheet, termo: args.termo, ocorrencias: hits.length, resultados: hits });
    }),
  );
}
