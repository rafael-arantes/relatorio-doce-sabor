import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { detectBlocks } from "../blocks.js";
import { guarded, ok, type Session } from "../session.js";

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

export function registerWorkbookTools(server: McpServer, session: Session): void {
  server.registerTool(
    "excel_workbook_info",
    {
      title: "Workbook info",
      description:
        "Describes the connected workbook (file name, where it came from, last modified, web link) and " +
        "lists every worksheet with its dimensions. Start here — it tells you which sheet names to pass to " +
        "the other tools.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {},
    },
    guarded(async () => {
      const problems = session.problems();
      if (problems.length > 0) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text:
                "The workbook source is not configured yet:\n" +
                problems.map((problem) => `- ${problem}`).join("\n") +
                "\n\nSee README.md for the setup walkthrough (or run `npm run doctor` in the excel-mcp folder).",
            },
          ],
        };
      }
      const source = await session.getSource();
      return ok(await source.describe());
    }),
  );

  server.registerTool(
    "excel_diagnostics",
    {
      title: "Configuration diagnostics",
      description:
        "Reports how this MCP server is configured (source mode, auth mode, target workbook), its limits, " +
        "and any setup problem it detects — without calling Microsoft. Use it when another tool fails with " +
        "an auth or file-not-found error.",
      annotations: { readOnlyHint: true, openWorldHint: false },
      inputSchema: {},
    },
    guarded(async () => {
      const config = session.config;
      return ok({
        fonte: config.source,
        autenticacao: config.source === "graph" ? config.auth : "n/a",
        tenant: config.source === "graph" ? config.tenantId : "n/a",
        clientId: config.source === "graph" ? `${config.clientId.slice(0, 8)}…` : "n/a",
        alvo: {
          workbookPath: config.workbookPath || null,
          driveId: config.driveId || null,
          itemId: config.itemId || null,
          shareUrl: config.shareUrl ? `${config.shareUrl.slice(0, 40)}…` : null,
          arquivoLocal: config.localFile || null,
        },
        fallbackDeDownload: config.graphFallback
          ? "ligado — se a API de workbook for recusada, o arquivo é baixado e lido localmente"
          : "desligado (EXCEL_GRAPH_FALLBACK=0)",
        cacheDeToken: config.tokenCache,
        limites: { maxRows: config.maxRows, maxCells: config.maxCells },
        problemas: session.problems(),
        dica:
          config.source === "graph" && config.auth === "device_code"
            ? "Se aparecer erro de credencial, rode `npm run login` (uma vez) nesta pasta."
            : null,
      });
    }),
  );

  server.registerTool(
    "excel_detect_blocks",
    {
      title: "Detect blocks in a worksheet",
      description:
        "Finds the disconnected 'blocks' of a hand-built spreadsheet — each supplier/expense group with its " +
        "own little table and its own Data/Valor columns. Returns every block's address, title label, header " +
        "row and data-row count. Use this first when a sheet has no single clean table.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name"),
        minCelulas: z.number().int().min(1).optional().describe("Ignore regions smaller than this (default 2)"),
        mesclarLinhas: z
          .number()
          .int()
          .min(0)
          .optional()
          .describe("Merge blocks separated by up to N blank rows (default 0)"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const { blocks, truncated } = detectBlocks(grid, {
        minCells: args.minCelulas ?? 2,
        mergeGapRows: args.mesclarLinhas ?? 0,
      });
      return ok({
        aba: grid.sheet,
        endereco: grid.address,
        totalBlocos: blocks.length,
        truncado: truncated,
        blocos: blocks.map((block) => ({
          endereco: block.address,
          rotulo: block.label,
          linhas: block.rowCount,
          colunas: block.columnCount,
          celulasPreenchidas: block.cellCount,
          cabecalho: block.headerRow,
          linhasDeDados: block.dataRowCount,
        })),
      });
    }),
  );
}
