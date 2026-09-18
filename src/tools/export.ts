import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { toCsv } from "../report.js";
import { toRecords, toTable } from "../tables.js";
import { guarded, ok, type Session } from "../session.js";
import type { Cell } from "../types.js";
import { sourceSchema } from "./schemas.js";
import { transactionRecords } from "./report.js";

export function registerExportTools(server: McpServer, session: Session): void {
  server.registerTool(
    "excel_export_csv",
    {
      title: "Export rows to CSV",
      description:
        "Writes either the normalized transactions or the raw sheet rows to a CSV file on disk — pt-BR " +
        "friendly (`;` separator, comma decimals, UTF-8 BOM so Excel opens it correctly). Use it to hand a " +
        "clean extract to another tool, archive a month, or prep a spreadsheet import.",
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name"),
        caminho: z.string().min(1).describe("Path of the .csv to write (absolute is safest)"),
        modo: z.enum(["transacoes", "bruto"]).optional().describe("Default 'transacoes'"),
        ano: z.number().int().optional().describe("Year hint for normalized dates"),
        delimitador: z.string().optional().describe("Defaults to ';'"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const modo = args.modo ?? "transacoes";

      let columns: string[];
      let rows: Array<Record<string, Cell>>;

      if (modo === "transacoes") {
        columns = ["data", "tipo", "categoria", "descricao", "valor", "conta", "confianca", "bloco", "celula"];
        rows = transactionRecords(grid, { aba: args.aba, ano: args.ano });
      } else {
        const table = toTable(grid);
        columns = table.headers;
        rows = toRecords(table);
      }

      const target = resolve(args.caminho);
      mkdirSync(dirname(target), { recursive: true });
      writeFileSync(target, toCsv(columns, rows, { delimiter: args.delimitador ?? ";" }), "utf8");

      return ok({
        arquivo: target,
        linhas: rows.length,
        colunas: columns,
        mensagem: `${rows.length} linhas gravadas em ${target}`,
      });
    }),
  );
}
