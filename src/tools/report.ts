import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { normalizeTransactions } from "../normalize.js";
import { buildReport, metricName } from "../report.js";
import { toRecords, toTable } from "../tables.js";
import { guarded, markdown, ok, type Session } from "../session.js";
import type { Cell, Grid } from "../types.js";
import { aggs, filterOps, filterSchema, metricSchema, sourceSchema } from "./schemas.js";

/** Maps normalized transactions into report-friendly records. */
export function transactionRecords(
  grid: Grid,
  options: Parameters<typeof normalizeTransactions>[1],
): Array<Record<string, Cell>> {
  return normalizeTransactions(grid, options).transacoes.map((t) => ({
    data: t.data,
    mes: t.data ? t.data.slice(0, 7) : null,
    tipo: t.tipo,
    categoria: t.categoria,
    descricao: t.descricao,
    valor: t.valor,
    conta: t.conta,
    bloco: t.origem.bloco,
    celula: t.origem.endereco,
    confianca: t.confianca,
  }));
}

export function registerNormalizeTools(server: McpServer, session: Session): void {
  server.registerTool(
    "excel_normalize_transactions",
    {
      title: "Normalize a messy sheet into transactions",
      description:
        "Turns any worksheet — including sheets built from many disconnected blocks — into clean transaction " +
        "rows shaped like the target schema: data, tipo, categoria, descricao, valor, conta. Payment methods " +
        "buried in the item label ('ovo caipira (dinheiro)') are pulled into `conta` and stripped from the " +
        "description. Every row keeps its provenance (block + cell address) and a confidence level, and the " +
        "response includes per-block diagnostics so nothing is a black box.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name"),
        ano: z
          .number()
          .int()
          .optional()
          .describe("Year for dates that only carry day+month (default: inferred from the sheet name)"),
        colunaData: z.number().int().optional().describe("Force data column (0-based, inside each block)"),
        colunaValor: z.number().int().optional(),
        colunaDescricao: z.number().int().optional(),
        colunaConta: z.number().int().optional(),
        colunaCategoria: z.number().int().optional(),
        colunaTipo: z.number().int().optional(),
        linhaCabecalho: z.number().int().optional().describe("Force header row index inside each block"),
        valorMinimo: z.number().optional().describe("Ignore amounts below this (default 0.01)"),
        apenasBlocos: z.string().optional().describe("Only normalize blocks whose label matches this regex"),
        limiteTransacoes: z.number().int().positive().optional().describe("Hard cap (default 5000)"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const result = normalizeTransactions(grid, {
        aba: args.aba,
        ano: args.ano,
        colunaData: args.colunaData,
        colunaValor: args.colunaValor,
        colunaDescricao: args.colunaDescricao,
        colunaConta: args.colunaConta,
        colunaCategoria: args.colunaCategoria,
        colunaTipo: args.colunaTipo,
        headerRowIndex: args.linhaCabecalho,
        valorMinimo: args.valorMinimo,
        apenasBlocos: args.apenasBlocos,
        limiteLinhas: args.limiteTransacoes,
      });

      const total = result.transacoes.reduce((sum, t) => sum + t.valor, 0);
      const porConta = new Map<string, number>();
      const porCategoria = new Map<string, number>();
      for (const t of result.transacoes) {
        const contaKey = t.conta ?? "(sem conta)";
        porConta.set(contaKey, (porConta.get(contaKey) ?? 0) + t.valor);
        const catKey = t.categoria ?? "(sem categoria)";
        porCategoria.set(catKey, (porCategoria.get(catKey) ?? 0) + t.valor);
      }

      const text =
        `**${result.transacoes.length} transações** normalizadas de \`${grid.sheet}\` — total ` +
        `R$ ${total.toFixed(2)}. Blocos: ${result.blocos.length} · linhas ignoradas: ${result.linhasIgnoradas}` +
        (result.truncado ? " · ⚠️ truncado" : "") +
        (result.avisos.length > 0 ? `\n\n⚠️ ${result.avisos.join("\n⚠️ ")}` : "");

      return markdown(text, {
        aba: grid.sheet,
        endereco: grid.address,
        transacoes: result.transacoes.map((t) => ({
          data: t.data,
          tipo: t.tipo,
          categoria: t.categoria,
          descricao: t.descricao,
          valor: t.valor,
          conta: t.conta,
          confianca: t.confianca,
          origem: t.origem,
        })),
        resumo: {
          quantidade: result.transacoes.length,
          valorTotal: Math.round(total * 100) / 100,
          semData: result.transacoes.filter((t) => t.data === null).length,
          semConta: result.transacoes.filter((t) => t.conta === null).length,
          porConta: Object.fromEntries(porConta),
          porCategoria: Object.fromEntries(porCategoria),
        },
        blocos: result.blocos,
        linhasIgnoradas: result.linhasIgnoradas,
        avisos: result.avisos,
        truncado: result.truncado,
      });
    }),
  );

  server.registerTool(
    "excel_build_report",
    {
      title: "Build a custom report",
      description:
        "The reporting workhorse: filters rows, groups them and computes metrics (sum/avg/count/" +
        "count_distinct/min/max/median/first/last). Returns a markdown table, the group rows, grand totals " +
        "and a share-of-total column. Two modes: 'transacoes' (default — normalizes the sheet into " +
        "Data/Tipo/Categoria/Descrição/Valor/Conta first) or 'bruto' (the sheet's own columns).",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: {
        aba: z.string().min(1).describe("Worksheet name"),
        modo: z
          .enum(["transacoes", "bruto"])
          .optional()
          .describe("'transacoes' normalizes first (default), 'bruto' uses the sheet's own columns"),
        agruparPor: z.array(z.string()).optional().describe("Columns to group by, e.g. ['categoria']"),
        filtros: z.array(filterSchema).optional(),
        metricas: z
          .array(metricSchema)
          .optional()
          .describe("Defaults to count + sum of the detected amount column"),
        ordenarPor: z.string().optional().describe("Output column to sort by (defaults to the first metric)"),
        ordemDesc: z.boolean().optional().describe("Sort descending (default false)"),
        limiteGrupos: z.number().int().positive().optional().describe("Max groups returned"),
        linhaCabecalho: z.number().int().optional(),
        ano: z.number().int().optional().describe("Year hint for normalized dates"),
        ...sourceSchema,
      },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const modo = args.modo ?? "transacoes";

      const records: Array<Record<string, Cell>> =
        modo === "transacoes"
          ? transactionRecords(grid, { aba: args.aba, ano: args.ano })
          : toRecords(toTable(grid, args.linhaCabecalho ?? 0));

      const report = buildReport(records, {
        groupBy: args.agruparPor,
        filters: (args.filtros ?? []) as never,
        metrics: args.metricas as never,
        sortBy: args.ordenarPor,
        sortDesc: args.ordemDesc,
        limit: args.limiteGrupos,
      });

      const text =
        `**Relatório — ${grid.sheet}** (modo \`${modo}\`)  \n` +
        `Linhas consideradas: ${report.registrosFiltrados} de ${report.registrosTotais}\n\n` +
        report.markdown +
        (report.avisos.length > 0 ? `\n\n⚠️ ${report.avisos.join("\n⚠️ ")}` : "");

      return markdown(text, {
        colunas: report.colunas,
        linhas: report.linhas,
        totais: report.totais,
        registrosTotais: report.registrosTotais,
        registrosFiltrados: report.registrosFiltrados,
        avisos: report.avisos,
      });
    }),
  );

  server.registerTool(
    "excel_list_metrics",
    {
      title: "List report columns and metric names",
      description:
        "Explains which column names, filter operators and aggregations excel_build_report accepts for a " +
        "sheet — the normalized transaction columns plus, in 'bruto' mode, the sheet's own headers. Call it " +
        "when unsure how to phrase a report.",
      annotations: { readOnlyHint: true, openWorldHint: true },
      inputSchema: { aba: z.string().min(1), ...sourceSchema },
    },
    guarded(async (args) => {
      const source = await session.getSource({ source: args.fonte, arquivoLocal: args.arquivoLocal });
      const grid = await source.readRange(args.aba);
      const table = toTable(grid);

      return ok({
        aba: grid.sheet,
        modoTransacoes: {
          colunas: [
            "data",
            "mes",
            "tipo",
            "categoria",
            "descricao",
            "valor",
            "conta",
            "bloco",
            "celula",
            "confianca",
          ],
          operadoresDeFiltro: filterOps,
          agregacoes: aggs,
          exemploMetrica: metricName({ agg: "sum", column: "valor", as: "total" }),
        },
        modoBruto: { colunas: table.headers },
        exemplos: [
          { agruparPor: ["categoria"], metricas: [{ agg: "sum", column: "valor", as: "total" }] },
          {
            filtros: [{ column: "mes", op: "eq", value: "2026-09" }],
            agruparPor: ["conta"],
            metricas: [
              { agg: "sum", column: "valor", as: "total" },
              { agg: "count", as: "lancamentos" },
            ],
          },
          {
            filtros: [{ column: "descricao", op: "icontains", value: "ambev" }],
            agruparPor: ["mes"],
          },
        ],
      });
    }),
  );
}
