import { z } from "zod";

export const filterOps = [
  "eq",
  "ne",
  "contains",
  "icontains",
  "startswith",
  "gt",
  "gte",
  "lt",
  "lte",
  "in",
  "not_in",
  "between",
  "is_empty",
  "not_empty",
  "regex",
  "month",
] as const;

export const aggs = [
  "sum",
  "avg",
  "count",
  "count_distinct",
  "min",
  "max",
  "median",
  "first",
  "last",
] as const;

export const sourceSchema = {
  fonte: z
    .enum(["graph", "graph_download", "local"])
    .optional()
    .describe(
      "Override the configured source for this call: 'graph' (workbook API), " +
        "'graph_download' (download the file and parse locally), 'local' (file on disk)",
    ),
  arquivoLocal: z
    .string()
    .optional()
    .describe("Read a specific .xlsx path instead of the configured workbook"),
};

export const filterSchema = z.object({
  column: z
    .string()
    .min(1)
    .describe("Column name; use 'mes' for YYYY-MM matching on the normalized rows"),
  op: z.enum(filterOps),
  value: z
    .union([z.string(), z.number(), z.boolean(), z.array(z.union([z.string(), z.number()]))])
    .optional()
    .describe("Comparison value; array for in/not_in/between, 'YYYY-MM' for op='month'"),
});

export const metricSchema = z.object({
  column: z.string().optional().describe("Omit for plain counting"),
  agg: z.enum(aggs),
  as: z.string().optional().describe("Output column name"),
});

export const headerRowSchema = {
  linhaCabecalho: z
    .number()
    .int()
    .optional()
    .describe("0-based header row index (default 0 auto-detects; -1 = no header)"),
};
