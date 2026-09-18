import type { Cell } from "./types.js";
import { isBlank } from "./types.js";
import {
  aggregate,
  formatNumber,
  matchesFilter,
  type Metric,
  type ReportResult,
  type ReportRow,
  type ReportSpec,
} from "./report-filters.js";

const VALUE_HEADER_HINTS = ["valor", "total", "preco", "montante", "vlr", "custo", "gasto", "quantia"];

function normalizeHeader(name: string): string {
  return name.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/** Picks the most plausible amount column when the caller didn't say. */
export function autoValueColumn(records: Array<Record<string, Cell>>): string | null {
  if (records.length === 0) return null;
  const headers = Object.keys(records[0] ?? {});

  const scored = headers
    .map((header) => {
      let numeric = 0;
      let filled = 0;
      for (const record of records) {
        const value = record[header];
        if (isBlank(value)) continue;
        filled++;
        if (typeof value === "number") numeric++;
      }
      return { header, filled, numeric };
    })
    .filter(
      (entry) => entry.filled >= Math.max(1, Math.floor(records.length * 0.5)) && entry.numeric > 0,
    );

  const hinted = scored.find((entry) =>
    VALUE_HEADER_HINTS.some((hint) => normalizeHeader(entry.header).includes(hint)),
  );
  if (hinted) return hinted.header;

  // Otherwise the column that is numeric for the largest share of its values.
  const best = scored.sort((a, b) => b.numeric / b.filled - a.numeric / a.filled)[0];
  return best?.header ?? null;
}

function defaultMetrics(records: Array<Record<string, Cell>>): Metric[] {
  const metrics: Metric[] = [{ agg: "count", as: "registros" }];
  const column = autoValueColumn(records);
  if (column) metrics.push({ agg: "sum", column, as: `total_${column}` });
  return metrics;
}

export function metricName(metric: Metric): string {
  return metric.as ?? (metric.column ? `${metric.agg}_${metric.column}` : metric.agg);
}

/**
 * Runs a pivot-style report over any set of records: filter, group, aggregate,
 * sort, and render as a markdown table plus the raw rows.
 */
export function buildReport(
  records: Array<Record<string, Cell>>,
  spec: ReportSpec = {},
): ReportResult {
  const avisos: string[] = [];
  const groupBy = spec.groupBy ?? [];
  const filters = spec.filters ?? [];

  const headers = records.length > 0 ? Object.keys(records[0]!) : [];
  for (const filter of filters) {
    if (!headers.includes(filter.column)) {
      avisos.push(
        `Filtro ignorado: a coluna "${filter.column}" não existe. Colunas disponíveis: ${headers.join(", ")}`,
      );
    }
  }
  for (const column of groupBy) {
    if (!headers.includes(column)) avisos.push(`Agrupamento ignorado: a coluna "${column}" não existe.`);
  }

  const usableFilters = filters.filter((f) => headers.includes(f.column));
  const filtered = records.filter((record) => usableFilters.every((f) => matchesFilter(record, f)));
  const metrics = spec.metrics && spec.metrics.length > 0 ? spec.metrics : defaultMetrics(filtered);

  for (const metric of metrics) {
    if (metric.column && !headers.includes(metric.column)) {
      avisos.push(`Métrica ignorada: a coluna "${metric.column}" não existe.`);
    }
  }
  const usableMetrics = metrics.filter(
    (m) => m.agg === "count" || (m.column !== undefined && headers.includes(m.column)),
  );
  const usableGroupBy = groupBy.filter((column) => headers.includes(column));

  const buckets = new Map<string, Array<Record<string, Cell>>>();
  if (usableGroupBy.length === 0) {
    buckets.set("__todos__", filtered);
  } else {
    for (const record of filtered) {
      const key = usableGroupBy
        .map((column) => (isBlank(record[column]) ? "(vazio)" : String(record[column])))
        .join(" | ");
      const bucket = buckets.get(key);
      if (bucket) bucket.push(record);
      else buckets.set(key, [record]);
    }
  }

  const buildRow = (rows: Array<Record<string, Cell>>, key: string): ReportRow => {
    const out: ReportRow = {};
    if (usableGroupBy.length > 0) {
      const parts = key.split(" | ");
      usableGroupBy.forEach((column, index) => {
        out[column] = parts[index] ?? null;
      });
    }
    for (const metric of usableMetrics) {
      // A metric without a column counts rows rather than reading a cell.
      const values: Cell[] =
        metric.column === undefined ? rows.map(() => 1) : rows.map((row) => row[metric.column!] ?? null);
      out[metricName(metric)] = aggregate(values, metric.agg);
    }
    return out;
  };

  const totais = buildRow(filtered, "__todos__");
  let linhas = [...buckets.entries()].map(([key, rows]) => buildRow(rows, key));

  // Share-of-total for sum metrics makes reports readable without extra math.
  const shareMetric = usableMetrics.find((m) => m.agg === "sum");
  if (shareMetric && linhas.length > 1) {
    const totalValue = totais[metricName(shareMetric)];
    const shareColumn = `${metricName(shareMetric)}_%`;
    for (const row of linhas) {
      const value = row[metricName(shareMetric)];
      row[shareColumn] =
        typeof value === "number" && typeof totalValue === "number" && totalValue !== 0
          ? Math.round((value / totalValue) * 10_000) / 100
          : null;
    }
  }

  const sortKey = spec.sortBy ?? usableMetrics[0]?.as ?? "";
  if (sortKey && linhas.length > 0 && sortKey in linhas[0]!) {
    const direction = spec.sortDesc ? -1 : 1;
    linhas = linhas.sort((a, b) => {
      const av = a[sortKey];
      const bv = b[sortKey];
      if (typeof av === "number" && typeof bv === "number") return (av - bv) * direction;
      return String(av ?? "").localeCompare(String(bv ?? ""), "pt-BR") * direction;
    });
  }
  if (spec.limit && linhas.length > spec.limit) linhas = linhas.slice(0, spec.limit);

  const colunas = linhas.length > 0 ? Object.keys(linhas[0]!) : Object.keys(totais);
  const markdown = renderMarkdown(colunas, linhas, totais, usableGroupBy.length > 0);

  return {
    colunas,
    linhas,
    totais,
    registrosTotais: records.length,
    registrosFiltrados: filtered.length,
    markdown,
    avisos,
  };
}

export function renderMarkdown(
  colunas: string[],
  linhas: ReportRow[],
  totais: ReportRow | null,
  includeTotals: boolean,
): string {
  if (colunas.length === 0) return "_Sem dados para os filtros informados._";

  const renderCell = (value: string | number | null, isPercent = false): string => {
    if (value === null) return "";
    if (typeof value === "number") return isPercent ? `${formatNumber(value, 2)}%` : formatNumber(value);
    return value.replace(/\|/g, "\\|");
  };

  const header = `| ${colunas.join(" | ")} |`;
  const separator = `| ${colunas.map(() => "---").join(" | ")} |`;
  const body = linhas
    .map(
      (row) =>
        `| ${colunas.map((column) => renderCell(row[column] ?? null, column.endsWith("_%"))).join(" | ")} |`,
    )
    .join("\n");

  const totalsRow =
    includeTotals && totais
      ? `\n| **TOTAL** | ${colunas
          .slice(1)
          .map((column) => renderCell(totais[column] ?? null, column.endsWith("_%")))
          .join(" | ")} |`
      : "";

  return [header, separator, `${body}${totalsRow}`].join("\n");
}

/** Serializes rows to CSV. Defaults to pt-BR conventions (`;` + comma decimals). */
export function toCsv(
  columns: string[],
  rows: Array<Record<string, Cell | number | string | null>>,
  options: { delimiter?: string; decimalComma?: boolean } = {},
): string {
  const { delimiter = ";", decimalComma = true } = options;

  const escape = (input: unknown): string => {
    let text = input === null || input === undefined ? "" : String(input);
    if (decimalComma && typeof input === "number" && text.includes(".")) {
      text = text.replace(".", ",");
    }
    if (text.includes(delimiter) || text.includes('"') || /[\r\n]/.test(text)) {
      return `"${text.replace(/"/g, '""')}"`;
    }
    return text;
  };

  const lines = [columns.map(escape).join(delimiter)];
  for (const row of rows) lines.push(columns.map((column) => escape(row[column])).join(delimiter));
  // BOM keeps accented headers intact when the CSV is opened in Excel.
  return `\uFEFF${lines.join("\r\n")}\r\n`;
}

export { aggregate };


