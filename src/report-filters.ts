import { stripAccents, toISODate, toNumber } from "./parse.js";
import type { Cell } from "./types.js";
import { isBlank } from "./types.js";

export type FilterOp =
  | "eq"
  | "ne"
  | "contains"
  | "icontains"
  | "startswith"
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "in"
  | "not_in"
  | "between"
  | "is_empty"
  | "not_empty"
  | "regex"
  | "month";

export interface Filter {
  column: string;
  op: FilterOp;
  value?: Cell | Cell[];
}

export type Agg = "sum" | "avg" | "count" | "count_distinct" | "min" | "max" | "median" | "first" | "last";

export interface Metric {
  /** Omit for count-style aggregations. */
  column?: string;
  agg: Agg;
  /** Display name; defaults to `<agg>_<column>`. */
  as?: string;
}

export interface ReportSpec {
  groupBy?: string[];
  filters?: Filter[];
  metrics?: Metric[];
  sortBy?: string;
  sortDesc?: boolean;
  limit?: number;
  /** Include every source row reference in the output (verbose). */
  includeOrigins?: boolean;
}

export type ReportRow = Record<string, string | number | null>;

export interface ReportResult {
  colunas: string[];
  linhas: ReportRow[];
  totais: ReportRow;
  registrosTotais: number;
  registrosFiltrados: number;
  markdown: string;
  avisos: string[];
}

const num = (value: unknown): number | null => toNumber(value as Cell);

function textOf(value: Cell | undefined): string {
  if (value === null || value === undefined) return "";
  return String(value);
}

function compareValues(a: Cell | undefined, b: Cell | undefined): number {
  const na = num(a);
  const nb = num(b);
  if (na !== null && nb !== null) return na - nb;
  const da = toISODate(a);
  const db = toISODate(b);
  if (da && db) return da < db ? -1 : da > db ? 1 : 0;
  return textOf(a).localeCompare(textOf(b), "pt-BR");
}

export function matchesFilter(record: Record<string, Cell>, filter: Filter): boolean {
  const raw = record[filter.column];
  const value = filter.value;
  const list = Array.isArray(value) ? (value as Cell[]) : null;

  switch (filter.op) {
    case "is_empty":
      return isBlank(raw);
    case "not_empty":
      return !isBlank(raw);
    case "eq":
      return compareValues(raw, value as Cell) === 0;
    case "ne":
      return compareValues(raw, value as Cell) !== 0;
    case "contains":
      return textOf(raw).includes(textOf(value as Cell));
    case "icontains":
      return stripAccents(textOf(raw).toLowerCase()).includes(
        stripAccents(textOf(value as Cell).toLowerCase()),
      );
    case "startswith":
      return stripAccents(textOf(raw).toLowerCase()).startsWith(
        stripAccents(textOf(value as Cell).toLowerCase()),
      );
    case "gt":
      return compareValues(raw, value as Cell) > 0;
    case "gte":
      return compareValues(raw, value as Cell) >= 0;
    case "lt":
      return compareValues(raw, value as Cell) < 0;
    case "lte":
      return compareValues(raw, value as Cell) <= 0;
    case "in":
      return (list ?? []).some((candidate) => compareValues(raw, candidate) === 0);
    case "not_in":
      return !(list ?? []).some((candidate) => compareValues(raw, candidate) === 0);
    case "between": {
      const [from, to] = list ?? [];
      return compareValues(raw, from) >= 0 && compareValues(raw, to) <= 0;
    }
    case "regex":
      return new RegExp(textOf(value as Cell), "i").test(textOf(raw));
    case "month": {
      const iso = toISODate(raw);
      return iso !== null && iso.slice(0, 7) === textOf(value as Cell);
    }
    default:
      return true;
  }
}

export function aggregate(values: Array<Cell | undefined>, agg: Agg): number | string | null {
  const numbers = values.map((v) => num(v)).filter((v): v is number => v !== null);
  const nonEmpty = values.filter((v) => !isBlank(v));

  switch (agg) {
    case "count":
      return nonEmpty.length;
    case "count_distinct":
      return new Set(nonEmpty.map((v) => textOf(v))).size;
    case "sum":
      return numbers.reduce((total, v) => total + v, 0);
    case "avg":
      return numbers.length ? numbers.reduce((t, v) => t + v, 0) / numbers.length : null;
    case "min":
      return numbers.length ? Math.min(...numbers) : null;
    case "max":
      return numbers.length ? Math.max(...numbers) : null;
    case "median": {
      if (numbers.length === 0) return null;
      const sorted = [...numbers].sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 === 0 ? (sorted[mid - 1]! + sorted[mid]!) / 2 : sorted[mid]!;
    }
    case "first":
      return nonEmpty.length ? textOf(nonEmpty[0]) : null;
    case "last":
      return nonEmpty.length ? textOf(nonEmpty[nonEmpty.length - 1]) : null;
    default:
      return null;
  }
}

export function formatNumber(value: number | string | null, decimals = 2): string {
  if (value === null) return "";
  if (typeof value === "string") return value;
  return new Intl.NumberFormat("pt-BR", {
    minimumFractionDigits: Number.isInteger(value) ? 0 : decimals,
    maximumFractionDigits: decimals,
  }).format(value);
}
