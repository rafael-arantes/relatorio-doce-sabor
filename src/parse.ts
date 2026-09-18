import type { Cell } from "./types.js";

const MONTHS_PT: Record<string, number> = {
  jan: 1, janeiro: 1,
  fev: 2, fevereiro: 2,
  mar: 3, marco: 3, março: 3,
  abr: 4, abril: 4,
  mai: 5, maio: 5,
  jun: 6, junho: 6,
  jul: 7, julho: 7,
  ago: 8, agosto: 8,
  set: 9, setembro: 9,
  out: 10, outubro: 10,
  nov: 11, novembro: 11,
  dez: 12, dezembro: 12,
};

export function stripAccents(input: string): string {
  return input.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Parses numbers written in Brazilian ("R$ 1.234,56"), US ("1,234.56") or
 * plain ("1234.5") form. Also understands parentheses for negatives and a
 * trailing percent sign.
 */
export function toNumber(raw: Cell | undefined): number | null {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === "number") return Number.isFinite(raw) ? raw : null;
  if (typeof raw === "boolean") return raw ? 1 : 0;

  let text = raw.trim();
  if (!text) return null;
  text = text.replace(/^(r\$|rs\.?|brl|\$)\s*/i, "").trim();

  let negative = false;
  if (/^\(.*\)$/.test(text)) {
    negative = true;
    text = text.slice(1, -1).trim();
  }
  const percent = text.endsWith("%");
  if (percent) text = text.slice(0, -1).trim();

  text = text.replace(/\s/g, "");
  const digitsOnly = text.replace(/[^0-9.,-]/g, "");
  if (!digitsOnly) return null;

  const lastComma = digitsOnly.lastIndexOf(",");
  const lastDot = digitsOnly.lastIndexOf(".");
  const decimals = Math.max(
    lastComma >= 0 ? digitsOnly.length - lastComma - 1 : -1,
    lastDot >= 0 ? digitsOnly.length - lastDot - 1 : -1,
  );

  let normalized = digitsOnly;
  if (lastComma >= 0 && lastDot >= 0) {
    // Both present: whichever comes last is the decimal separator.
    normalized =
      lastComma > lastDot
        ? digitsOnly.replace(/\./g, "").replace(",", ".")
        : digitsOnly.replace(/,/g, "");
  } else if (lastComma >= 0) {
    // Only commas: "1,5" is decimal, a 3-digit group is thousands (pt-BR).
    normalized = decimals === 3 ? digitsOnly.replace(/,/g, "") : digitsOnly.replace(",", ".");
  } else if (decimals === 3 && /^-?\d{1,3}(\.\d{3})+$/.test(digitsOnly)) {
    // "1.234" / "12.345.678" => thousands separators.
    normalized = digitsOnly.replace(/\./g, "");
  }

  const value = Number.parseFloat(normalized);
  if (!Number.isFinite(value)) return null;
  const signed = negative ? -Math.abs(value) : value;
  return percent ? signed / 100 : signed;
}

const MS_PER_DAY = 86_400_000;
/** Excel's day 0, so serial 1 == 1900-01-01. */
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

export function excelSerialToDate(serial: number): Date {
  return new Date(EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY);
}

function iso(date: Date): string {
  return date.toISOString().slice(0, 10);
}

export interface DateParseOptions {
  /** Used when the cell only carries day + month. */
  defaultYear?: number;
  /** Accept bare numbers as Excel serial dates (default true). */
  serialNumbers?: boolean;
}

/**
 * Parses the date shapes that show up in hand-made spreadsheets:
 * 12/06/2026, 12/06, 12-06-26, 2026-06-12, "12 de junho", "12/jun", Excel
 * serial numbers and real Date objects. Returns YYYY-MM-DD or null.
 */
export function toISODate(
  raw: Cell | Date | undefined,
  options: DateParseOptions = {},
): string | null {
  const { defaultYear, serialNumbers = true } = options;
  if (raw === null || raw === undefined) return null;

  if (raw instanceof Date) {
    return Number.isNaN(raw.getTime()) ? null : iso(raw);
  }

  if (typeof raw === "number") {
    if (!serialNumbers) return null;
    // Plausible Excel serial window: 1950-01-01 .. 2080-01-01.
    if (raw >= 18_263 && raw <= 65_780) return iso(excelSerialToDate(raw));
    return null;
  }

  if (typeof raw !== "string") return null;

  const text = raw.trim();
  if (!text) return null;

  const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;

  const numeric = /^(\d{1,2})[/\-.](\d{1,2})(?:[/\-.](\d{2,4}))?$/.exec(text);
  if (numeric) {
    const day = Number(numeric[1]);
    const month = Number(numeric[2]);
    if (day < 1 || day > 31 || month < 1 || month > 12) return null;
    let year = numeric[3] ? Number(numeric[3]) : defaultYear;
    if (!year) return null;
    if (year < 100) year += year < 70 ? 2000 : 1900;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  // "12 de junho" | "12 junho 2026" | "12/jun" | "12 jun"
  const named =
    /^(\d{1,2})\s*(?:de|do|[/\-. ])\s*([a-zçãé]+)\.?\s*(?:de|do|[/\-. ])?\s*(\d{2,4})?$/.exec(
      stripAccents(text.toLowerCase()),
    );
  if (named) {
    const month = MONTHS_PT[named[2] ?? ""];
    if (!month) return null;
    const day = Number(named[1]);
    if (day < 1 || day > 31) return null;
    let year = named[3] ? Number(named[3]) : defaultYear;
    if (!year) return null;
    if (year < 100) year += year < 70 ? 2000 : 1900;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  return null;
}

/** True when the value is *only* a date, nothing else appended. */
export function looksLikeDate(raw: Cell | undefined, options: DateParseOptions = {}): boolean {
  if (raw === null || raw === undefined) return false;
  if (typeof raw === "number") return toISODate(raw, options) !== null;
  if (typeof raw !== "string") return false;
  return toISODate(raw, options) !== null;
}

export function titleCase(input: string): string {
  return input
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => (word.length <= 2 ? word : word[0]!.toUpperCase() + word.slice(1)))
    .join(" ");
}

