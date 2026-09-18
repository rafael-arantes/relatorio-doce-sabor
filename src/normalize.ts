import { detectBlocks, profileColumns, type Block } from "./blocks.js";
import { gridAddress } from "./cells.js";
import { extractPaymentHint } from "./payment.js";
import { stripAccents, titleCase, toISODate, toNumber } from "./parse.js";
import { looksLikeHeaderRow } from "./tables.js";
import type { Cell, Grid } from "./types.js";
import { columnLetter, isBlank } from "./types.js";

export type Tipo = "entrada" | "saida";

export interface Transaction {
  data: string | null;
  tipo: Tipo;
  categoria: string | null;
  descricao: string | null;
  valor: number;
  conta: string | null;
  /** Provenance so every generated row can be traced back to a cell. */
  origem: {
    bloco: string | null;
    endereco: string;
    aba: string;
  };
  /** "alta" = header matched, "media" = date+value detected, "baixa" = guess. */
  confianca: "alta" | "media" | "baixa";
}

export interface NormalizeOptions {
  aba?: string;
  /** Year used for dates that only carry day + month. */
  ano?: number;
  /** Force the header row index inside each block. */
  headerRowIndex?: number;
  /** Force specific grid column indices (0-based) instead of auto-detection. */
  colunaData?: number;
  colunaValor?: number;
  colunaDescricao?: number;
  colunaConta?: number;
  colunaCategoria?: number;
  colunaTipo?: number;
  /** Ignore amounts below this (filters out stray numbers). Default 0.01. */
  valorMinimo?: number;
  /** Only normalize blocks whose label matches this regex (case/accent-insensitive). */
  apenasBlocos?: string;
  limiteLinhas?: number;
}

export interface NormalizeResult {
  transacoes: Transaction[];
  blocos: Array<{
    endereco: string;
    rotulo: string | null;
    linhas: number;
    colunaData: string | null;
    colunaValor: string | null;
    colunaDescricao: string | null;
    metodo: "cabecalho" | "heuristica" | "forcado" | "ignorado";
  }>;
  avisos: string[];
  linhasIgnoradas: number;
  truncado: boolean;
}

const HEADER_CANDIDATES = {
  data: ["data", "date", "dia", "dt", "vencimento", "data lancamento", "data pagamento", "competencia"],
  valor: ["valor", "vlr", "preco", "total", "montante", "quantia", "custo", "gasto", "saida", "rs", "r$"],
  descricao: [
    "descricao",
    "descriçao",
    "item",
    "produto",
    "fornecedor",
    "detalhe",
    "historico",
    "especificacao",
    "lancamento",
    "observacao",
    "nome",
  ],
  conta: ["conta", "pagamento", "forma", "meio", "tipo de pagamento", "forma de pagamento", "banco", "pix", "dinheiro"],
  categoria: ["categoria", "classificacao", "grupo", "setor", "centro de custo", "rubrica", "tipo"],
  tipo: ["tipo", "natureza", "operacao", "entrada saida", "debito credito"],
};

function normalizeKey(input: string): string {
  return stripAccents(input.toLowerCase()).replace(/[^a-z0-9]+/g, " ").trim();
}

/** Finds which column (0-based, inside `header`) holds a canonical field. */
function mapHeader(
  header: Cell[],
  candidates: string[],
): number | null {
  const keys = header.map((cell) => normalizeKey(isBlank(cell) ? "" : String(cell)));
  for (const candidate of candidates) {
    const wanted = normalizeKey(candidate);
    const exact = keys.findIndex((key) => key === wanted);
    if (exact >= 0) return exact;
  }
  for (const candidate of candidates) {
    const wanted = normalizeKey(candidate);
    const partial = keys.findIndex((key) => key.length > 0 && (key.includes(wanted) || wanted.includes(key)));
    if (partial >= 0) return partial;
  }
  return null;
}

const REVENUE_LABEL = /faturamento|entrada|receita|venda|recebimento|bruto/i;

/** Infers year + month from a sheet name like "junho26", "SETEMBRO" or "agosto26". */
export function inferPeriodFromSheetName(sheetName: string): { ano: number | null; mes: number | null } {
  const key = stripAccents(sheetName.toLowerCase());
  const months: Record<string, number> = {
    jan: 1, fev: 2, mar: 3, abr: 4, mai: 5, jun: 6,
    jul: 7, ago: 8, set: 9, out: 10, nov: 11, dez: 12,
  };
  let mes: number | null = null;
  for (const [prefix, number] of Object.entries(months)) {
    if (key.includes(prefix)) {
      mes = number;
      break;
    }
  }

  let ano: number | null = null;
  const fullYear = /(19|20)\d{2}/.exec(key);
  if (fullYear) {
    ano = Number(fullYear[0]);
  } else {
    // "junho26" / "set-26" style two-digit year.
    const shortYear = /(\d{2})\s*$/.exec(key);
    if (shortYear) {
      const candidate = Number(shortYear[1]);
      if (candidate >= 0 && candidate <= 99) {
        ano = candidate < 70 ? 2000 + candidate : 1900 + candidate;
      }
    }
  }
  return { ano, mes };
}

/** Builds a grid covering one block's rows, starting `skipRows` below its top. */
function subGrid(grid: Grid, block: Block, skipRows: number): Grid {
  const startRow = block.startRow + skipRows;
  if (startRow > block.startRow + block.rowCount - 1) {
    return { ...grid, rows: [], rowCount: 0, columnCount: block.columnCount, startRow: grid.startRow + startRow };
  }
  const rows = grid.rows
    .slice(startRow, block.startRow + block.rowCount)
    .map((row) => row.slice(block.startColumn, block.startColumn + block.columnCount));
  while (rows.length < 1) rows.push([]);
  return {
    ...grid,
    rows,
    rowCount: rows.length,
    columnCount: block.columnCount,
    startRow: grid.startRow + startRow,
    startColumn: grid.startColumn + block.startColumn,
  };
}

const TOTAL_LABEL = /^\s*(total|subtotal|soma|saldo|resumo)\b/i;

/** Guesses Data/Valor/Descrição columns inside a header-less block. */
function heuristicColumns(sub: Grid): {
  dataColumn: number | null;
  valorColumn: number | null;
  descricaoColumn: number | null;
} {
  const profile = profileColumns(sub, 0);
  const usable = profile.filter((p) => p.filled >= 1);

  const dateCandidates = usable
    .filter((p) => p.dateHits >= 2 && p.dateShare >= 0.5)
    .sort((a, b) => b.dateShare - a.dateShare || b.dateHits - a.dateHits);
  const dataColumn = dateCandidates[0]?.column ?? null;

  const numeric = usable
    .filter((p) => p.column !== dataColumn && p.numberHits >= 1 && p.numberShare >= 0.6)
    .sort((a, b) => a.column - b.column);
  let valorColumn: number | null = null;
  if (dataColumn !== null) {
    const adjacent = numeric.find((p) => p.column > dataColumn);
    valorColumn = adjacent ? adjacent.column : (numeric[numeric.length - 1]?.column ?? null);
  } else if (numeric.length > 0) {
    // No date anywhere: the last numeric column is the most likely amount.
    valorColumn = numeric[numeric.length - 1]!.column;
  }

  const textCandidates = usable
    .filter((p) => p.column !== dataColumn && p.column !== valorColumn && p.textHits >= 1)
    .sort((a, b) => b.textHits - a.textHits || b.distinct - a.distinct || a.column - b.column);
  const descricaoColumn = textCandidates[0]?.column ?? null;

  return { dataColumn, valorColumn, descricaoColumn };
}

/**
 * Turns any sheet into Aba 1-shaped transactions — the schema the report layer
 * (and the client's new spreadsheet) expects:
 * `Data | Tipo | Categoria | Descrição | Valor | Conta`.
 *
 * Three strategies per block, in order of confidence:
 * 1. `cabecalho`  — the block has a real header row (Data/Valor/Descrição/...)
 * 2. `heuristica` — no header, so column roles are inferred from cell types
 * 3. `forcado`    — the caller pinned the columns explicitly
 *
 * Payment methods hidden in item labels ("ovo caipira (dinheiro)") are pulled
 * out into `conta` and stripped from `descricao`.
 */
export function normalizeTransactions(
  grid: Grid,
  options: NormalizeOptions = {},
): NormalizeResult {
  const { valorMinimo = 0.01, limiteLinhas = 5000 } = options;
  const aba = options.aba ?? grid.sheet;
  const period = inferPeriodFromSheetName(aba);
  const defaultYear = options.ano ?? period.ano ?? new Date().getUTCFullYear();

  const blockFilter = options.apenasBlocos
    ? new RegExp(stripAccents(options.apenasBlocos).toLowerCase(), "i")
    : null;
  const forced =
    options.colunaData !== undefined ||
    options.colunaValor !== undefined ||
    options.colunaDescricao !== undefined;

  const { blocks, truncated: blocksTruncated } = detectBlocks(grid, { minCells: 2 });
  const transacoes: Transaction[] = [];
  const blocos: NormalizeResult["blocos"] = [];
  const avisos: string[] = [];
  let linhasIgnoradas = 0;
  let truncado = blocksTruncated;

  for (const block of blocks) {
    const label = block.label;
    if (blockFilter && !blockFilter.test(stripAccents(label ?? block.address).toLowerCase())) {
      blocos.push({
        endereco: block.address,
        rotulo: label,
        linhas: block.rowCount,
        colunaData: null,
        colunaValor: null,
        colunaDescricao: null,
        metodo: "ignorado",
      });
      continue;
    }

    const headerRowIndex = options.headerRowIndex ?? block.headerRowOffset;
    const localHeader = (grid.rows[block.startRow + headerRowIndex] ?? []).slice(
      block.startColumn,
      block.startColumn + block.columnCount,
    );

    const fromHeader = {
      data: mapHeader(localHeader, HEADER_CANDIDATES.data),
      valor: mapHeader(localHeader, HEADER_CANDIDATES.valor),
      descricao: mapHeader(localHeader, HEADER_CANDIDATES.descricao),
      conta: mapHeader(localHeader, HEADER_CANDIDATES.conta),
      categoria: mapHeader(localHeader, HEADER_CANDIDATES.categoria),
      tipo: mapHeader(localHeader, HEADER_CANDIDATES.tipo),
    };

    let metodo: "cabecalho" | "heuristica" | "forcado" = "heuristica";
    let dataColumn: number | null = null;
    let valorColumn: number | null = null;
    let descricaoColumn: number | null = null;
    let contaColumn: number | null = null;
    let categoriaColumn: number | null = null;
    let tipoColumn: number | null = null;
    let dataStartOffset = headerRowIndex;

    if (forced) {
      metodo = "forcado";
      dataColumn = options.colunaData ?? null;
      valorColumn = options.colunaValor ?? null;
      descricaoColumn = options.colunaDescricao ?? null;
      contaColumn = options.colunaConta ?? null;
      categoriaColumn = options.colunaCategoria ?? null;
      tipoColumn = options.colunaTipo ?? null;
    } else if (fromHeader.valor !== null) {
      metodo = "cabecalho";
      dataColumn = fromHeader.data;
      valorColumn = fromHeader.valor;
      descricaoColumn = fromHeader.descricao;
      contaColumn = fromHeader.conta;
      categoriaColumn = fromHeader.categoria;
      tipoColumn = fromHeader.tipo;
      dataStartOffset = headerRowIndex + 1;
    } else {
      const guess = heuristicColumns(subGrid(grid, block, headerRowIndex));
      dataColumn = guess.dataColumn;
      valorColumn = guess.valorColumn;
      descricaoColumn = guess.descricaoColumn;
      // The block has no recognizable field names, but a header row may still be
      // sitting there (e.g. "Item | Valor") — don't emit it as a transaction.
      if (looksLikeHeaderRow(localHeader)) dataStartOffset = headerRowIndex + 1;
      if (valorColumn === null) {
        avisos.push(
          `Bloco ${block.address}${label ? ` ("${label}")` : ""}: nenhuma coluna de valor reconhecida — ignorado.`,
        );
        blocos.push({
          endereco: block.address,
          rotulo: label,
          linhas: block.rowCount,
          colunaData: null,
          colunaValor: null,
          colunaDescricao: null,
          metodo: "ignorado",
        });
        continue;
      }
    }

    if (valorColumn === null) {
      avisos.push(`Bloco ${block.address}: coluna de valor não informada — ignorado.`);
      continue;
    }

    const isRevenue = REVENUE_LABEL.test(stripAccents(label ?? ""));
    const categoriaFromLabel = label && !isRevenue && label.length <= 30 ? titleCase(label) : null;

    let blockRows = 0;
    const lastRow = block.startRow + block.rowCount - 1;

    for (let r = block.startRow + dataStartOffset; r <= lastRow; r++) {
      const row = grid.rows[r] ?? [];
      const at = (offset: number | null): Cell =>
        offset === null ? null : (row[block.startColumn + offset] ?? null);

      const valorRaw = at(valorColumn);
      const valor = toNumber(valorRaw);
      const descRaw = at(descricaoColumn);
      const descricaoText = typeof descRaw === "string" ? descRaw : null;

      if (valor === null || Math.abs(valor) < valorMinimo) {
        if (!isBlank(valorRaw) || !isBlank(descRaw)) linhasIgnoradas++;
        continue;
      }
      if (descricaoText && TOTAL_LABEL.test(descricaoText)) {
        linhasIgnoradas++;
        continue;
      }

      const hint = extractPaymentHint(descricaoText);
      const contaCell = at(contaColumn);
      const catCell = at(categoriaColumn);
      const tipoCell = at(tipoColumn);
      const dataCell = at(dataColumn);

      const conta = isBlank(contaCell) ? (hint?.conta ?? null) : titleCase(String(contaCell).trim());

      const tipoText = isBlank(tipoCell) ? "" : stripAccents(String(tipoCell).toLowerCase());
      let tipo: Tipo;
      if (tipoText.startsWith("entrada") || tipoText.includes("receita") || tipoText.includes("credito")) {
        tipo = "entrada";
      } else if (tipoText.startsWith("saida") || tipoText.includes("despesa") || tipoText.includes("debito")) {
        tipo = "saida";
      } else {
        tipo = isRevenue ? "entrada" : "saida";
      }

      const descricaoFinal = hint?.cleaned ?? descricaoText ?? label ?? null;

      transacoes.push({
        data: toISODate(dataCell, { defaultYear }),
        tipo,
        categoria: isBlank(catCell) ? categoriaFromLabel : titleCase(String(catCell).trim()),
        descricao:
          descricaoFinal && descricaoFinal.trim() !== "" ? descricaoFinal.trim() : null,
        valor: Math.abs(valor),
        conta,
        origem: {
          bloco: label,
          endereco: gridAddress(grid, r, block.startColumn + valorColumn),
          aba,
        },
        confianca:
          metodo === "cabecalho" || metodo === "forcado"
            ? "alta"
            : dataColumn !== null
              ? "media"
              : "baixa",
      });
      blockRows++;

      if (transacoes.length >= limiteLinhas) {
        truncado = true;
        break;
      }
    }

    blocos.push({
      endereco: block.address,
      rotulo: label,
      linhas: blockRows,
      colunaData: dataColumn === null ? null : columnLetter(block.startColumn + dataColumn),
      colunaValor: columnLetter(block.startColumn + valorColumn),
      colunaDescricao:
        descricaoColumn === null ? null : columnLetter(block.startColumn + descricaoColumn),
      metodo,
    });

    if (metodo === "heuristica" && dataColumn === null) {
      avisos.push(
        `Bloco ${block.address}${label ? ` ("${label}")` : ""}: sem coluna de data detectada — ` +
          "as datas ficaram nulas; informe `ano`/`colunaData` se precisar delas.",
      );
    }
    if (truncado) break;
  }

  return { transacoes, blocos, avisos, linhasIgnoradas, truncado };
}



