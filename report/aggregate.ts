import type { Config } from "../src/config.js";
import { openWorkbook } from "../src/open.js";
import { stripAccents, titleCase, toNumber } from "../src/parse.js";
import { toRecords, toTable } from "../src/tables.js";
import type { Cell } from "../src/types.js";
import { isBlank } from "../src/types.js";

export interface ReportItem {
  nome: string;
  valor: number;
  pct?: number;
  quantidade?: number;
}

export interface ReportData {
  geradoEm: string;
  fonte: string;
  fonteAtualizadaEm: string | null;
  meses: string[];
  porMes: Record<string, unknown>;
}

function canonicalCategoria(raw: Cell | undefined): string {
  if (isBlank(raw)) return "Sem categoria";
  return titleCase(String(raw).trim());
}

function canonicalForma(raw: Cell | undefined): string {
  if (isBlank(raw)) return "Outro";
  const key = stripAccents(String(raw).trim().toLowerCase());
  if (key === "pix") return "Pix";
  if (key === "dinheiro" || key === "especie") return "Dinheiro";
  if (key === "cartao" || key === "credito" || key === "debito") return "Cartão";
  if (key === "boleto" || key === "ted" || key === "doc") return "Boleto";
  return titleCase(String(raw).trim()) || "Outro";
}

function monthOf(raw: Cell | undefined): string | null {
  if (typeof raw !== "string") return null;
  const match = /^(\d{4})-(\d{2})/.exec(raw.trim());
  return match ? `${match[1]}-${match[2]}` : null;
}

function pct(value: number, total: number): number {
  return total > 0 ? Math.round((value / total) * 1000) / 10 : 0;
}

function rollup(list: Array<{ nome: string; valor: number; quantidade?: number }>, total: number) {
  return list
    .filter((item) => Math.abs(item.valor) > 0.004)
    .sort((a, b) => b.valor - a.valor)
    .map((item) => ({
      ...item,
      valor: Math.round(item.valor * 100) / 100,
      pct: pct(item.valor, total),
    }));
}

/** Lê a planilha e agrega tudo em um único objeto pronto para o data.json. */
export async function buildReportData(config: Config): Promise<ReportData> {
  const source = await openWorkbook(config);

  const readTab = async (nome: string) => {
    const grid = await source.readRange(nome);
    return toRecords(toTable(grid));
  };

  const lancamentosRaw = await readTab("Lançamentos");
  const entradasRaw = await readTab("Entradas");
  const fixasRaw = await readTab("Contas fixas");

  const meses = new Set<string>();
  const lancamentos = lancamentosRaw.flatMap((r) => {
    const mes = monthOf(r["Data"]);
    const valor = toNumber(r["Valor"]);
    if (!mes || valor === null) return [];
    meses.add(mes);
    return [{
      mes,
      fornecedor: String(r["Fornecedor"] ?? "—").trim(),
      categoria: canonicalCategoria(r["Categoria"]),
      forma: canonicalForma(r["Forma de pagamento"]),
      valor,
    }];
  });

  const entradas = entradasRaw.flatMap((r) => {
    const mes = monthOf(r["Data"]);
    const valor = toNumber(r["Valor"]);
    if (!mes || valor === null) return [];
    meses.add(mes);
    return [{ mes, forma: canonicalForma(r["Forma de pagamento"]), valor }];
  });

  const fixas = fixasRaw.flatMap((r) => {
    const valor = toNumber(r["Valor"]);
    const nome = String(r["Nome"] ?? "—").trim();
    if (!nome || valor === null) return [];
    return [{ nome, valor }];
  });

  const mesesOrdenados = [...meses].sort();
  const fixasTotal = fixas.reduce((s, f) => s + f.valor, 0);

  const porMes: Record<string, unknown> = {};
  for (const mes of mesesOrdenados) {
    const lanc = lancamentos.filter((l) => l.mes === mes);
    const entr = entradas.filter((e) => e.mes === mes);

    const saidasTotal = lanc.reduce((s, l) => s + l.valor, 0);
    const entradasTotal = entr.reduce((s, e) => s + e.valor, 0);
    const despesasTotais = saidasTotal + fixasTotal;

    const categoriaMap = new Map<string, { valor: number; quantidade: number }>();
    const fornecedorMap = new Map<string, { valor: number; quantidade: number }>();
    const formaSaidaMap = new Map<string, number>();
    const formaEntradaMap = new Map<string, number>();

    for (const l of lanc) {
      const c = categoriaMap.get(l.categoria) ?? { valor: 0, quantidade: 0 };
      c.valor += l.valor;
      c.quantidade += 1;
      categoriaMap.set(l.categoria, c);

      const f = fornecedorMap.get(l.fornecedor) ?? { valor: 0, quantidade: 0 };
      f.valor += l.valor;
      f.quantidade += 1;
      fornecedorMap.set(l.fornecedor, f);

      formaSaidaMap.set(l.forma, (formaSaidaMap.get(l.forma) ?? 0) + l.valor);
    }
    for (const e of entr) formaEntradaMap.set(e.forma, (formaEntradaMap.get(e.forma) ?? 0) + e.valor);

    porMes[mes] = {
      entradas: entradasTotal,
      saidas: saidasTotal,
      fixas: fixasTotal,
      despesasTotais,
      saldo: entradasTotal - saidasTotal,
      lucro: entradasTotal - despesasTotais,
      lancamentos: lanc.length,
      categorias: rollup(
        [...categoriaMap].map(([nome, v]) => ({ nome, valor: v.valor, quantidade: v.quantidade })),
        saidasTotal,
      ),
      fornecedores: rollup(
        [...fornecedorMap].map(([nome, v]) => ({ nome, valor: v.valor, quantidade: v.quantidade })),
        saidasTotal,
      ),
      saidasPorForma: rollup(
        [...formaSaidaMap].map(([nome, valor]) => ({ nome, valor })),
        saidasTotal,
      ),
      entradasPorForma: rollup(
        [...formaEntradaMap].map(([nome, valor]) => ({ nome, valor })),
        entradasTotal,
      ),
      fixasDetalhe: rollup(fixas.map((f) => ({ nome: f.nome, valor: f.valor })), fixasTotal),
    };
  }

  const description = await source.describe();
  await source.close();

  return {
    geradoEm: new Date().toISOString(),
    fonte: description.name,
    fonteAtualizadaEm: description.lastModified ?? null,
    meses: mesesOrdenados,
    porMes,
  };
}
