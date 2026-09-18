#!/usr/bin/env node
/**
 * Gera o data.json do relatório financeiro a partir da planilha Excel Online.
 *
 * Uso:  npm run report        (lê a planilha via Graph e regrava report/data.json)
 *
 * O relatório em si (report/index.html) é 100% estático: ele só carrega o
 * data.json. Para atualizar os números, rode este script e faça deploy dos
 * arquivos de novo (ou só do data.json).
 *
 * Para atualização automática/sob demanda, veja `server.ts` e `npm run report:serve`.
 */
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { loadConfig } from "../src/config.js";
import { buildReportData } from "./aggregate.js";

const OUT = resolve(import.meta.dirname, "data.json");

async function main(): Promise<void> {
  const payload = await buildReportData(loadConfig());
  mkdirSync(dirname(OUT), { recursive: true });
  writeFileSync(OUT, JSON.stringify(payload, null, 2), "utf8");

  console.log(`✅ ${OUT}`);
  console.log(`   meses: ${payload.meses.join(", ") || "(nenhum)"}`);
  console.log(`   gerado em: ${payload.geradoEm}`);
  if (payload.meses.length === 0) console.log("⚠️  Nenhum mês encontrado — verifique as datas.");
}

main().catch((error: unknown) => {
  console.error("❌", error instanceof Error ? error.message : error);
  process.exit(1);
});
