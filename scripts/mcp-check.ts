#!/usr/bin/env node
/**
 * Boots the compiled MCP server over stdio, speaks the real protocol, and runs a
 * full read → report round trip against a local fixture. This is the test that
 * proves the wiring works for a real client (Claude Code / Claude Desktop).
 *
 * Usage: npm run build && npm run check
 */
import { rmSync } from "node:fs";
import { resolve } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { buildFixture } from "./fixture.js";

let failures = 0;
function check(label: string, condition: boolean, detail?: unknown): void {
  if (condition) {
    console.log(`  ✅ ${label}`);
  } else {
    failures++;
    console.log(`  ❌ ${label}`);
    if (detail !== undefined) console.log(`     obtido: ${JSON.stringify(detail).slice(0, 600)}`);
  }
}

function textOf(result: unknown): string {
  const content = (result as { content?: Array<{ type: string; text?: string }> }).content ?? [];
  return content.map((part) => part.text ?? "").join("\n");
}

function jsonOf(result: unknown): unknown {
  const raw = textOf(result);
  const fence = /```json\n([\s\S]*?)\n```/.exec(raw);
  const payload = fence ? fence[1]! : raw;
  try {
    return JSON.parse(payload);
  } catch {
    return null;
  }
}

async function main(): Promise<void> {
  const projectRoot = resolve(import.meta.dirname, "..");
  const tmpDir = resolve(projectRoot, ".tokens", "tmp");
  const fixture = await buildFixture(resolve(tmpDir, "PLANILHA-mcp-check.xlsx"));
  const serverEntry = resolve(projectRoot, "dist/src/index.js");

  console.log("excel-mcp protocol check\n═══════════════════════════════════════════════════════");
  console.log(`server  : ${serverEntry}`);
  console.log(`fixture : ${fixture}\n`);

  const client = new Client({ name: "excel-mcp-check", version: "0.1.0" });
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [serverEntry],
    env: {
      ...process.env,
      EXCEL_SOURCE: "local",
      EXCEL_LOCAL_FILE: fixture,
      EXCEL_DEBUG: "0",
    } as Record<string, string>,
    stderr: "pipe",
  });

  await client.connect(transport);

  console.log("1) Handshake");
  check("servidor conectado", client.getServerVersion()?.name === "excel-mcp", client.getServerVersion());

  console.log("\n2) tools/list");
  const { tools } = await client.listTools();
  const names = tools.map((tool) => tool.name).sort();
  console.log(`   ${names.join(", ")}`);
  const expected = [
    "excel_build_report",
    "excel_column_profile",
    "excel_detect_blocks",
    "excel_diagnostics",
    "excel_export_csv",
    "excel_list_metrics",
    "excel_normalize_transactions",
    "excel_read_range",
    "excel_read_sheet",
    "excel_search",
    "excel_workbook_info",
  ];
  check(`${expected.length} ferramentas registradas`, names.length === expected.length, names);
  check("nomes esperados", expected.every((name) => names.includes(name)), names);
  check("todas têm descrição", tools.every((tool) => Boolean(tool.description && tool.description.length > 20)));
  check(
    "read-only exceto export",
    tools.every((tool) => tool.name === "excel_export_csv" || tool.annotations?.readOnlyHint === true),
  );

  console.log("\n3) excel_workbook_info");
  const info = jsonOf(await client.callTool({ name: "excel_workbook_info", arguments: {} })) as {
    name?: string;
    sheets?: Array<{ name: string }>;
  };
  check("nome do arquivo", info?.name === "PLANILHA-mcp-check.xlsx", info?.name);
  check("abas listadas", info?.sheets?.length === 2, info?.sheets);

  console.log("\n4) excel_detect_blocks");
  const blocks = jsonOf(
    await client.callTool({ name: "excel_detect_blocks", arguments: { aba: "junho26" } }),
  ) as { totalBlocos?: number; blocos?: Array<{ rotulo: string }> };
  check("4 blocos", blocks?.totalBlocos === 4, blocks?.totalBlocos);
  check(
    "rótulos",
    blocks?.blocos?.map((b) => b.rotulo).join(",") === "MERCADO,BEBIDAS,QUITANDA,FATURAMENTO",
    blocks?.blocos?.map((b) => b.rotulo),
  );

  console.log("\n5) excel_build_report (normalized mode)");
  const report = jsonOf(
    await client.callTool({
      name: "excel_build_report",
      arguments: {
        aba: "junho26",
        agruparPor: ["categoria"],
        metricas: [{ agg: "sum", column: "valor", as: "total" }],
        ordenarPor: "total",
        ordemDesc: true,
      },
    }),
  ) as { totais?: { total?: number }; linhas?: unknown[] };
  check("4 grupos", report?.linhas?.length === 4, report?.linhas?.length);
  check("total 13314.52", Math.abs(Number(report?.totais?.total) - 13314.52) < 0.001, report?.totais);

  console.log("\n6) excel_normalize_transactions");
  const normalized = jsonOf(
    await client.callTool({ name: "excel_normalize_transactions", arguments: { aba: "junho26" } }),
  ) as { resumo?: { quantidade?: number; porConta?: Record<string, number> } };
  check("8 transações", normalized?.resumo?.quantidade === 8, normalized?.resumo?.quantidade);
  check(
    "contas extraídas dos rótulos",
    normalized?.resumo?.porConta?.PIX === 320.5 && normalized?.resumo?.porConta?.Dinheiro === 88.9,
    normalized?.resumo?.porConta,
  );

  console.log("\n7) excel_list_metrics / excel_search / excel_column_profile");
  const metrics = jsonOf(
    await client.callTool({ name: "excel_list_metrics", arguments: { aba: "Lançamentos" } }),
  ) as { modoBruto?: { colunas?: string[] } };
  check(
    "colunas brutas expostas",
    metrics?.modoBruto?.colunas?.includes("Descrição") === true,
    metrics?.modoBruto?.colunas,
  );

  const search = jsonOf(
    await client.callTool({ name: "excel_search", arguments: { aba: "junho26", termo: "ambev" } }),
  ) as { ocorrencias?: number };
  check("busca encontra 'ambev'", search?.ocorrencias === 1, search?.ocorrencias);

  const profile = jsonOf(
    await client.callTool({ name: "excel_column_profile", arguments: { aba: "junho26" } }),
  ) as { colunas?: Array<{ letra: string }> };
  check(
    "perfil alcança a coluna K",
    profile?.colunas?.[(profile.colunas?.length ?? 0) - 1]?.letra === "K",
    profile?.colunas?.map((c) => c.letra),
  );

  console.log("\n8) excel_read_sheet + excel_export_csv");
  const sheet = jsonOf(
    await client.callTool({ name: "excel_read_sheet", arguments: { aba: "Lançamentos", formato: "records" } }),
  ) as { colunas?: string[]; linhasTotais?: number };
  check("linhas lidas na aba limpa", sheet?.linhasTotais === 4, sheet?.linhasTotais);

  const exported = jsonOf(
    await client.callTool({
      name: "excel_export_csv",
      arguments: { aba: "junho26", caminho: resolve(tmpDir, "export.csv") },
    }),
  ) as { linhas?: number; arquivo?: string };
  check("8 linhas exportadas", exported?.linhas === 8, exported?.linhas);

  console.log("\n9) Erro amigável em argumento inválido");
  const bad = await client.callTool({
    name: "excel_read_sheet",
    arguments: { aba: "aba-que-nao-existe" },
  });
  const badText = textOf(bad);
  check("erro devolvido como resultado, sem crash", (bad as { isError?: boolean }).isError === true, badText.slice(0, 200));
  check("mensagem cita as abas disponíveis", badText.includes("junho26"), badText.slice(0, 300));

  await client.close();
  rmSync(tmpDir, { recursive: true, force: true });
  console.log("\n═══════════════════════════════════════════════════════");
  if (failures === 0) {
    console.log("✅ protocolo MCP validado ponta a ponta");
  } else {
    console.log(`❌ ${failures} verificação(ões) falharam`);
  }
  process.exit(failures === 0 ? 0 : 1);
}

main().catch((error: unknown) => {
  console.error("\n❌ protocol check crashed:", error instanceof Error ? error.stack : error);
  process.exit(1);
});
