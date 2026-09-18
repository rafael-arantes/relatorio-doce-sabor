#!/usr/bin/env node
/**
 * Offline smoke test: builds a fixture that mimics the client's messy monthly
 * sheet (per-supplier blocks, Data/Valor pairs, payment method buried in the
 * item label), then runs the whole pipeline — blocks, normalize, report, CSV —
 * and asserts the results. No network, no Azure, no credentials needed.
 */
import { readFileSync, rmSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { TokenProvider, AuthError } from "../src/auth.js";
import { signInWithBrowser } from "../src/auth-code.js";
import { loadConfig } from "../src/config.js";
import { detectBlocks } from "../src/blocks.js";
import { normalizeTransactions } from "../src/normalize.js";
import { buildReport, toCsv } from "../src/report.js";
import { encodeDrivePath, encodeSharingUrl, GraphError } from "../src/sources/graph-api.js";
import { GraphWorkbookSource, isWorkbookApiUnavailable } from "../src/sources/graph.js";
import { LocalWorkbookSource, parseA1 } from "../src/sources/local.js";
import { toRecords, toTable } from "../src/tables.js";
import { transactionRecords } from "../src/tools/report.js";
import { buildFixture } from "./fixture.js";

let failures = 0;
function check(label: string, condition: boolean, detail?: unknown): void {
  if (condition) {
    console.log(`  ✅ ${label}`);
  } else {
    failures++;
    console.log(`  ❌ ${label}`);
    if (detail !== undefined) console.log(`     obtido: ${JSON.stringify(detail)}`);
  }
}

const config = loadConfig();
const tmpDir = resolve(config.projectRoot, ".tokens", "tmp");
const fixturePath = resolve(tmpDir, "PLANILHA-fixture.xlsx");

async function main(): Promise<void> {
  console.log("excel-mcp smoke test (offline)\n═══════════════════════════════════════════════════════");
  await buildFixture(fixturePath);
  console.log(`Fixture: ${fixturePath}\n`);

  const source = await LocalWorkbookSource.open(config, fixturePath);

  console.log("1) Workbook description");
  const description = await source.describe();
  check("duas abas encontradas", description.sheets.length === 2, description.sheets.map((s) => s.name));
  check(
    "nomes das abas",
    description.sheets.map((s) => s.name).join(",") === "junho26,Lançamentos",
    description.sheets.map((s) => s.name),
  );

  console.log("\n2) Block detection on the messy tab");
  const juneGrid = await source.readRange("junho26");
  const { blocks } = detectBlocks(juneGrid);
  check(
    "4 blocos detectados",
    blocks.length === 4,
    blocks.map((b) => `${b.address}:${b.label}`),
  );
  check(
    "rótulos dos blocos",
    blocks.map((b) => b.label).join(",") === "MERCADO,BEBIDAS,QUITANDA,FATURAMENTO",
    blocks.map((b) => b.label),
  );
  check("bloco MERCADO abrange A1:B4", blocks[0]?.address === "A1:B4", blocks[0]?.address);

  console.log("\n3) Normalization into the Aba 1 schema");
  const normalized = normalizeTransactions(juneGrid, { aba: "junho26" });
  check("8 transações", normalized.transacoes.length === 8, normalized.transacoes.length);
  check("sem avisos", normalized.avisos.length === 0, normalized.avisos);

  const byBlock = (name: string) => normalized.transacoes.filter((t) => t.origem.bloco === name);

  const mercado = byBlock("MERCADO");
  check("MERCADO: valor numérico", mercado[0]?.valor === 1234.56, mercado[0]?.valor);
  check("MERCADO: data normalizada", mercado[0]?.data === "2026-06-02", mercado[0]?.data);
  check("MERCADO: categoria vinda do rótulo", mercado[0]?.categoria === "Mercado", mercado[0]?.categoria);
  check("MERCADO: tipo=saida", mercado[0]?.tipo === "saida", mercado[0]?.tipo);
  check("MERCADO: confiança alta", mercado[0]?.confianca === "alta", mercado[0]?.confianca);

  const bebidas = byBlock("BEBIDAS");
  check("BEBIDAS: conta extraída do rótulo (pix)", bebidas[0]?.conta === "PIX", bebidas[0]?.conta);
  check("BEBIDAS: descrição limpa", bebidas[0]?.descricao === "ambev", bebidas[0]?.descricao);
  check("BEBIDAS: conta dinheiro", bebidas[1]?.conta === "Dinheiro", bebidas[1]?.conta);
  check("BEBIDAS: categoria do rótulo", bebidas[0]?.categoria === "Bebidas", bebidas[0]?.categoria);

  const quitanda = byBlock("QUITANDA");
  check("QUITANDA: 'R$ 1.234,56' → 1234.56", quitanda[0]?.valor === 1234.56, quitanda[0]?.valor);
  check("QUITANDA: '45,90' → 45.9", quitanda[1]?.valor === 45.9, quitanda[1]?.valor);

  const faturamento = byBlock("FATURAMENTO");
  check(
    "FATURAMENTO: tipo=entrada",
    faturamento.every((t) => t.tipo === "entrada"),
    faturamento.map((t) => t.tipo),
  );
  check(
    "FATURAMENTO: ano inferido do nome da aba (2026)",
    faturamento[0]?.data === "2026-06-02",
    faturamento[0]?.data,
  );
  check("FATURAMENTO: sem categoria", faturamento[0]?.categoria === null, faturamento[0]?.categoria);

  console.log("\n4) Report over normalized rows");
  const records = transactionRecords(juneGrid, { aba: "junho26" });
  const byCategoria = buildReport(records, {
    groupBy: ["categoria"],
    metrics: [{ agg: "sum", column: "valor", as: "total" }],
    sortBy: "total",
    sortDesc: true,
  });
  check(
    "4 grupos de categoria",
    byCategoria.linhas.length === 4,
    byCategoria.linhas.map((r) => r.categoria),
  );
  check(
    "total geral = 13314.52",
    Math.abs(Number(byCategoria.totais.total) - 13314.52) < 0.001,
    byCategoria.totais.total,
  );
  check("markdown gerado", byCategoria.markdown.includes("| categoria | total |"), byCategoria.markdown);

  const byMonth = buildReport(records, {
    groupBy: ["mes"],
    filters: [{ column: "mes", op: "eq", value: "2026-06" }],
    metrics: [{ agg: "count", as: "lancamentos" }],
  });
  check("filtro por mês retorna 1 grupo", byMonth.linhas.length === 1, byMonth.linhas);
  // Only MERCADO and FATURAMENTO carry a date column in the fixture, so the
  // other two blocks legitimately produce rows with `data: null`.
  check("filtro por mês conta 4 lançamentos datados", byMonth.linhas[0]?.lancamentos === 4, byMonth.linhas[0]);
  check(
    "filtro is_empty encontra os 4 sem data",
    buildReport(records, {
      filters: [{ column: "data", op: "is_empty" }],
      metrics: [{ agg: "count", as: "sem_data" }],
    }).totais.sem_data === 4,
  );

  const onlyAmbev = buildReport(records, {
    filters: [{ column: "descricao", op: "icontains", value: "AMBEV" }],
    metrics: [{ agg: "sum", column: "valor", as: "total" }],
  });
  check("busca textual ignora caixa", onlyAmbev.totais.total === 320.5, onlyAmbev.totais.total);

  console.log("\n5) Raw-column mode (clean tab)");
  const targetGrid = await source.readRange("Lançamentos");
  const table = toTable(targetGrid);
  check(
    "cabeçalhos da aba limpa",
    table.headers.join(",") === "Data,Tipo,Categoria,Descrição,Valor,Conta,Status",
    table.headers,
  );
  const rawReport = buildReport(toRecords(table), {
    groupBy: ["Tipo"],
    filters: [{ column: "Status", op: "eq", value: "pago" }],
    metrics: [{ agg: "sum", column: "Valor", as: "total" }],
  });
  check("só linhas 'pago' consideradas", rawReport.registrosFiltrados === 3, rawReport.registrosFiltrados);

  console.log("\n6) CSV export");
  const csv = toCsv(
    ["data", "descricao", "valor"],
    normalized.transacoes.map((t) => ({ data: t.data, descricao: t.descricao, valor: t.valor })),
  );
  check("CSV com decimal vírgula", csv.includes("1234,56"), csv.split("\r\n")[1]);
  check("CSV com BOM", csv.startsWith("\uFEFF"));
  check("CSV com cabeçalho ;", csv.includes("data;descricao;valor"));

  await source.close();

  console.log("\n7) Graph request encoding");
  // Reference token produced by an independent implementation (Python
  // base64.urlsafe_b64encode for the sample URL in Microsoft's docs) — this
  // catches a subtly wrong base64url step, which silently 404s against /shares.
  const sampleUrl = "https://onedrive.live.com/redir?resid=1231244193912!12&authKey=1201919!12921!1";
  check(
    "encodeSharingUrl bate com a implementação de referência",
    encodeSharingUrl(sampleUrl) ===
      "u!aHR0cHM6Ly9vbmVkcml2ZS5saXZlLmNvbS9yZWRpcj9yZXNpZD0xMjMxMjQ0MTkzOTEyITEyJmF1dGhLZXk9MTIwMTkxOSExMjkyMSEx",
    encodeSharingUrl(sampleUrl),
  );
  check("sem padding '=' e sem '+' nem '/'", !/[=+/]/.test(encodeSharingUrl(sampleUrl).slice(2)));
  check(
    "encodeDrivePath preserva as barras e codifica cada segmento",
    encodeDrivePath("/Documentos/Planilha Doce Sabor.xlsx") ===
      "/Documentos/Planilha%20Doce%20Sabor.xlsx",
    encodeDrivePath("/Documentos/Planilha Doce Sabor.xlsx"),
  );
  check(
    "encodeDrivePath codifica acentos",
    encodeDrivePath("/Documentos/PLANILHAÇÃO.xlsx") === "/Documentos/PLANILHA%C3%87%C3%83O.xlsx",
    encodeDrivePath("/Documentos/PLANILHAÇÃO.xlsx"),
  );

  console.log("\n8) Classificação de erro para o fallback de download");
  check(
    "403 → indisponível (usa fallback)",
    isWorkbookApiUnavailable(new GraphError("nope", 403, "accessDenied")),
  );
  check(
    "invalidRequest → indisponível",
    isWorkbookApiUnavailable(new GraphError("nope", 400, "invalidRequest")),
  );
  check("500 → indisponível", isWorkbookApiUnavailable(new GraphError("boom", 503, "serviceUnavailable")));
  check(
    "404 sheet errado → NÃO usa fallback",
    !isWorkbookApiUnavailable(new GraphError("Worksheet not found", 404, "itemNotFound")),
  );
  check("erro não-Graph → NÃO usa fallback", !isWorkbookApiUnavailable(new Error("rede caiu")));

  console.log("\n9) Leitura a partir de bytes (mesmo caminho do graph_download)");
  const fromBytes = await LocalWorkbookSource.fromBuffer(config, readFileSync(fixturePath), {
    name: "PLANILHA-bytes.xlsx",
    source: "graph",
    location: "drive=test item=test",
  });
  const bytesDescription = await fromBytes.describe();
  check("descreve como graph", bytesDescription.source === "graph", bytesDescription.source);
  check("nome do meta preservado", bytesDescription.name === "PLANILHA-bytes.xlsx", bytesDescription.name);
  check("abas lidas dos bytes", bytesDescription.sheets.length === 2, bytesDescription.sheets.length);
  const bytesGrid = await fromBytes.readRange("junho26");
  check("mesma grade do arquivo em disco", bytesGrid.rowCount === juneGrid.rowCount, bytesGrid.rowCount);
  check(
    "mesmo bloco detectado a partir dos bytes",
    detectBlocks(bytesGrid).blocks.length === 4,
    detectBlocks(bytesGrid).blocks.length,
  );
  const latin = await LocalWorkbookSource.fromBuffer(config, Buffer.from("not-an-xlsx"), {
    name: "broken.xls",
    source: "local",
    location: "/tmp/broken.xls",
  }).catch((error: Error) => error.message);
  check(
    "arquivo que não é .xlsx dá erro claro",
    typeof latin === "string" && latin.includes("re-saved as .xlsx"),
    latin,
  );

  console.log("\n10) parseA1");
  check(
    "intervalo normalizado",
    JSON.stringify(parseA1("B2:D200")) ===
      JSON.stringify({ start: { row: 2, column: 2 }, end: { row: 200, column: 4 } }),
    parseA1("B2:D200"),
  );
  check(
    "intervalo invertido é corrigido",
    JSON.stringify(parseA1("D200:B2")) ===
      JSON.stringify({ start: { row: 2, column: 2 }, end: { row: 200, column: 4 } }),
    parseA1("D200:B2"),
  );
  check(
    "aceita 'Aba!A1:B2' e cifrão",
    JSON.stringify(parseA1("junho26!$A$1:$B$2")) ===
      JSON.stringify({ start: { row: 1, column: 1 }, end: { row: 2, column: 2 } }),
    parseA1("junho26!$A$1:$B$2"),
  );

  console.log("\n11) Fallback automático: API de workbook recusada → download + leitura local");
  const tokenCache = resolve(tmpDir, "fake-token.json");
  writeFileSync(
    tokenCache,
    JSON.stringify({
      access_token: "fake-token",
      refresh_token: "fake-refresh",
      expires_at: Date.now() + 3_600_000,
    }),
  );
  const fakeConfig = {
    ...config,
    source: "graph" as const,
    auth: "device_code" as const,
    tokenCache,
    workbookPath: "/Documentos/PLANILHA.xlsx",
    driveId: "",
    itemId: "",
    shareUrl: "",
    localFile: "",
    graphFallback: true,
  };

  const realFetch = globalThis.fetch;
  const calls: string[] = [];
  let workbookApiDenied = true;

  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
    calls.push(url);

    if (url.includes("/workbook/")) {
      // Exactly what a personal account or app-only token gets back.
      if (workbookApiDenied) {
        return new Response(
          JSON.stringify({ error: { code: "accessDenied", message: "The workbook API is not available." } }),
          { status: 403, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(JSON.stringify({ value: [] }), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    }

    if (url.includes("@microsoft.graph.downloadUrl")) {
      return new Response(
        JSON.stringify({
          id: "ITEM-1",
          "@microsoft.graph.downloadUrl": "https://download.invalid/PLANILHA.xlsx",
        }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }

    if (url.startsWith("https://download.invalid/")) {
      return new Response(readFileSync(fixturePath), {
        status: 200,
        headers: {
          "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "content-length": String(readFileSync(fixturePath).byteLength),
        },
      });
    }

    // Drive-item resolution.
    return new Response(
      JSON.stringify({
        id: "ITEM-1",
        name: "PLANILHA.xlsx",
        size: readFileSync(fixturePath).byteLength,
        lastModifiedDateTime: "2026-09-01T00:00:00Z",
        file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        parentReference: { driveId: "DRIVE-1" },
      }),
      { status: 200, headers: { "content-type": "application/json" } },
    );
  }) as typeof globalThis.fetch;

  try {
    const tokens = new TokenProvider(fakeConfig);
    const graphSource = await GraphWorkbookSource.connect(fakeConfig, tokens);
    check("resolveu o arquivo sem chamar a API de workbook", calls.some((u) => u.includes("root:")), calls.length);

    const degradedGrid = await graphSource.readRange("junho26");
    check("não quebrou: caiu para o download", graphSource.degraded, graphSource.degraded);
    check(
      "motivo do fallback registrado",
      graphSource.fallbackInfo?.includes("accessDenied") === true,
      graphSource.fallbackInfo,
    );
    check(
      "tentou primeiro a API de workbook",
      calls.some((u) => u.includes("/workbook/worksheets")),
      calls.filter((u) => u.includes("workbook")),
    );
    check(
      "baixou o arquivo depois da recusa",
      calls.some((u) => u.startsWith("https://download.invalid/")),
      calls.filter((u) => u.includes("download")),
    );
    check("dados lidos corretamente no modo degradado", detectBlocks(degradedGrid).blocks.length === 4);
    const degradedNormalized = normalizeTransactions(degradedGrid, { aba: "junho26" });
    check("normalização igual no modo degradado", degradedNormalized.transacoes.length === 8);
    check(
      "descrição descreve a origem como graph",
      (await graphSource.describe()).source === "graph",
      (await graphSource.describe()).source,
    );

    // The next call must reuse the already-downloaded copy, not re-hit the API.
    const before = calls.filter((u) => u.includes("/workbook/")).length;
    await graphSource.readRange("Lançamentos");
    check(
      "não volta a bater na API de workbook",
      calls.filter((u) => u.includes("/workbook/")).length === before,
      { antes: before, depois: calls.filter((u) => u.includes("/workbook/")).length },
    );

    // With the fallback disabled the original error must surface untouched.
    workbookApiDenied = true;
    globalThis.fetch = (async (input: string | URL | Request) => {
      const url = typeof input === "string" ? input : input instanceof URL ? input.href : input.url;
      if (url.includes("/workbook/")) {
        return new Response(
          JSON.stringify({ error: { code: "accessDenied", message: "nope" } }),
          { status: 403, headers: { "content-type": "application/json" } },
        );
      }
      return new Response(
        JSON.stringify({ id: "ITEM-1", name: "x.xlsx", parentReference: { driveId: "DRIVE-1" } }),
        { status: 200, headers: { "content-type": "application/json" } },
      );
    }) as typeof globalThis.fetch;
    const strictTokens = new TokenProvider(fakeConfig);
    const strictSource = await GraphWorkbookSource.connect(
      { ...fakeConfig, graphFallback: false },
      strictTokens,
    );
    const strictError = await strictSource.readRange("junho26").catch((error: Error) => error);
    check(
      "EXCEL_GRAPH_FALLBACK=0 propaga o erro original",
      strictError instanceof GraphError && strictError.graphCode === "accessDenied",
      strictError instanceof Error ? strictError.message : strictError,
    );
  } finally {
    globalThis.fetch = realFetch;
  }

  console.log("\n12) Login pelo navegador (authorization code + PKCE) — encanamento");
  const authConfig = { ...fakeConfig, scopes: ["offline_access", "Files.Read.All", "Sites.Read.All"] };

  const runBrowserFlow = async (
    forge: (redirectUri: string, state: string) => void,
  ): Promise<{ error: unknown; url: URL | null }> => {
    let authorize: URL | null = null;
    const error = await signInWithBrowser(authConfig, {
      port: 0, // let the OS pick a free port
      openBrowser: false,
      timeoutSeconds: 15,
      onUrl: (url) => {
        authorize = new URL(url);
        const redirect = authorize.searchParams.get("redirect_uri") ?? "";
        const state = authorize.searchParams.get("state") ?? "";
        // Fire the forged callback once the listener is up.
        setTimeout(() => forge(redirect, state), 50);
      },
    }).then(
      () => null,
      (err: unknown) => err,
    );
    // Give the server a moment to notice the request arrived.
    await new Promise((r) => setTimeout(r, 300));
    return { error, url: authorize };
  };

  const forged = await runBrowserFlow((redirect, state) => {
    void fetch(`${redirect}?code=forjado&state=${state}-errado`).catch(() => {});
  });
  check("PKCE S256 presente na URL", forged.url?.searchParams.get("code_challenge_method") === "S256");
  check("response_type=code", forged.url?.searchParams.get("response_type") === "code");
  check("redirect é loopback local", /^http:\/\/localhost:\d+\/callback$/.test(forged.url?.searchParams.get("redirect_uri") ?? ""), forged.url?.searchParams.get("redirect_uri"));
  check("scopes enviados", forged.url?.searchParams.get("scope")?.includes("Files.Read.All") === true);
  check(
    "state forjado é rejeitado (segurança)",
    forged.error instanceof AuthError && /State inválido/.test(forged.error.message),
    forged.error instanceof Error ? forged.error.message : forged.error,
  );

  const denied = await runBrowserFlow((redirect) => {
    void fetch(`${redirect}?error=access_denied&error_description=usuario+recusou`).catch(() => {});
  });
  check(
    "recusa do usuário é tratada",
    denied.error instanceof AuthError && /Autorização recusada/.test(denied.error.message),
    denied.error instanceof Error ? denied.error.message : denied.error,
  );

  // Same state, bogus code → this really calls Microsoft's token endpoint, which
  // proves the exchange is shaped correctly (PKCE + redirect_uri + grant_type).
  const exchange = await runBrowserFlow((redirect, state) => {
    void fetch(`${redirect}?code=um-codigo-invalido&state=${state}`).catch(() => {});
  });
  check(
    "troca do code chega ao Microsoft e falha como invalid_grant",
    exchange.error instanceof AuthError &&
      /Troca do code por token falhou/.test(exchange.error.message) &&
      /invalid_grant|AADSTS/.test(exchange.error.message),
    exchange.error instanceof Error ? exchange.error.message.slice(0, 200) : exchange.error,
  );

  rmSync(tmpDir, { recursive: true, force: true });

  console.log("\n═══════════════════════════════════════════════════════");
  if (failures === 0) {
    console.log("✅ todos os testes passaram");
  } else {
    console.log(`❌ ${failures} teste(s) falharam`);
    process.exit(1);
  }
}

main().catch((error: unknown) => {
  console.error("\n❌ smoke test crashed:", error instanceof Error ? error.stack : error);
  process.exit(1);
});

