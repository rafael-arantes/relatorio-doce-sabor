#!/usr/bin/env node
/**
 * Servidor do relatório financeiro.
 *
 * Serve os arquivos estáticos de `report/` e mantém o `data.json` atualizado:
 *   - ao subir (primeira carga);
 *   - em intervalos regulares (`REFRESH_MINUTES`, padrão 60);
 *   - sob demanda, via `POST /refresh` (o botão "Atualizar" do relatório usa isso).
 *
 * Este é o processo que precisa do token do Microsoft Graph (`.tokens/graph.json`
 * ou as variáveis do `.env`) — por isso roda num servidor seu (Coolify), não no
 * navegador do cliente.
 *
 * Env:
 *   PORT              porta (padrão 8787)
 *   REFRESH_MINUTES   intervalo entre atualizações (padrão 60)
 *   REFRESH_TOKEN     segredo opcional exigido em `POST /refresh?token=…`
 */
import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { readFileSync, existsSync } from "node:fs";
import { join, resolve } from "node:path";
import { loadConfig, findProjectRoot } from "../src/config.js";
import { buildReportData, type ReportData } from "./aggregate.js";

const ROOT = process.env.REPORT_ROOT || resolve(findProjectRoot(), "report");
const PORT = Number(process.env.PORT || 8787);
const REFRESH_MINUTES = Math.max(1, Number(process.env.REFRESH_MINUTES || 60));
const REFRESH_TOKEN = process.env.REFRESH_TOKEN || "";

const MIME: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

let cached: ReportData | null = null;
let cachedAt = 0;
let inflight: Promise<void> | null = null;

async function refresh(): Promise<void> {
  if (inflight) return inflight;
  inflight = (async () => {
    try {
      cached = await buildReportData(loadConfig());
      cachedAt = Date.now();
      console.log(`[report] dados atualizados em ${new Date().toISOString()}`);
    } catch (error) {
      console.error("[report] falha ao atualizar dados:", error instanceof Error ? error.message : error);
    } finally {
      inflight = null;
    }
  })();
  return inflight;
}

function send(res: ServerResponse, status: number, body: string | Buffer, type: string): void {
  res.writeHead(status, {
    "content-type": type,
    "cache-control": "no-store",
    "access-control-allow-origin": "*",
  });
  res.end(body);
}

async function handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
  const url = new URL(req.url ?? "/", `http://localhost:${PORT}`);
  const pathname = url.pathname;

  if (req.method === "POST" && pathname === "/refresh") {
    if (REFRESH_TOKEN && url.searchParams.get("token") !== REFRESH_TOKEN) {
      send(res, 401, JSON.stringify({ ok: false, error: "token inválido" }), "application/json; charset=utf-8");
      return;
    }
    await refresh();
    send(res, 200, JSON.stringify({ ok: true, geradoEm: cached?.geradoEm ?? null }), "application/json; charset=utf-8");
    return;
  }

  if (pathname === "/data.json") {
    // Se estiver desatualizado, atualiza antes de servir (uma chamada por processo).
    const ageMinutes = cachedAt ? (Date.now() - cachedAt) / 60_000 : Infinity;
    if (!cached || ageMinutes >= REFRESH_MINUTES) await refresh();
    send(res, 200, JSON.stringify(cached ?? { meses: [], porMes: {} }), "application/json; charset=utf-8");
    return;
  }

  // Arquivos estáticos.
  const safe = pathname === "/" ? "/index.html" : pathname;
  const file = resolve(join(ROOT, decodeURIComponent(safe)));
  if (!file.startsWith(ROOT) || !existsSync(file)) {
    send(res, 404, "não encontrado", "text/plain; charset=utf-8");
    return;
  }
  const ext = safe.slice(safe.lastIndexOf("."));
  send(res, 200, readFileSync(file), MIME[ext] ?? "application/octet-stream");
}

const server = createServer((req, res) => {
  handle(req, res).catch((error) => {
    console.error("[report] erro na requisição:", error);
    if (!res.headersSent) send(res, 500, "erro interno", "text/plain; charset=utf-8");
  });
});

server.listen(PORT, () => {
  console.log(`[report] servindo ${ROOT} em http://localhost:${PORT}`);
  console.log(`[report] atualização automática a cada ${REFRESH_MINUTES} min`);
  console.log(`[report] atualização manual: POST /refresh${REFRESH_TOKEN ? "?token=…" : ""}`);
  void refresh();
  setInterval(() => void refresh(), REFRESH_MINUTES * 60_000);
});
