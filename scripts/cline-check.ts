#!/usr/bin/env node
/**
 * Verifies that the excel MCP server works under the exact conditions Cline
 * launches it in: reads the real Cline settings file, spawns the configured
 * command/args/env, from a FOREIGN cwd (Cline may run from `/` or a workspace),
 * and drives it over the real MCP protocol.
 *
 * Usage: npm run cline-check
 */
import { execFileSync } from "node:child_process";
import { readFileSync, rmSync } from "node:fs";
import { homedir } from "node:os";
import { resolve } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { buildFixture } from "./fixture.js";

const SETTINGS_PATH =
  process.env.CLINE_MCP_SETTINGS_PATH?.trim() ||
  resolve(homedir(), ".cline", "data", "settings", "cline_mcp_settings.json");

let failures = 0;
function check(label: string, condition: boolean, detail?: unknown): void {
  if (condition) {
    console.log(`  ✅ ${label}`);
  } else {
    failures++;
    console.log(`  ❌ ${label}`);
    if (detail !== undefined) console.log(`     obtido: ${JSON.stringify(detail).slice(0, 400)}`);
  }
}

interface ClineEntry {
  transport?: {
    type?: string;
    command?: string;
    args?: string[];
    env?: Record<string, string>;
    cwd?: string;
  };
  type?: string;
  command?: string;
  args?: string[];
  env?: Record<string, string>;
  cwd?: string;
  disabled?: boolean;
  timeout?: number;
}

interface Normalized {
  command?: string;
  args: string[];
  env: Record<string, string>;
  cwd?: string;
  disabled?: boolean;
  timeout?: number;
}

/** Accepts both the flat and the nested-transport entry forms. */
function normalize(entry: ClineEntry): Normalized {
  const source = entry.transport ?? entry;
  return {
    command: source.command,
    args: source.args ?? [],
    env: source.env ?? {},
    cwd: source.cwd,
    disabled: entry.disabled,
    timeout: entry.timeout,
  };
}

function textOf(result: unknown): string {
  const content = (result as { content?: Array<{ type: string; text?: string }> }).content ?? [];
  return content.map((part) => part.text ?? "").join("\n");
}

function jsonOf(result: unknown): Record<string, unknown> | null {
  const raw = textOf(result);
  const fence = /```json\n([\s\S]*?)\n```/.exec(raw);
  try {
    return JSON.parse(fence ? fence[1]! : raw) as Record<string, unknown>;
  } catch {
    return null;
  }
}

async function main(): Promise<void> {
  console.log("excel-mcp × Cline check\n═══════════════════════════════════════════════════════");
  console.log(`settings: ${SETTINGS_PATH}\n`);

  let entry: ClineEntry | undefined;
  try {
    const settings = JSON.parse(readFileSync(SETTINGS_PATH, "utf8")) as {
      mcpServers?: Record<string, ClineEntry>;
    };
    entry = settings.mcpServers?.excel;
  } catch (error) {
    console.log(`❌ could not read the Cline settings file: ${(error as Error).message}`);
    process.exit(1);
  }

  check("servidor 'excel' presente em mcpServers", Boolean(entry));
  if (!entry) process.exit(1);

  const config = normalize(entry);
  check("tem command", Boolean(config.command), config.command);
  check(
    "args apontam para o build",
    config.args.some((flag) => flag.endsWith("index.js")),
    config.args,
  );
  check("não está desabilitado", config.disabled !== true);
  if (!config.command) process.exit(1);

  console.log("\n1) O binário do node indicado existe e roda");
  try {
    const version = execFileSync(config.command, ["--version"], { encoding: "utf8" }).trim();
    check(`command executa (${version})`, /^v\d+/.test(version), version);
  } catch (error) {
    check("command executa", false, (error as Error).message);
    process.exit(1);
  }

  console.log("\n2) O entrypoint do servidor existe");
  const entryPath = config.args[config.args.length - 1] ?? "";
  let entryExists = true;
  try {
    readFileSync(entryPath);
  } catch {
    entryExists = false;
  }
  check("dist/src/index.js existe (rode `npm run build`)", entryExists, entryPath);
  if (!entryExists) process.exit(1);

  console.log("\n3) Spawn idêntico ao do Cline, a partir de um cwd alheio");
  const projectRoot = resolve(import.meta.dirname, "..");
  const tmpDir = resolve(projectRoot, ".tokens", "tmp");
  const fixture = await buildFixture(resolve(tmpDir, "PLANILHA-cline-check.xlsx"));
  console.log("   cwd: /   (o servidor não pode depender do cwd)");
  console.log("   fixture apontada por EXCEL_LOCAL_FILE\n");

  const client = new Client({ name: "cline-mcp-check", version: "0.1.0" });
  const transport = new StdioClientTransport({
    command: config.command,
    args: config.args,
    env: { ...config.env, EXCEL_SOURCE: "local", EXCEL_LOCAL_FILE: fixture },
    cwd: "/", // deliberately not the project directory
    stderr: "pipe",
  });

  await client.connect(transport);
  check("handshake", client.getServerVersion()?.name === "excel-mcp", client.getServerVersion());

  const { tools } = await client.listTools();
  console.log(`   ${tools.length} ferramentas: ${tools.map((tool) => tool.name).slice(0, 4).join(", ")}, …`);
  check("11 ferramentas expostas", tools.length === 11, tools.length);

  const diagnostics = jsonOf(await client.callTool({ name: "excel_diagnostics", arguments: {} }));
  check("excel_diagnostics responde", diagnostics?.fonte === "local", diagnostics?.fonte);

  const info = jsonOf(await client.callTool({ name: "excel_workbook_info", arguments: {} }));
  check("excel_workbook_info lê a fixture com cwd=/", info?.name === "PLANILHA-cline-check.xlsx", info?.name);

  const report = jsonOf(
    await client.callTool({
      name: "excel_build_report",
      arguments: {
        aba: "junho26",
        agruparPor: ["categoria"],
        metricas: [{ agg: "sum", column: "valor", as: "total" }],
      },
    }),
  ) as { totais?: { total?: number } } | null;
  check(
    "relatório correto mesmo com cwd=/",
    Math.abs(Number(report?.totais?.total) - 13314.52) < 0.001,
    report?.totais,
  );

  await client.close();
  rmSync(tmpDir, { recursive: true, force: true });

  console.log("\n═══════════════════════════════════════════════════════");
  console.log(
    failures === 0
      ? "✅ pronto: o Cline consegue fazer spawn e usar estas ferramentas"
      : `❌ ${failures} verificação(ões) falharam`,
  );
  process.exit(failures === 0 ? 0 : 1);
}

main().catch((error: unknown) => {
  console.error("\n❌ cline-check falhou:", error instanceof Error ? error.stack : error);
  process.exit(1);
});
