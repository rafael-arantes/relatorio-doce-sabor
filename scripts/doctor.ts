#!/usr/bin/env node
/**
 * Setup checker: prints the effective configuration, validates it, and — when
 * possible — connects once and lists the workbook's sheets so you know the exact
 * sheet names to use.
 */
import { loadConfig } from "../src/config.js";
import { Session } from "../src/session.js";
import { hasCachedCredentials } from "../src/auth.js";

async function main(): Promise<void> {
  const config = loadConfig();
  const session = new Session(config);

  console.log("excel-mcp doctor\n═══════════════════════════════════════════════════════");
  console.log(`fonte            : ${config.source}`);
  console.log(`autenticação     : ${config.source === "graph" ? config.auth : "n/a (arquivo local)"}`);
  if (config.source === "graph") {
    console.log(`tenant           : ${config.tenantId}`);
    console.log(`client id        : ${config.clientId}`);
    const cached = hasCachedCredentials(config.tokenCache);
    console.log(`credenciais      : ${cached ? "presentes no cache ✅" : "ausentes — rode `npm run login` ❌"}`);
  }
  console.log(`alvo             :`);
  console.log(`  workbook path  : ${config.workbookPath || "—"}`);
  console.log(`  drive/item id  : ${config.driveId ? `${config.driveId}/${config.itemId}` : "—"}`);
  console.log(`  share url      : ${config.shareUrl ? "definida" : "—"}`);
  console.log(`  arquivo local  : ${config.localFile || "—"}`);
  console.log(`limites          : ${config.maxRows} linhas / ${config.maxCells} células por leitura`);

  const problems = session.problems();
  console.log("\nProblemas de configuração:");
  console.log(problems.length === 0 ? "  nenhum ✅" : problems.map((p) => `  - ${p}`).join("\n"));

  if (problems.length > 0) {
    console.log("\nConfigure o .env (veja .env.example) e rode de novo.");
    process.exit(1);
  }

  console.log("\nConectando ao workbook…");
  try {
    const source = await session.getSource();
    const description = await source.describe();
    console.log(`✅ ${description.name}  (${description.source})`);
    if (description.lastModified) console.log(`   modificado em: ${description.lastModified}`);
    if (description.webUrl) console.log(`   link: ${description.webUrl}`);
    console.log("   abas:");
    for (const sheet of description.sheets) {
      const size = sheet.rowCount && sheet.columnCount ? ` — ${sheet.rowCount} linhas × ${sheet.columnCount} colunas` : "";
      console.log(`     • ${sheet.name}${size}`);
    }
  } catch (error) {
    console.log(`❌ ${error instanceof Error ? error.message : String(error)}`);
    const hint = (error as { hint?: string }).hint;
    if (hint) console.log(`   💡 ${hint}`);
    process.exit(1);
  }
}

main().catch((error: unknown) => {
  console.error("❌", error instanceof Error ? error.message : error);
  process.exit(1);
});
