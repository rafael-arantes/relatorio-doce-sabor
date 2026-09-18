#!/usr/bin/env node
/**
 * One-time interactive sign-in for the delegated (device-code) auth flow.
 *
 * Designed to survive being run in a captured/background terminal:
 *   - the code is written to .tokens/pending-device-code.json *immediately*, so
 *     you can re-read it even if the output scrolled away;
 *   - a heartbeat every ~10s proves the process is still waiting for you;
 *   - if the code expires mid-way, a fresh one is requested automatically;
 *   - `--resume` reuses a pending code, `--fresh` always requests a new one.
 *
 * Usage: npm run login [-- --resume] [-- --fresh] [-- --no-renew]
 */
import { existsSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { loadConfig, GRAPH_CLI_CLIENT_ID, type Config } from "../src/config.js";
import { AuthError, TokenProvider, persistToken, readCachedToken, type TokenRecord } from "../src/auth.js";
import { signInWithBrowser } from "../src/auth-code.js";
import { pollDeviceCode, startDeviceCode, type DeviceCodeChallenge } from "../src/device-code.js";
import { resolveDriveItem } from "../src/sources/graph-api.js";

const argv = process.argv.slice(2);
const flag = (name: string): boolean => argv.includes(name);

function pendingPath(tokenCache: string): string {
  return tokenCache.replace(/[^/]+$/, "pending-device-code.json");
}

function savePending(path: string, challenge: DeviceCodeChallenge): void {
  writeFileSync(path, JSON.stringify({ ...challenge, created_at: Date.now() }, null, 2), {
    mode: 0o600,
  });
}

function loadPending(path: string): (DeviceCodeChallenge & { created_at?: number }) | null {
  if (!existsSync(path)) return null;
  try {
    const parsed = JSON.parse(readFileSync(path, "utf8")) as DeviceCodeChallenge & {
      created_at?: number;
    };
    return parsed.device_code ? parsed : null;
  } catch {
    return null;
  }
}

function describe(config: Config, challenge: DeviceCodeChallenge): void {
  const complete = challenge.verification_uri_complete;
  const deadline = new Date(Date.now() + Number(challenge.expires_in || 900) * 1000);
  console.log("──────────────────────────────────────────────────────────────");
  console.log("  Sign in with the Microsoft account that can open the workbook");
  console.log("──────────────────────────────────────────────────────────────");
  console.log(`\n  1. Open:   ${complete ?? challenge.verification_uri}`);
  if (!complete) console.log(`  2. Enter:  ${challenge.user_code}`);
  console.log(
    `\n  Valid until ${deadline.toLocaleTimeString()} (${Math.round(
      Number(challenge.expires_in || 900) / 60,
    )} minutos).`,
  );
  console.log("  If it expires, a new code is requested automatically — keep this window open.");
  console.log(`  Scopes:   ${config.scopes.join(", ")}`);
  console.log(`  Saved to: ${pendingPath(config.tokenCache)}`);
  console.log("\n  Waiting for you to finish signing in (Ctrl+C to abort)…\n");
}

async function finalize(config: Config, token: TokenRecord): Promise<void> {
  persistToken(config, token);
  const pending = pendingPath(config.tokenCache);
  if (existsSync(pending)) rmSync(pending);
  console.log(`\n✅ Credentials cached at ${config.tokenCache}`);

  const tokens = new TokenProvider(config);
  try {
    const item = await resolveDriveItem(config, tokens);
    console.log(`✅ Workbook reachable: ${item.name}`);
    console.log(`   drive=${item.parentReference?.driveId} item=${item.id}`);
    console.log("   Put those in .env as EXCEL_DRIVE_ID / EXCEL_ITEM_ID to skip the share-link lookup.");
  } catch (error) {
    const hint = (error as { hint?: string }).hint;
    console.log("⚠️  Signed in, but the workbook could not be resolved.");
    console.log(`   ${error instanceof Error ? error.message : String(error)}`);
    if (hint) console.log(`   💡 ${hint}`);
  }
}

async function main(): Promise<void> {
  const config = loadConfig();
  const pending = pendingPath(config.tokenCache);
  const renew = !flag("--no-renew");

  if (config.auth === "client_secret") {
    console.log("EXCEL_AUTH=client_secret: acquiring an app-only token…");
    await new TokenProvider(config).getToken();
    console.log(`✅ App-only token cached at ${config.tokenCache}`);
    return;
  }

  const existing = readCachedToken(config);
  if (existing?.refresh_token) {
    console.log(`An existing credential cache was found at ${config.tokenCache}.`);
    console.log("Signing in again replaces it.\n");
  }

  // Browser flow (authorization code + PKCE) — needs a redirect URI registered
  // on the app, so it only works with YOUR OWN app registration, never with the
  // first-party "Microsoft Graph Command Line Tools" client.
  if (flag("--browser")) {
    if (config.clientId === GRAPH_CLI_CLIENT_ID) {
      console.error(
        "❌ O fluxo pelo navegador precisa de um app próprio, porque o cliente público " +
          "padrão não tem `http://localhost` como redirect URI registrado.\n\n" +
          "   Registre um app no Microsoft Entra (documentado no README, seção 'mode B') com:\n" +
          "   • Allow public client flows = Yes\n" +
          "   • Redirect URI: http://localhost\n" +
          "   • Permissão delegada: Files.Read.All\n" +
          "   e coloque o Application (client) ID em EXCEL_CLIENT_ID. Depois rode `npm run login -- --browser`.\n\n" +
          "   Ou, sem registrar nada, use o fluxo por código de dispositivo: `npm run login` (padrão).",
      );
      process.exit(1);
    }
    console.log("Login pelo navegador (authorization code + PKCE, redirect para localhost).\n");
    const token = await signInWithBrowser(config, {
      timeoutSeconds: 300,
      openBrowser: !flag("--no-browser"),
      onUrl: (url) => {
        console.log("Se o navegador não abrir sozinho, cole esta URL manualmente:\n");
        console.log(`  ${url}\n`);
        console.log("  Autorize com a conta Microsoft que tem acesso à planilha.");
        console.log("  Ao final, o navegador volta sozinho para 127.0.0.1 e o token é salvo.\n");
        console.log("  Aguardando a autorização (Ctrl+C para abortar)…");
      },
    });
    await finalize(config, token);
    return;
  }

  // Device-code flow (default): works with the first-party client, no registration.
  console.log("Login por código de dispositivo.\n");
  console.log(
    "Dica: use uma janela anônima/privada e entre com a conta Microsoft PESSOAL " +
      "que tem acesso à planilha (não a conta de trabalho), para evitar que uma sessão " +
      "antiga do navegador atrapalhe.\n",
  );
  let challenge: DeviceCodeChallenge | null = null;
  if (!flag("--fresh")) {
    const stored = loadPending(pending);
    if (stored) {
      const age = Date.now() - (stored.created_at ?? 0);
      if (age < Number(stored.expires_in || 900) * 1000) {
        console.log(
          `Reusing the pending code (created ${Math.round(age / 60_000)} min ago). ` +
            "Use `npm run login -- --fresh` for a new one.\n",
        );
        challenge = stored;
      }
    }
  }
  challenge ??= await startDeviceCode(config);
  savePending(pending, challenge);
  describe(config, challenge);

  const token = await pollDeviceCode(config, challenge, {
    onWaiting: (elapsed, remaining) => {
      const mm = String(Math.floor(remaining / 60)).padStart(2, "0");
      const ss = String(remaining % 60).padStart(2, "0");
      console.log(`  … waiting (${elapsed}s elapsed, ${mm}:${ss} until this code expires)`);
    },
    onBadCode: () => {
      console.log(
        "  ⚠️  The browser rejected that code (typo, or a code from an earlier run).\n" +
          `      Use exactly this one: ${challenge?.user_code}  — still waiting…`,
      );
    },
    onExpired: renew
      ? async () => {
          console.log("\n  ⌛ That code expired. Requesting a fresh one…\n");
          const fresh = await startDeviceCode(config);
          savePending(pending, fresh);
          describe(config, fresh);
          return fresh;
        }
      : undefined,
  });

  await finalize(config, token);
}

main().catch((error: unknown) => {
  if (error instanceof AuthError) {
    console.error(`\n❌ ${error.message}`);
    if (error.hint) console.error(`   ${error.hint}`);
  } else {
    console.error("\n❌", error instanceof Error ? error.message : error);
  }
  process.exit(1);
});
