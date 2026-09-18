import { createHash, randomBytes } from "node:crypto";
import { execFile } from "node:child_process";
import { createServer, type Server } from "node:http";
import type { Config } from "./config.js";
import { AUTHORITY, AuthError, expiryFromNow, type TokenRecord } from "./auth.js";

export interface BrowserSignInOptions {
  /** Preferred loopback port; 0 lets the OS pick a free one. */
  port?: number;
  /** How long to wait for the user to finish in the browser. */
  timeoutSeconds?: number;
  /** Try to open the default browser (set false when headless). */
  openBrowser?: boolean;
  /** Called once the server is listening, with the URL the user must visit. */
  onUrl?: (url: string) => void;
}

function base64url(input: Buffer): string {
  return input.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/** Minimal, dependency-free localhost catcher for the OAuth redirect. */
export class LoopbackReceiver {
  private server: Server | null = null;
  port = 0;

  private constructor() {}

  static async start(port: number, handler: (url: URL) => { status: number; html: string }): Promise<LoopbackReceiver> {
    const receiver = new LoopbackReceiver();
    const server = createServer((req, res) => {
      const url = new URL(req.url ?? "/", `http://127.0.0.1:${receiver.port}`);
      const { status, html } = handler(url);
      res.writeHead(status, { "content-type": "text/html; charset=utf-8" });
      res.end(html);
    });
    receiver.server = server;

    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      // 127.0.0.1 only: never expose the callback on the network.
      server.listen(port, "127.0.0.1", () => resolve());
    });

    const address = server.address();
    receiver.port = typeof address === "object" && address ? address.port : port;
    return receiver;
  }

  async close(): Promise<void> {
    if (!this.server) return;
    await new Promise<void>((resolve) => this.server?.close(() => resolve()));
    this.server = null;
  }
}

function page(title: string, body: string, accent = "#16a34a"): string {
  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<title>${title}</title><style>
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#0f1115;color:#e8eaed;
display:flex;align-items:center;justify-content:center;height:100vh;margin:0}
.card{max-width:520px;padding:40px;border-radius:16px;background:#171a21;border:1px solid #262b36}
h1{margin:0 0 12px;font-size:20px;color:${accent}}p{margin:0;line-height:1.6;color:#a8b0bd}
code{background:#0f1115;padding:2px 6px;border-radius:6px;color:#e8eaed}
</style></head><body><div class="card"><h1>${title}</h1><p>${body}</p></div></body></html>`;
}

/**
 * Authorization-code + PKCE flow over a loopback redirect — the flow Microsoft
 * recommends for desktop apps.
 *
 * It exists because the device-code flow forces the user to retype a code on a
 * completely different web page, which is both error-prone and, on accounts where
 * two device-login front-ends exist (AAD vs consumer), prone to "the code has
 * expired" rejections even while the token endpoint still reports the code as
 * pending. Here the browser hands the code straight back to us, so there is
 * nothing to copy and nothing to mistype.
 */
export async function signInWithBrowser(
  config: Config,
  options: BrowserSignInOptions = {},
): Promise<TokenRecord> {
  const { port = 53_123, timeoutSeconds = 300, openBrowser = process.platform === "darwin" } = options;

  const verifier = base64url(randomBytes(32));
  const challenge = base64url(createHash("sha256").update(verifier).digest());
  const state = base64url(randomBytes(16));

  let settled = false;
  let onHit: (url: URL) => void = () => {};

  const handleCallback = (url: URL): { status: number; html: string } => {
    if (url.pathname !== "/callback") {
      return { status: 404, html: page("Rota desconhecida", "Nada para ver aqui.", "#dc2626") };
    }
    if (settled) {
      return { status: 200, html: page("Já recebido", "Este login já foi processado. Pode fechar esta aba.") };
    }
    settled = true;
    onHit(url);
    return {
      status: 200,
      html: url.searchParams.get("error")
        ? page("Login não concluído", "Você pode fechar esta aba e voltar ao terminal.", "#dc2626")
        : page("Login concluído ✅", "Pode fechar esta aba e voltar ao terminal."),
    };
  };

  // Ports are irrelevant to Microsoft for loopback redirects, so fall back to an
  // OS-assigned one if the preferred port is busy.
  let receiver: LoopbackReceiver;
  try {
    receiver = await LoopbackReceiver.start(port, handleCallback);
  } catch (error) {
    if ((error as { code?: string }).code !== "EADDRINUSE") throw error;
    receiver = await LoopbackReceiver.start(0, handleCallback);
  }

  const redirectUri = `http://localhost:${receiver.port}/callback`;
  const authorizeUrl = `${AUTHORITY}/${config.tenantId}/oauth2/v2.0/authorize?${new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: config.scopes.join(" "),
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account",
  }).toString()}`;

  options.onUrl?.(authorizeUrl);
  if (openBrowser) {
    execFile("open", [authorizeUrl], () => {
      /* if it fails the user still has the printed URL */
    });
  }

  const result = await new Promise<URL>((resolve, reject) => {
    const timer = setTimeout(() => {
      if (!settled) reject(new AuthError(`Ninguém concluiu o login em ${timeoutSeconds}s.`));
    }, timeoutSeconds * 1000);

    onHit = (url) => {
      clearTimeout(timer);
      resolve(url);
    };
  }).finally(() => receiver.close());

  const error = result.searchParams.get("error");
  if (error) {
    throw new AuthError(
      `Autorização recusada (${error}): ${result.searchParams.get("error_description") ?? ""}`,
    );
  }
  if (result.searchParams.get("state") !== state) {
    throw new AuthError("State inválido na resposta de autorização — login abortado por segurança.");
  }
  const code = result.searchParams.get("code");
  if (!code) throw new AuthError("A resposta de autorização não trouxe um code.");

  const res = await fetch(`${AUTHORITY}/${config.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "authorization_code",
      client_id: config.clientId,
      code,
      redirect_uri: redirectUri,
      code_verifier: verifier,
      scope: config.scopes.join(" "),
    }).toString(),
  });
  const json = (await res.json()) as Record<string, unknown>;
  if (!res.ok || json.error) {
    throw new AuthError(
      `Troca do code por token falhou (${String(json.error ?? res.status)}): ${String(
        json.error_description ?? "",
      )}`,
    );
  }

  return {
    access_token: String(json.access_token),
    refresh_token: json.refresh_token ? String(json.refresh_token) : undefined,
    expires_at: expiryFromNow(json.expires_in),
    scope: json.scope ? String(json.scope) : undefined,
  };
}
