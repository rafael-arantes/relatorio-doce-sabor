import { chmodSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import type { Config } from "./config.js";
import { debugLog } from "./config.js";

export const GRAPH = "https://graph.microsoft.com";
export const AUTHORITY = "https://login.microsoftonline.com";

export interface TokenRecord {
  access_token: string;
  refresh_token?: string;
  expires_at: number;
  scope?: string;
}

export class AuthError extends Error {
  constructor(message: string, readonly hint?: string) {
    super(message);
    this.name = "AuthError";
  }
}

export function readCache(path: string): TokenRecord | null {
  if (!existsSync(path)) return null;
  try {
    const parsed = JSON.parse(readFileSync(path, "utf8")) as TokenRecord;
    if (typeof parsed.access_token !== "string") return null;
    return parsed;
  } catch {
    return null;
  }
}

export function writeCache(path: string, record: TokenRecord): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, JSON.stringify(record, null, 2), { mode: 0o600 });
  try {
    chmodSync(path, 0o600);
  } catch {
    /* best effort on filesystems without chmod */
  }
}

export function expiryFromNow(expiresInSeconds: unknown): number {
  const seconds = Number.isFinite(expiresInSeconds) ? Number(expiresInSeconds) : 3600;
  // 60s safety margin so a token never expires mid-request.
  return Date.now() + Math.max(seconds - 60, 30) * 1000;
}

/**
 * Scope string for the /token endpoint.
 * - client credentials always uses `<resource>/.default`
 * - delegated scopes are sent as bare permission names (or absolute URLs as-is)
 */
export function scopeFor(config: Config): string {
  if (config.auth === "client_secret") return `${GRAPH}/.default`;
  return config.scopes.join(" ");
}

export function readCachedToken(config: Config): TokenRecord | null {
  return readCache(config.tokenCache);
}

/** True when a token cache exists on disk (fresh or not). */
export function hasCachedCredentials(path: string): boolean {
  const record = readCache(path);
  return Boolean(record && (record.refresh_token || record.access_token));
}

export function persistToken(config: Config, token: TokenRecord): void {
  writeCache(config.tokenCache, token);
}

/**
 * Acquires a Microsoft Graph access token, transparently refreshing it and
 * persisting rotations back to the token cache.
 *
 * Delegated (device_code) mode never prompts interactively from inside the MCP
 * server — a stdio server has no way to show a user code. Instead it throws an
 * actionable AuthError telling you to run `npm run login` once.
 */
export class TokenProvider {
  private cached: TokenRecord | null;
  private inflight: Promise<string> | null = null;

  constructor(private readonly config: Config) {
    this.cached = readCache(config.tokenCache);
  }

  /** True when a usable token is already cached (no network call needed). */
  hasFreshToken(): boolean {
    return Boolean(this.cached && this.cached.expires_at > Date.now());
  }

  async getToken(): Promise<string> {
    const cached = this.cached;
    if (cached && cached.expires_at > Date.now()) {
      debugLog(this.config, "using cached access token");
      return cached.access_token;
    }
    // Collapse concurrent refreshes into a single network round-trip.
    this.inflight ??= this.mint().finally(() => {
      this.inflight = null;
    });
    return this.inflight;
  }

  private async mint(): Promise<string> {
    if (this.config.auth === "client_secret") {
      const token = await this.clientCredentialsToken();
      this.cached = token;
      persistToken(this.config, token);
      return token.access_token;
    }

    const refresh = this.cached?.refresh_token || this.config.refreshTokenEnv;
    if (!refresh) {
      throw new AuthError(
        "No cached Microsoft credentials for the delegated (device-code) flow.",
        "Run `npm run login` once in this folder — it prints a URL + code, you sign in, " +
          `and the refresh token is cached at ${this.config.tokenCache} for the server to reuse.`,
      );
    }

    const response = await this.postToken({
      grant_type: "refresh_token",
      refresh_token: refresh,
      client_id: this.config.clientId,
      scope: scopeFor(this.config),
    });

    if (response.error || !response.access_token) {
      throw new AuthError(
        `Refresh token rejected (${response.error ?? "unknown"}): ${response.error_description ?? ""}`,
        "The cached credentials expired or were revoked. Run `npm run login` again.",
      );
    }

    // Microsoft rotates refresh tokens; keep the newest one.
    const token: TokenRecord = {
      access_token: response.access_token,
      refresh_token: response.refresh_token ?? refresh,
      expires_at: expiryFromNow(response.expires_in),
      scope: response.scope,
    };
    this.cached = token;
    persistToken(this.config, token);
    return token.access_token;
  }

  private async clientCredentialsToken(): Promise<TokenRecord> {
    const response = await this.postToken({
      grant_type: "client_credentials",
      client_id: this.config.clientId,
      client_secret: this.config.clientSecret,
      scope: `${GRAPH}/.default`,
    });
    if (response.error || !response.access_token) {
      throw new AuthError(
        `client_credentials failed (${response.error ?? "unknown"}): ${response.error_description ?? ""}`,
        "App-only access needs an Azure app registration with the *application* " +
          "permissions Files.Read.All and Sites.Read.All granted by an admin.",
      );
    }
    return {
      access_token: response.access_token,
      expires_at: expiryFromNow(response.expires_in),
      scope: response.scope,
    };
  }

  private async postToken(body: Record<string, string>): Promise<TokenResponse> {
    const url = `${AUTHORITY}/${this.config.tenantId}/oauth2/v2.0/token`;
    const res = await fetch(url, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(body).toString(),
    });
    return (await res.json()) as TokenResponse;
  }
}

interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  error?: string;
  error_description?: string;
}

