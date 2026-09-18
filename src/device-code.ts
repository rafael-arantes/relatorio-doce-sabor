import type { Config } from "./config.js";
import { AUTHORITY, AuthError, expiryFromNow, type TokenRecord } from "./auth.js";

export interface DeviceCodeChallenge {
  device_code: string;
  user_code: string;
  verification_uri: string;
  verification_uri_complete?: string;
  message: string;
  interval: number;
  expires_in: number;
}

export interface PollHooks {
  /** Called every ~10s while waiting, so the user can see it's alive. */
  onWaiting?: (elapsedSeconds: number, remainingSeconds: number) => void;
  /** The user typed a wrong or already-used code; we keep waiting. */
  onBadCode?: (description: string) => void;
  /**
   * Called when the code expires. Return a fresh challenge to keep waiting, or
   * null to give up. Lets a long sign-in survive the 15-minute window.
   */
  onExpired?: () => Promise<DeviceCodeChallenge | null>;
}

/** Step 1 of the interactive device-code flow (used by scripts/login.ts). */
export async function startDeviceCode(config: Config): Promise<DeviceCodeChallenge> {
  const res = await fetch(`${AUTHORITY}/${config.tenantId}/oauth2/v2.0/devicecode`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: config.clientId,
      scope: config.scopes.join(" "),
    }).toString(),
  });
  const json = (await res.json()) as Record<string, unknown>;
  if (!res.ok || json.error) {
    throw new AuthError(
      `devicecode request failed: ${String(json.error ?? res.status)} ${String(
        json.error_description ?? "",
      )}`,
    );
  }
  return json as unknown as DeviceCodeChallenge;
}

/** Step 2: poll until the user finishes signing in. */
export async function pollDeviceCode(
  config: Config,
  initialChallenge: DeviceCodeChallenge,
  hooks: PollHooks = {},
): Promise<TokenRecord> {
  let challenge = initialChallenge;
  let lastHeartbeat = 0;

  for (;;) {
    const started = Date.now();
    const deadline = started + Number(challenge.expires_in || 900) * 1000;
    let interval = Math.max(Number(challenge.interval) || 5, 1) * 1000;

    while (Date.now() < deadline) {
      await new Promise((r) => setTimeout(r, interval));

      const res = await fetch(`${AUTHORITY}/${config.tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: { "content-type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "urn:ietf:params:oauth:grant-type:device_code",
          device_code: challenge.device_code,
          client_id: config.clientId,
        }).toString(),
      });
      const json = (await res.json()) as Record<string, unknown>;
      const error = json.error ? String(json.error) : "";
      const accessToken = json.access_token ? String(json.access_token) : "";

      if (accessToken) {
        return {
          access_token: accessToken,
          refresh_token: json.refresh_token ? String(json.refresh_token) : undefined,
          expires_at: expiryFromNow(json.expires_in),
          scope: json.scope ? String(json.scope) : undefined,
        };
      }

      if (error === "authorization_pending") {
        const elapsed = Math.round((Date.now() - started) / 1000);
        // A heartbeat roughly every 10s proves the process is still waiting.
        if (elapsed - lastHeartbeat >= 10) {
          lastHeartbeat = elapsed;
          hooks.onWaiting?.(elapsed, Math.max(Math.round((deadline - Date.now()) / 1000), 0));
        }
        continue;
      }
      if (error === "slow_down") {
        interval += 5000;
        continue;
      }
      if (error === "authorization_declined") {
        throw new AuthError("Sign-in was declined in the browser.");
      }
      if (error === "bad_verification_code") {
        // The typed code doesn't exist (typo, or a code from an earlier run).
        hooks.onBadCode?.(String(json.error_description ?? ""));
        continue;
      }
      if (error === "expired_token") {
        break; // handled by the renewal branch below
      }
      throw new AuthError(
        `Token polling failed (${error || res.status}): ${String(json.error_description ?? "")}`,
      );
    }

    // The code expired. Ask for a fresh one instead of dying, so a slow sign-in
    // can't leave the user stranded with a dead code.
    const renewed = await hooks.onExpired?.();
    if (!renewed) throw new AuthError("O código de dispositivo expirou antes do login ser concluído.");
    challenge = renewed;
    lastHeartbeat = 0;
  }
}
