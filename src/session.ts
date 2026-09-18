import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { loadConfig, validateConfig, type Config } from "./config.js";
import { openWorkbook } from "./open.js";
import type { WorkbookSource } from "./types.js";
import { LocalWorkbookSource } from "./sources/local.js";

export class Session {
  readonly config: Config;
  private source: WorkbookSource | null = null;
  private localOverride: string | null = null;

  constructor(config: Config = loadConfig()) {
    this.config = config;
  }

  /** Setup problems that would block every tool. */
  problems(): string[] {
    return validateConfig(this.config);
  }

  async getSource(
    options: { source?: "graph" | "graph_download" | "local"; arquivoLocal?: string } = {},
  ): Promise<WorkbookSource> {
    if (options.arquivoLocal && options.arquivoLocal !== this.localOverride) {
      await this.source?.close();
      this.source = await LocalWorkbookSource.open(this.config, options.arquivoLocal);
      this.localOverride = options.arquivoLocal;
      return this.source;
    }
    if (!this.source) {
      this.source = await openWorkbook(this.config, {
        forceSource: options.source,
        localFile: options.arquivoLocal,
      });
    }
    return this.source;
  }

  /** Drops the cached handle, e.g. after switching workbooks. */
  async reset(): Promise<void> {
    await this.source?.close();
    this.source = null;
    this.localOverride = null;
  }
}

export function ok(data: unknown, extraText?: string): CallToolResult {
  const text = JSON.stringify(data, null, 2);
  return {
    content: [
      { type: "text", text: extraText ? `${extraText}\n\n\`\`\`json\n${text}\n\`\`\`` : text },
    ],
  };
}

export function markdown(text: string, data?: unknown): CallToolResult {
  const payload = JSON.stringify(data ?? {}, null, 2);
  return {
    content: [
      { type: "text", text },
      { type: "text", text: `\`\`\`json\n${payload}\n\`\`\`` },
    ],
  };
}

export function fail(error: unknown): CallToolResult {
  const message = error instanceof Error ? error.message : String(error);
  const hint = (error as { hint?: string }).hint;
  return {
    isError: true,
    content: [{ type: "text", text: hint ? `${message}\n\n💡 ${hint}` : message }],
  };
}

/** Wraps a handler so any thrown error becomes a readable tool error. */
export function guarded<Args>(
  handler: (args: Args) => Promise<CallToolResult>,
): (args: Args) => Promise<CallToolResult> {
  return async (args: Args) => {
    try {
      return await handler(args);
    } catch (error) {
      return fail(error);
    }
  };
}
