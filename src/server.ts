import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig } from "./config.js";
import { Session } from "./session.js";
import { registerExportTools } from "./tools/export.js";
import { registerReadTools } from "./tools/read.js";
import { registerNormalizeTools } from "./tools/report.js";
import { registerWorkbookTools } from "./tools/workbook.js";

export const SERVER_INFO = {
  name: "excel-mcp",
  version: "0.1.0",
} as const;

export function createServer(session: Session = new Session(loadConfig())): McpServer {
  const server = new McpServer(SERVER_INFO, {
    instructions:
      "Read-only access to an Excel Online (.xlsx) workbook — either live through Microsoft Graph or from " +
      "a local/downloaded file. Typical flow: excel_workbook_info to see the sheets, excel_detect_blocks " +
      "or excel_column_profile to understand a messy sheet, then excel_build_report (or " +
      "excel_normalize_transactions) to produce the report. All tools are read-only except " +
      "excel_export_csv, which writes a CSV to disk. Dates and money in this workbook are Brazilian " +
      "formatted (dd/mm/yyyy, R$ 1.234,56) and are parsed accordingly.",
  });

  registerWorkbookTools(server, session);
  registerReadTools(server, session);
  registerNormalizeTools(server, session);
  registerExportTools(server, session);

  return server;
}

export async function main(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);

  // stdout belongs to the protocol; logs must go to stderr.
  console.error("[excel-mcp] ready on stdio");

  const shutdown = async (signal: string): Promise<void> => {
    console.error(`[excel-mcp] ${signal} received, shutting down`);
    try {
      await server.close();
    } finally {
      process.exit(0);
    }
  };
  process.on("SIGINT", () => void shutdown("SIGINT"));
  process.on("SIGTERM", () => void shutdown("SIGTERM"));
}
