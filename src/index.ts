#!/usr/bin/env node
import { main } from "./server.js";

main().catch((error: unknown) => {
  console.error("[excel-mcp] fatal:", error instanceof Error ? error.message : error);
  process.exit(1);
});
