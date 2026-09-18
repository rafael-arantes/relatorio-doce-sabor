# excel-mcp

MCP server that reads an **Excel Online** workbook (OneDrive / SharePoint via Microsoft Graph) — or a
local `.xlsx` — and turns it into **custom reports** you can ask for in plain language from Claude Code
or Claude Desktop.

Read-only by design: the only tool that writes anything is `excel_export_csv`, and it writes a CSV next
to you, never back into the workbook.

It is built for the *actual* shape of real-world spreadsheets, not just tidy tables:

- **Disconnected blocks.** A sheet where each supplier/expense group has its own little table with its
  own `Data`/`Valor` columns is detected automatically (4-connected region growing).
- **Payment method hidden in the label.** `"ovo caipira (dinheiro)"` → `descricao: "ovo caipira"`,
  `conta: "Dinheiro"`.
- **Brazilian number/date formats.** `R$ 1.234,56` → `1234.56`; `12/06`, `12-06-26`, `12 de junho`,
  `2026-06-12` and Excel serial dates → `YYYY-MM-DD`. The year is inferred from the sheet name
  (`junho26`, `SETEMBRO`) when a date doesn't carry one.
- **Provenance and confidence.** Every normalized row knows which block and cell it came from, and how
  confident the parser is about it.

---

## Quick start (60 seconds, no Azure, no credentials)

Prove the whole thing works against any `.xlsx` you already have:

```bash
cd /Volumes/Storage/www/doce-sabor/excel-mcp
npm install
cp .env.example .env
# edit .env:
#   EXCEL_SOURCE=local
#   EXCEL_LOCAL_FILE=/absolute/path/to/PLANILHA.xlsx

npm run doctor     # shows the config, connects, lists every tab
npm run smoke      # offline self-test (fixture, 25 assertions)
npm run check      # boots the real MCP server over stdio and drives it
```

Then restart your MCP client and ask: *"What tabs does the spreadsheet have?"*

---

## Connecting to Excel Online (Microsoft Graph)

`EXCEL_LOCAL_FILE` is the escape hatch; the main path is Graph. Pick one **source mode** and one
**auth mode**.

### Source modes

| `EXCEL_SOURCE` | Reads via | Use when |
|---|---|---|
| `graph` | The Graph **workbook** API (`/workbook/worksheets/...`) | The file lives in a Microsoft 365 **work/school** account. Best fidelity: live cell values, no download. |
| `graph_download` | Graph to resolve + download, then exceljs locally | The file is in a **personal** Microsoft account, or you're using `EXCEL_AUTH=client_secret`. |
| `local` | A `.xlsx` on disk (or fetched from a share link) | Offline use, archived year-end files, quick testing. |
| `auto` (default) | `local` if a local file is set, otherwise `graph` | You don't want to think about it. |

> **Why `graph_download` exists.** Microsoft documents the workbook API as
> **not supported for personal Microsoft accounts** and **not supported for application
> (app-only) permissions** — only delegated work/school access. The `driveItem/content` API, by
> contrast, supports delegated personal accounts (`Files.Read`) *and* application tokens
> (`Files.Read.All`). So when the workbook API isn't available, that's the route that still works.
>
> In fact you usually don't have to choose: with `EXCEL_SOURCE=graph` the server tries the workbook
> API first and, if it gets a 403/400/5xx refusal, downloads the file once, keeps it in memory, and
> serves every later call from there — printing one line to stderr explaining why. Disable that with
> `EXCEL_GRAPH_FALLBACK=0`.

The one trade-off of `graph_download`: it reads the file **as last saved**, so if someone has the
workbook open in Excel with unsaved edits, the download may be stale (the workbook API would see the
live values). That's the price of working on accounts the workbook API refuses.

### 1. Pick a way to identify the file (set exactly one)

| Setting | Use it when | Example |
|---|---|---|
| `EXCEL_WORKBOOK_PATH` | You know the file's path in the drive | `/Documentos/PLANILHA.xlsx` (OneDrive personal) or `/Shared Documents/Financeiro/PLANILHA.xlsx` (SharePoint) |
| `EXCEL_DRIVE_ID` + `EXCEL_ITEM_ID` | Most precise; `npm run login` prints both | `b!AbCd...` / `01ABCD...` |
| `EXCEL_SHARE_URL` | The client sent you a share link | `https://1drv.ms/x/s!AbCd...` |

`EXCEL_SHARE_URL` works in **both** modes: in `graph` mode Microsoft resolves the link for the
signed-in account, and in `local` mode the file is downloaded and parsed offline.

### 2. Pick an auth mode

#### A. Device code with the built-in Microsoft client (fastest to try)

Leave `EXCEL_CLIENT_ID` empty. The server falls back to Microsoft's first-party
*Microsoft Graph Command Line Tools* public client, so there is **no Azure app registration** — but
each user must consent interactively, and the consent screen names Microsoft's app rather than yours.

```bash
# .env
EXCEL_SOURCE=graph
EXCEL_AUTH=device_code
EXCEL_TENANT_ID=common          # works for personal (outlook/hotmail/live) AND work accounts
EXCEL_WORKBOOK_PATH=/Documentos/PLANILHA.xlsx

npm run login                   # prints a URL + code, you sign in once, token is cached
```

#### B. Device code with your own Azure app registration (recommended for client data)

Do this if the workbook belongs to a client and you want the consent screen to name *your* app, or if
you want the token to stop working when you revoke the registration.

1. Azure Portal → **App registrations** → *New registration*
   - Name: `excel-mcp`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts**
   - Redirect URI: leave empty (device code doesn't need one)
2. **Authentication** → *Advanced settings* → **Allow public client flows = Yes** → Save
3. **API permissions** → Add → Microsoft Graph → *Delegated permissions* → add
   `Files.Read.All` and `Sites.Read.All` → *Grant admin consent* (or the user consents on first sign-in)
4. Copy the **Application (client) ID** into `.env` as `EXCEL_CLIENT_ID`, then `npm run login`.

#### C. App-only / client credentials (unattended, no user sign-in)

An Azure app registration with the **application** permission `Files.Read.All` and admin consent.
The Graph workbook API rejects app-only tokens, so this **must** be paired with
`EXCEL_SOURCE=graph_download` — the server warns you at startup if you forget.

```bash
# .env
EXCEL_SOURCE=graph_download
EXCEL_AUTH=client_secret
EXCEL_TENANT_ID=<your-tenant-id>      # a real tenant id, not "common"
EXCEL_CLIENT_ID=<client-id>
EXCEL_CLIENT_SECRET=<secret>
EXCEL_DRIVE_ID=<drive-id>
EXCEL_ITEM_ID=<item-id>
```

For SharePoint sites, an app-only app needs `Sites.Selected` (with the site granted) or a broad
`Sites.Read.All`. `client_secret` mode doesn't use `npm run login` — the token is fetched and cached
automatically.

### Notes on Graph mode

- The file must be saved as **`.xlsx`**. Legacy `.xls` and IRM-protected workbooks are rejected by the
  Graph workbook API (the server says so explicitly instead of returning a cryptic 400). The
  `graph_download`/`local` paths surface the same limit as a clear "re-save as .xlsx" message.
- `EXCEL_AUTH=device_code` requires a one-time `npm run login`. The MCP server itself never prompts —
  it can't, a stdio server has no way to show a verification code — so a missing token produces an
  actionable error telling you to run it.
- Tokens are cached at `.tokens/graph.json` (`chmod 600`) and refreshed automatically, including
  Microsoft's refresh-token rotation.
- Requests retry on 429/5xx honouring `Retry-After`.
- The download path keeps the file **in memory only** — nothing is written to disk. Download URLs are
  preauthenticated and short-lived, so they're used immediately and never persisted.
- `EXCEL_SHARE_URL` requires the `/shares` lookup, whose documented least-privileged delegated
  permission is `Files.ReadWrite`. If a share link 403s, add `Files.ReadWrite` to `EXCEL_SCOPES`, or
  switch to `EXCEL_DRIVE_ID` + `EXCEL_ITEM_ID` (which needs only `Files.Read.All`).

---

## Registering the server in your MCP client

### Cline (desktop app — this machine)

Cline reads MCP servers from:

```
~/.cline/data/settings/cline_mcp_settings.json     # overridable via CLINE_MCP_SETTINGS_PATH
```

(That path comes from the app binary itself — `resolveMcpSettingsPath()` — not from a guess.)

This file has already been created with the `excel` server:

```json
{
  "mcpServers": {
    "excel": {
      "type": "stdio",
      "command": "/Users/rafaelarantes/.local/share/fnm/aliases/default/bin/node",
      "args": ["/Volumes/Storage/www/doce-sabor/excel-mcp/dist/src/index.js"],
      "env": {},
      "disabled": false,
      "timeout": 300
    }
  }
}
```

Three deliberate choices:

- **Absolute node path** (`~/.local/share/fnm/aliases/default/bin/node`). This machine's node comes
  from `fnm`, and a GUI app does not inherit your shell's `PATH` — a bare `"command": "node"` would
  most likely fail with ENOENT. The `default` alias symlink survives node upgrades, unlike the
  version-pinned path.
- **`timeout: 300`** — seconds, not milliseconds. Normalizing a full year of messy blocks against
  Graph can exceed the default budget.
- **No `cwd`** — the server locates its own package root by walking up from its module path, so it
  doesn't care where it's launched from. Proved by `npm run cline-check`, which spawns it from `/`.

Cline also accepts a nested-transport form if your build prefers it:

```json
"excel": { "transport": { "type": "stdio", "command": "…/node", "args": ["…/index.js"] }, "timeout": 300 }
```

Verify the wiring **without touching the UI**:

```bash
npm run cline-check   # reads the real settings file, spawns it as Cline does, drives the tools
```

Then restart Cline (or toggle the server in its MCP panel) — settings are read on load.

**On approvals:** nothing is auto-approved, so Cline will ask before each call. The ten read-only
tools can be trusted to run unattended, but `excel_export_csv` writes to disk — leave that one
gated. To opt in, use Cline's MCP panel to enable auto-approve per tool, or add
`"autoApprove": ["excel_workbook_info", "excel_diagnostics", "excel_read_sheet", "excel_read_range", "excel_detect_blocks", "excel_column_profile", "excel_search", "excel_normalize_transactions", "excel_build_report", "excel_list_metrics"]`
to the entry (`autoApprove` is accepted by Cline but isn't part of the transport schema I verified).

### Claude Code

`.mcp.json` files are already in place at `/Volumes/Storage/www/doce-sabor/.mcp.json` and
`/Volumes/Storage/www/doce-sabor/excel-mcp/.mcp.json`. Claude Code picks them up when you open that
folder and asks you to approve the server once. To add it anywhere else:

```bash
claude mcp add excel --scope user -- node /Volumes/Storage/www/doce-sabor/excel-mcp/dist/src/index.js
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "node",
      "args": ["/Volumes/Storage/www/doce-sabor/excel-mcp/dist/src/index.js"],
      "env": { "EXCEL_SOURCE": "graph" }
    }
  }
}
```

`npm run build` must have been run at least once, and `npm run login` once for device-code auth.
`env` entries here override `.env`, so keep credentials in `.env` and only switch modes here.
To skip the build step entirely, use
`"command": "npx", "args": ["tsx", "/…/excel-mcp/src/index.ts"]`.

---

## Tools

| Tool | What it does |
|---|---|
| `excel_workbook_info` | File name, source, last modified, web link, and every tab with its dimensions. **Start here.** |
| `excel_diagnostics` | How the server is configured + what's broken. No network call. |
| `excel_read_sheet` | Read a tab's used range. `formato`: `records` (headers + objects), `rows` (raw grid), `markdown`, `csv`. |
| `excel_read_range` | Read one explicit A1 range (e.g. `B2:H60`). |
| `excel_column_profile` | Per-column stats: filled/blank, % that parse as dates / numbers / text, sample values. |
| `excel_detect_blocks` | Find the disconnected blocks in a messy sheet — address, title label, header row, data-row count. |
| `excel_search` | Case/accent-insensitive substring search, each hit with its A1 address and full row. |
| `excel_normalize_transactions` | Turn any tab into `data, tipo, categoria, descricao, valor, conta` rows with provenance + confidence. |
| `excel_build_report` | **The reporting workhorse.** Filter → group → aggregate (`sum/avg/count/count_distinct/min/max/median/first/last`) → markdown table + totals + share-of-total. |
| `excel_list_metrics` | Which column names, filter operators and aggregations a tab accepts, plus worked examples. |
| `excel_export_csv` | Write normalized or raw rows to a CSV (`;` separator, comma decimals, UTF-8 BOM — opens cleanly in pt-BR Excel). |

All tools are annotated `readOnlyHint: true` except `excel_export_csv`.

### Two report modes

- **`modo: "transacoes"`** (default) — normalizes the tab first, then reports over the clean
  `Data/Tipo/Categoria/Descrição/Valor/Conta` schema. Use this on the client's messy monthly tabs.
- **`modo: "bruto"`** — reports over the sheet's own columns. Use this on already-clean tabs
  (Aba 1), or when you need a column the normalizer doesn't map.

### Filter operators

`eq`, `ne`, `contains`, `icontains`, `startswith`, `gt`, `gte`, `lt`, `lte`, `in`, `not_in`, `between`,
`is_empty`, `not_empty`, `regex`, `month` (matches `YYYY-MM`). Comparisons are number- and date-aware,
so `{"column": "data", "op": "between", "value": ["2026-06-01", "2026-06-30"]}` works on normalized rows.

---

## Example prompts

> "List the tabs in the spreadsheet and tell me which one looks like the raw expense data."

> "On tab `junho26`, detect the blocks and tell me how many suppliers there are."

> "Normalize `agosto26` into the Aba 1 schema and show me the total by categoria and by conta."

> "Build a report of September spending grouped by conta, only saídas, sorted by total descending."

> "Compare julho26 vs agosto26 spending per categoria — a table with the difference."

> "Export the normalized `agosto26` rows to `/Volumes/Storage/www/doce-sabor/agosto26-normalizado.csv`."

---

## Applying it to Doce Sabor

The client's file is a hand-built Excel workbook (`PLANILHA.xlsx`) with one tab per month
(`junho26`, `julho26`, `agosto26`, `SETEMBRO`) holding ~15 disconnected supplier blocks
(MERCADO, BEBIDAS, DIVERSOS, EMBALAGENS, QUITANDA, FIXOS, …), each with its own `Data`/`Valor` pair,
and the payment method embedded in the item label.

The pipeline that maps it onto the new schema:

```
PLANILHA.xlsx (Excel Online)
   │  excel_detect_blocks          → the ~15 blocks, with labels and headers
   │  excel_column_profile         → which column is the date, which is the amount
   │  excel_normalize_transactions → Data/Tipo/Categoria/Descrição/Valor/Conta
   ▼
excel_build_report  → painel-style rollups (categoria × conta × mês)
excel_export_csv    → hand-off file for the caixa app / Aba 1 import
```

Block labels become `categoria` (MERCADO → `Mercado`, FIXOS → `Fixos`), and a `FATURAMENTO`/`RECEITA`
style label flips `tipo` to `entrada` — exactly the mapping the client's new Aba 1 expects. The `conta`
column is recovered from the `(dinheiro)` / `(pix)` annotations that made the old sheet unfilterable.
Because every row carries `origem.endereco` and `confianca`, a parse can be audited cell by cell before
anything is written anywhere.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `No cached Microsoft credentials…` | Run `npm run login` once in this folder. |
| `Graph recusou a API de workbook … baixando o arquivo` on stderr | Not an error — the account (personal, or app-only) can't use the workbook API, so the server switched to the download path. Set `EXCEL_SOURCE=graph_download` to skip the failed attempt. |
| Everything fails with 400 `invalidRequest` on `/workbook/` | The workbook API isn't available for that account/file type. Use `EXCEL_SOURCE=graph_download`. |
| 404 `Could not resolve workbook via EXCEL_WORKBOOK_PATH` | Path is drive-relative and case-sensitive: `/Documentos/…` on pt-BR OneDrive, `/Documents/…` on en-US. `npm run login` prints the working `drive`/`item` ids — put those in `.env` instead. |
| `Graph 403 accessDenied` on the `/shares` lookup | Share links document `Files.ReadWrite` as their least-privileged delegated permission. Add it to `EXCEL_SCOPES`, or use `EXCEL_DRIVE_ID` + `EXCEL_ITEM_ID`. |
| `Graph 403 accessDenied` elsewhere | The signed-in account can't read the file, or the app's permissions were never consented. |
| `re-saved as .xlsx` | The file isn't `.xlsx` (legacy `.xls`, or IRM-protected). |
| Downloaded values look stale | Someone has the workbook open with unsaved edits. `graph_download` reads the last saved version by design. |
| Share link fails in `local` mode | The link is restricted to specific people. Use `EXCEL_SOURCE=graph` + `npm run login`. |
| Only some blocks found | A gap of a fully empty row/column is what *defines* a block boundary. Raise `mesclarLinhas` to merge across spacer rows, or lower `minCelulas`. |
| Dates came out empty | That block has no date column (normal for some suppliers). The response names which blocks had none; force one with `colunaData`. |
| Amount looks wrong | Run `excel_column_profile` — if several columns are numeric, the wrong one may be picked. Force it with `colunaValor`. |
| `EXCEL_MAX_ROWS` hit | Reads cap at 5000 rows / 200k cells by default; raise them in `.env`. |

---

## Security

- **Read-only.** Only `GET` requests are ever issued against Graph.
- The token cache (`.tokens/graph.json`, `chmod 600`) holds a refresh token — treat it like a password.
  It is already in `.gitignore`.
- Graph mode honours whatever access the signed-in account has; the server never widens it.
- The download fallback holds the workbook **in memory only**; no copy is written to disk.
- Prefer your own Azure app registration (mode B) over the shared public client for client data, so you
  control revocation.
- `excel_export_csv` writes wherever you point it — keep exports of client financial data out of the repo.

---

## Scripts

| Command | What it does |
|---|---|
| `npm run doctor` | Prints the effective config, validates it, connects once and lists the tabs. |
| `npm run login` | One-time interactive device-code sign-in; caches the token and verifies the workbook is reachable. |
| `npm run smoke` | Offline self-test: fixture → blocks → normalize → report → CSV, request encoding, and the simulated workbook-API-failure → download fallback. |
| `npm run check` | Boots the compiled server over stdio, speaks MCP, drives all 11 tools end to end. |
| `npm run cline-check` | Reads the real Cline MCP settings, spawns the server exactly as Cline does (from cwd `/`), and drives the tools. |
| `npm run build` | Compiles to `dist/`. |
| `npm run dev` | Runs the server from TypeScript with `tsx`. |

## Layout

```
src/
  index.ts        entry point (stdio)
  server.ts       McpServer wiring + model instructions
  config.ts       env loading/validation (finds the package root itself)
  auth.ts         token record, cache, TokenProvider (refresh + client_credentials)
  device-code.ts  the interactive half of the device-code flow
  open.ts         picks the backend
  sources/
    graph.ts        Excel Online via the workbook API + download fallback (read-only)
    graph-api.ts    shared Graph fetch, drive-item resolution, error hints
    graph-download.ts  resolve via Graph, download the .xlsx, parse locally
    local.ts        .xlsx via exceljs (from a path or from bytes)
    download.ts     share link → raw .xlsx download
  cells.ts        cell normalization, value/text merge, used-range trimming
  types.ts        Grid / SheetInfo / WorkbookSource contracts
  tables.ts       header detection and records
  blocks.ts       block detection (flood fill) + column profiling
  parse.ts        pt-BR numbers, flexible dates
  payment.ts      payment-method extraction from labels
  normalize.ts    messy sheet → transaction rows
  report-filters.ts  filters, aggregations, formatting
  report.ts       buildReport, markdown, CSV
  tools/          the 11 MCP tools
scripts/          doctor, login, smoke, mcp-check, fixture
```
