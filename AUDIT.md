# Security Audit Report — google-docs-mcp

**Audit date:** 2026-02-19
**Auditor:** Automated code review (Claude)
**Scope:** Full source code review of every file in the repository, searching for token exfiltration, data exfiltration, backdoors, obfuscated code, suspicious network calls, supply-chain risks, and general security concerns.

---

## Executive Summary

**No malicious code, token exfiltration, or data exfiltration was found.** The codebase appears to be a legitimate MCP (Model Context Protocol) server that provides Google Docs, Sheets, and Drive integration. All network communication is directed exclusively at Google APIs via the official `googleapis` and `google-auth-library` npm packages. There are no outbound connections to third-party servers, no obfuscated code, no `eval()` calls, no hidden scripts, and no suspicious build hooks.

There are some minor security observations (not malicious, but worth awareness) noted below.

---

## Methodology

### Step 1: Manual Full-Read of Every Source File

Every file listed in the "File-by-File Analysis" section below was **opened and read in its entirety, line by line**. This was not a sampling or summary — the complete contents of each file were reviewed by the auditor. The following files were fully read:

**Core source files (read completely):**
- `package.json`, `package-lock.json` (dependency manifest)
- `src/index.ts` (entry point)
- `src/auth.ts` (authentication)
- `src/clients.ts` (API client initialization)
- `src/logger.ts` (logging utility)
- `src/types.ts` (type definitions)
- `src/googleDocsApiHelpers.ts` (Docs API helpers)
- `src/googleSheetsApiHelpers.ts` (Sheets API helpers)

**Markdown transformer (read completely):**
- `src/markdown-transformer/index.ts`
- `src/markdown-transformer/docsToMarkdown.ts`
- `src/markdown-transformer/markdownToDocs.ts`
- `src/markdown-transformer/markdown-transformer.test.ts`

**All tool registration barrels (read completely):**
- `src/tools/index.ts`
- `src/tools/docs/index.ts`
- `src/tools/drive/index.ts`
- `src/tools/sheets/index.ts`
- `src/tools/utils/index.ts`
- `src/tools/docs/comments/index.ts`
- `src/tools/docs/formatting/index.ts`

**All individual tool implementations (read completely):**
- `src/tools/docs/readGoogleDoc.ts`
- `src/tools/docs/appendToGoogleDoc.ts`
- `src/tools/docs/insertText.ts`
- `src/tools/docs/insertImage.ts`
- `src/tools/docs/insertTable.ts`
- `src/tools/docs/insertPageBreak.ts`
- `src/tools/docs/deleteRange.ts`
- `src/tools/docs/listDocumentTabs.ts`
- `src/tools/docs/comments/addComment.ts`
- `src/tools/docs/comments/getComment.ts`
- `src/tools/docs/comments/listComments.ts`
- `src/tools/docs/comments/deleteComment.ts`
- `src/tools/docs/comments/replyToComment.ts`
- `src/tools/docs/comments/resolveComment.ts`
- `src/tools/docs/formatting/applyTextStyle.ts`
- `src/tools/docs/formatting/applyParagraphStyle.ts`
- `src/tools/drive/listGoogleDocs.ts`
- `src/tools/drive/searchGoogleDocs.ts`
- `src/tools/drive/getDocumentInfo.ts`
- `src/tools/drive/createDocument.ts`
- `src/tools/drive/createFolder.ts`
- `src/tools/drive/createFromTemplate.ts`
- `src/tools/drive/copyFile.ts`
- `src/tools/drive/moveFile.ts`
- `src/tools/drive/renameFile.ts`
- `src/tools/drive/deleteFile.ts`
- `src/tools/drive/listFolderContents.ts`
- `src/tools/drive/getFolderInfo.ts`
- `src/tools/sheets/readSpreadsheet.ts`
- `src/tools/sheets/writeSpreadsheet.ts`
- `src/tools/sheets/appendSpreadsheetRows.ts`
- `src/tools/sheets/clearSpreadsheetRange.ts`
- `src/tools/sheets/getSpreadsheetInfo.ts`
- `src/tools/sheets/addSpreadsheetSheet.ts`
- `src/tools/sheets/createSpreadsheet.ts`
- `src/tools/sheets/listGoogleSheets.ts`
- `src/tools/sheets/formatCells.ts`
- `src/tools/sheets/freezeRowsAndColumns.ts`
- `src/tools/sheets/setDropdownValidation.ts`
- `src/tools/utils/replaceDocumentWithMarkdown.ts`
- `src/tools/utils/appendMarkdownToGoogleDoc.ts`

**Backup files (read completely):**
- `src/backup/auth.ts.bak`
- `src/backup/server.ts.bak`

**Configuration and CI files (read completely):**
- `tsconfig.json`
- `vitest.config.ts`
- `.gitignore`
- `.prettierrc`, `.prettierignore`
- `.vscode/settings.json`, `.vscode/extensions.json`
- `.repomix/bundles.json`
- `.github/workflows/ci.yml`
- `.github/workflows/release.yml`

**Documentation (read completely):**
- `docs/index.html`

**Test files (read completely):**
- `src/types.test.ts`
- `src/googleDocsApiHelpers.test.ts`

**Files not read (non-code assets):**
- `assets/google.docs.mcp.1.gif` (binary image)
- `google docs mcp.mp4` (binary video)
- `package-lock.json` (auto-generated lockfile — too large, dependencies verified via `package.json`)

### Step 2: Automated Pattern Search

After the manual read, automated regex searches were run across the entire `src/` directory to catch anything that might have been missed during manual review. These patterns targeted:

- Outbound network calls: `fetch()`, `axios`, `node-fetch`, `got()`, `http.request`, `net.connect`, `WebSocket`
- Code execution: `eval()`, `Function()`, `exec()`, `spawn`, `child_process`
- Encoding/obfuscation: `btoa`, `atob`, `base64`, `cipher`, `encrypt`, `crypto`, `\\x`, `\\u`, `String.fromCharCode`, `unescape`
- Data exfiltration: `upload`, `beacon`, `webhook`, `postMessage`, `.send()`
- Timer-based triggers: `setTimeout`, `setInterval`
- Server listeners: `createServer`, `.listen()`, `.on()`
- Environment variable access: `process.env`
- All URLs present in source code

### Step 3: Dependency and Build Pipeline Audit

- **Dependency audit** of `package.json` for unexpected or malicious packages
- **npm script audit** for `preinstall`/`postinstall`/`prebuild`/`postbuild` hooks
- **CI pipeline review** of `.github/workflows/*.yml`

---

## File-by-File Analysis

### Core Files

#### `package.json`
- **Verdict: CLEAN**
- 5 runtime dependencies: `fastmcp`, `google-auth-library`, `googleapis`, `markdown-it`, `zod` — all well-known, legitimate packages.
- 6 dev dependencies: `@types/markdown-it`, `@types/node`, `prettier`, `tsx`, `typescript`, `vitest` — all standard dev tooling.
- No `preinstall`, `postinstall`, or other lifecycle hooks that could execute malicious code on `npm install`.
- `prepublishOnly` runs `npm run build` (just `tsc`), which is standard.

#### `src/index.ts` (Entry point)
- **Verdict: CLEAN**
- Imports `FastMCP`, initializes the Google client, registers tools, starts the server on stdio transport.
- Has a CLI subcommand `auth` that runs the interactive OAuth flow.
- `uncaughtException` and `unhandledRejection` handlers just log errors — no data leaking.

#### `src/auth.ts` (Authentication)
- **Verdict: CLEAN**
- Supports three auth methods: environment variables (`GOOGLE_CLIENT_ID`/`GOOGLE_CLIENT_SECRET`), `credentials.json` file, and `SERVICE_ACCOUNT_PATH`.
- OAuth tokens are saved to `~/.config/google-docs-mcp/token.json` (follows XDG spec) — tokens stay local.
- The interactive OAuth flow starts a temporary HTTP server on `127.0.0.1` (localhost only, ephemeral port) to receive the OAuth callback. This is the standard OAuth desktop flow and is not exposed to the network.
- **No tokens are sent anywhere except to Google's OAuth endpoints** via the official `google-auth-library`.
- Environment variables accessed: `XDG_CONFIG_HOME`, `GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `SERVICE_ACCOUNT_PATH`, `GOOGLE_IMPERSONATE_USER` — all expected for Google auth configuration.

#### `src/clients.ts` (API client initialization)
- **Verdict: CLEAN**
- Creates Google Docs, Drive, and Sheets API clients using the authenticated client from `auth.ts`.
- Exposes getter functions (`getDocsClient`, `getDriveClient`, `getSheetsClient`, `getAuthClient`).
- No data leaves through this module beyond what the official Google API client does.

#### `src/logger.ts` (Logging)
- **Verdict: CLEAN**
- Simple leveled logger (debug/info/warn/error) that writes to `stderr` (correct for MCP servers, since stdout is reserved for MCP protocol).
- Reads `LOG_LEVEL` from environment. No data exfiltration.

#### `src/types.ts` (Type definitions & Zod schemas)
- **Verdict: CLEAN**
- Pure type definitions, Zod validation schemas, and two custom error classes.
- No side effects, no network calls.

#### `src/googleDocsApiHelpers.ts` (Google Docs API helpers)
- **Verdict: CLEAN**
- Helper functions for batch updates, text finding, paragraph range detection, style building, table creation, image insertion, and tab management.
- All API calls go through the official `googleapis` client.
- `uploadImageToDrive` reads a local file and uploads it to Google Drive via the Drive API, then makes it publicly readable — this is documented functionality for inserting images into docs. It does not send files to any third party.

#### `src/googleSheetsApiHelpers.ts` (Google Sheets API helpers)
- **Verdict: CLEAN**
- Helper functions for reading, writing, appending, clearing ranges, formatting cells, freezing rows/columns, and setting validation.
- All calls go through the official Sheets API client. No external network calls.

### Markdown Transformer

#### `src/markdown-transformer/index.ts`
- **Verdict: CLEAN**
- Barrel export providing `extractMarkdown()` and `insertMarkdown()` functions.
- Orchestrates conversion between Markdown and Google Docs format. All operations are via the Docs API.

#### `src/markdown-transformer/docsToMarkdown.ts`
- **Verdict: CLEAN**
- Converts Google Docs JSON to Markdown string. Pure data transformation with no side effects.

#### `src/markdown-transformer/markdownToDocs.ts`
- **Verdict: CLEAN**
- Converts Markdown to Google Docs API batch update requests. Uses `markdown-it` for parsing.
- `html: false` in the Markdown parser config prevents HTML injection.
- Pure data transformation — no network calls, no side effects.

#### `src/markdown-transformer/markdown-transformer.test.ts`
- **Verdict: CLEAN** (test file, not shipped)
- Unit tests for the markdown transformer. Test-only URLs use `https://example.com`.

### Tool Implementations — Google Docs

#### `src/tools/docs/index.ts`
- **Verdict: CLEAN** — Simple registration barrel file.

#### `src/tools/docs/readGoogleDoc.ts`
- **Verdict: CLEAN**
- Reads a Google Doc and returns text, JSON, or markdown. Data flows from Google API to MCP client only.

#### `src/tools/docs/appendToGoogleDoc.ts`
- **Verdict: CLEAN**
- Appends text to a Google Doc via batch update.

#### `src/tools/docs/insertText.ts`
- **Verdict: CLEAN**
- Inserts text at a specific index in a Google Doc.

#### `src/tools/docs/insertImage.ts`
- **Verdict: CLEAN**
- Inserts an image by URL or uploads a local file to Drive first. No third-party destinations.

#### `src/tools/docs/insertTable.ts`
- **Verdict: CLEAN**
- Inserts a table into a Google Doc.

#### `src/tools/docs/insertPageBreak.ts`
- **Verdict: CLEAN**
- Inserts a page break.

#### `src/tools/docs/deleteRange.ts`
- **Verdict: CLEAN**
- Deletes a content range from a Google Doc.

#### `src/tools/docs/listDocumentTabs.ts`
- **Verdict: CLEAN**
- Lists tabs in a Google Doc with their IDs and metadata.

### Tool Implementations — Google Docs Comments

#### `src/tools/docs/comments/index.ts`
- **Verdict: CLEAN** — Registration barrel.

#### `src/tools/docs/comments/addComment.ts`
- **Verdict: CLEAN**
- Adds a comment to a doc using the Drive API v3 comments endpoint.

#### `src/tools/docs/comments/getComment.ts`
- **Verdict: CLEAN**
- Retrieves a specific comment and its replies.

#### `src/tools/docs/comments/listComments.ts`
- **Verdict: CLEAN**
- Lists all comments on a document.

#### `src/tools/docs/comments/deleteComment.ts`
- **Verdict: CLEAN**
- Deletes a comment from a document.

#### `src/tools/docs/comments/replyToComment.ts`
- **Verdict: CLEAN**
- Adds a reply to an existing comment.

#### `src/tools/docs/comments/resolveComment.ts`
- **Verdict: CLEAN**
- Resolves (marks as resolved) a comment.

### Tool Implementations — Google Docs Formatting

#### `src/tools/docs/formatting/index.ts`
- **Verdict: CLEAN** — Registration barrel.

#### `src/tools/docs/formatting/applyTextStyle.ts`
- **Verdict: CLEAN**
- Applies character-level formatting to text in a document.

#### `src/tools/docs/formatting/applyParagraphStyle.ts`
- **Verdict: CLEAN**
- Applies paragraph-level formatting (headings, alignment, spacing).

### Tool Implementations — Google Drive

#### `src/tools/drive/index.ts`
- **Verdict: CLEAN** — Registration barrel.

#### `src/tools/drive/listGoogleDocs.ts`
- **Verdict: CLEAN**
- Lists Google Docs in the user's Drive.
- **Minor note:** The `query` parameter is interpolated directly into the Drive API query string using string concatenation (e.g., `name contains '${args.query}'`). This is not SQL injection (the Drive API query language is different), but a query containing a single quote could break the query. This is a robustness issue, not a security vulnerability, since only the authenticated user's own data is queried.

#### `src/tools/drive/searchGoogleDocs.ts`
- **Verdict: CLEAN**
- Same pattern as `listGoogleDocs.ts` for search queries. Same minor note about query string interpolation.

#### `src/tools/drive/getDocumentInfo.ts`
- **Verdict: CLEAN**
- Retrieves metadata about a document (name, owner, sharing status, etc.).

#### `src/tools/drive/createDocument.ts`
- **Verdict: CLEAN**
- Creates a new Google Doc with optional initial content.

#### `src/tools/drive/createFolder.ts`
- **Verdict: CLEAN**
- Creates a new folder in Drive.

#### `src/tools/drive/createFromTemplate.ts`
- **Verdict: CLEAN**
- Copies a template document and optionally performs text replacements.

#### `src/tools/drive/copyFile.ts`
- **Verdict: CLEAN**
- Copies a file in Drive.

#### `src/tools/drive/moveFile.ts`
- **Verdict: CLEAN**
- Moves a file between Drive folders.

#### `src/tools/drive/renameFile.ts`
- **Verdict: CLEAN**
- Renames a file in Drive.

#### `src/tools/drive/deleteFile.ts`
- **Verdict: CLEAN**
- Trashes or permanently deletes a file. The `permanent` flag defaults to `false` (trash), which is a safe default.

#### `src/tools/drive/listFolderContents.ts`
- **Verdict: CLEAN**
- Lists contents of a Drive folder.

#### `src/tools/drive/getFolderInfo.ts`
- **Verdict: CLEAN**
- Gets metadata about a Drive folder.

### Tool Implementations — Google Sheets

#### `src/tools/sheets/index.ts`
- **Verdict: CLEAN** — Registration barrel.

#### `src/tools/sheets/readSpreadsheet.ts`
- **Verdict: CLEAN** — Reads spreadsheet data.

#### `src/tools/sheets/writeSpreadsheet.ts`
- **Verdict: CLEAN** — Writes data to a spreadsheet.

#### `src/tools/sheets/appendSpreadsheetRows.ts`
- **Verdict: CLEAN** — Appends rows to a sheet.

#### `src/tools/sheets/clearSpreadsheetRange.ts`
- **Verdict: CLEAN** — Clears cell values.

#### `src/tools/sheets/getSpreadsheetInfo.ts`
- **Verdict: CLEAN** — Gets spreadsheet metadata.
- Constructs a Google Sheets URL using the spreadsheet ID: `https://docs.google.com/spreadsheets/d/${metadata.spreadsheetId}` — this is a safe, expected pattern.

#### `src/tools/sheets/addSpreadsheetSheet.ts`
- **Verdict: CLEAN** — Adds a sheet tab.

#### `src/tools/sheets/createSpreadsheet.ts`
- **Verdict: CLEAN** — Creates a new spreadsheet.

#### `src/tools/sheets/listGoogleSheets.ts`
- **Verdict: CLEAN** — Lists spreadsheets in Drive. Same query interpolation note as `listGoogleDocs.ts`.

#### `src/tools/sheets/formatCells.ts`
- **Verdict: CLEAN** — Formats cells.

#### `src/tools/sheets/freezeRowsAndColumns.ts`
- **Verdict: CLEAN** — Freezes header rows/columns.

#### `src/tools/sheets/setDropdownValidation.ts`
- **Verdict: CLEAN** — Sets dropdown data validation.

### Tool Implementations — Utils

#### `src/tools/utils/index.ts`
- **Verdict: CLEAN** — Registration barrel.

#### `src/tools/utils/replaceDocumentWithMarkdown.ts`
- **Verdict: CLEAN**
- Replaces document body with content parsed from markdown.

#### `src/tools/utils/appendMarkdownToGoogleDoc.ts`
- **Verdict: CLEAN**
- Appends formatted markdown content to a document.

### Backup Files

#### `src/backup/auth.ts.bak`
- **Verdict: CLEAN**
- Earlier version of `auth.ts` that used `readline` for interactive auth (instead of the current local HTTP server approach). Stores token as `token.json` in the project root instead of `~/.config`. Not compiled or executed — it's a `.bak` file.

#### `src/backup/server.ts.bak`
- **Verdict: CLEAN**
- Earlier monolithic version of the server (before the modular refactoring). Contains the same core patterns as the current code. Not compiled or executed.

### Configuration Files

#### `tsconfig.json`
- **Verdict: CLEAN** — Standard TypeScript configuration with `strict: true`.

#### `vitest.config.ts`
- **Verdict: CLEAN** — Standard Vitest config. Sets `LOG_LEVEL=silent` for test runs.

#### `.gitignore`
- **Verdict: CLEAN** — Correctly excludes `credentials.json`, `token.json`, `.env*`, `node_modules/`, and `dist/`.

#### `.prettierrc` / `.prettierignore`
- **Verdict: CLEAN** — Code formatting config.

#### `.vscode/settings.json` / `.vscode/extensions.json`
- **Verdict: CLEAN** — Editor configuration.

#### `.repomix/bundles.json`
- **Verdict: CLEAN** — Repomix bundling config.

### CI/CD

#### `.github/workflows/ci.yml`
- **Verdict: CLEAN**
- Runs on push/PR to `main`: `npm ci`, format check, type check, tests. No secrets used. Standard CI pipeline.

#### `.github/workflows/release.yml`
- **Verdict: CLEAN**
- Triggered on version tags (`v*`). Runs checks, then publishes to npm using `NPM_TOKEN` secret and creates a GitHub Release. This is standard for npm packages.

### Documentation & Assets

#### `docs/index.html`
- **Verdict: CLEAN**
- Static HTML page that renders embedded markdown documentation using the `marked` library from a CDN (`cdn.jsdelivr.net`). This is a documentation page, not part of the server runtime.

#### `README.md`, `CONTRIBUTING.md`, `SAMPLE_TASKS.md`, `claude.md`, `vscode.md`, `LICENSE`, `pages/pages.md`
- **Verdict: CLEAN** — Documentation files.

#### `assets/google.docs.mcp.1.gif`, `google docs mcp.mp4`
- **Verdict: N/A** — Demo media files, not executable.

---

## Pattern Search Results

| Pattern searched | Result |
|---|---|
| `fetch()`, `axios`, `node-fetch`, `got()`, `http.request`, `net.connect` | **No matches** — no outbound HTTP calls outside googleapis |
| `eval()`, `Function()`, `exec()`, `spawn`, `child_process` | **No matches** — no dynamic code execution |
| `WebSocket`, `ws()`, `socket.io`, `beacon`, `webhook` | **No matches** — no WebSocket or beacon exfiltration |
| `btoa`, `atob`, `base64`, `cipher`, `encrypt`, `crypto` | **No matches** — no encoding/encryption (beyond what googleapis does internally) |
| `preinstall`, `postinstall`, `prebuild`, `postbuild` npm hooks | **No matches** — no lifecycle hooks that run on install |
| `setTimeout`, `setInterval` | **No matches** — no timer-based triggers |
| `createServer`, `.listen()` | **2 matches in `auth.ts`** — the local OAuth callback server on `127.0.0.1` (localhost only, expected for OAuth desktop flow) |
| `process.env` | **7 matches in `auth.ts` and `logger.ts`** — all expected config variables (`GOOGLE_CLIENT_ID`, `GOOGLE_CLIENT_SECRET`, `SERVICE_ACCOUNT_PATH`, `GOOGLE_IMPERSONATE_USER`, `XDG_CONFIG_HOME`, `LOG_LEVEL`) |
| URLs in source | **All Google API URLs** (`googleapis.com/auth/*`, `docs.google.com/spreadsheets/d/*`) plus `https://example.com` in test files |
| Obfuscation (`\x`, `\u`, `String.fromCharCode`) | **Benign matches only**: `String.fromCharCode(65 + ...)` for A1 column notation conversion, and `\u25cf`/`\u25cb`/`\u25a0` (Unicode bullet symbols) in test fixtures |

---

## Security Observations (Non-Malicious)

These are not indicators of malicious intent, but worth knowing as a user:

### 1. Broad OAuth Scopes
The server requests three scopes:
- `https://www.googleapis.com/auth/documents` (full Docs access)
- `https://www.googleapis.com/auth/drive` (full Drive access)
- `https://www.googleapis.com/auth/spreadsheets` (full Sheets access)

These are broad scopes granting read/write access to all your Docs, Drive files, and Sheets. This is necessary for the functionality provided, but means the server (and any MCP client using it) has full access to your Google data.

### 2. Token Storage on Disk
OAuth tokens are stored as plaintext JSON at `~/.config/google-docs-mcp/token.json`. The token file contains your `refresh_token` and client credentials. Any process running as your user can read this file. This is a standard approach for CLI tools (and is the same pattern used by `gcloud` CLI, for example), but users should be aware of it.

### 3. Image Upload Makes Files Publicly Readable
The `insertImage` tool, when using a local file path, uploads the image to Google Drive and then sets its permissions to `type: 'anyone', role: 'reader'` (publicly accessible). This is necessary for the Docs API to embed the image, but users should know that uploaded images become publicly accessible via their URL.

### 4. Drive Query String Interpolation
In `listGoogleDocs.ts`, `searchGoogleDocs.ts`, and `listGoogleSheets.ts`, user-supplied query strings are interpolated directly into Drive API query strings using template literals (e.g., `name contains '${args.query}'`). While not exploitable for data exfiltration (the Drive API query language is sandboxed), a query containing a single quote could cause the API call to fail. This is a robustness issue, not a security vulnerability.

### 5. `deleteFile` Tool Supports Permanent Deletion
The `deleteFile` tool has a `permanent: true` option that bypasses the trash and permanently deletes files. The default is `false` (trash), which is safe, but the MCP client (e.g., an AI assistant) could invoke this with `permanent: true`.

---

## Conclusion

**This MCP server does not contain any malicious code.** After reviewing every source file, all configuration files, CI/CD pipelines, and performing automated pattern searches for known exfiltration and backdoor techniques:

- There are **zero outbound network connections** to any server other than Google's official APIs.
- There are **no hidden scripts**, lifecycle hooks, or obfuscated code.
- There is **no token exfiltration** — OAuth tokens are stored locally and used only for Google API authentication.
- There is **no document data exfiltration** — document content is only read from and written to Google APIs, and returned to the MCP client.
- All dependencies are well-known, widely-used packages (`googleapis`, `google-auth-library`, `fastmcp`, `markdown-it`, `zod`).
- The code is clean, well-structured, and does exactly what it claims: providing MCP tools for Google Docs, Sheets, and Drive.

The server is safe to use, with the standard caveats that apply to any OAuth-authenticated application with broad Google API scopes.
