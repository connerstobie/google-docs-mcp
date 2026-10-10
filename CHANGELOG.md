# Changelog

## 1.9.1

### Security
- **27 dependency advisories patched, all inside existing ranges (BILL-176).** Found by an audit of every repo on 9 Oct 2026. Critical: protobufjs (arbitrary code execution, now 7.6.6, through `@google-cloud/firestore` and `@grpc/grpc-js`) and proxy-addr (IP spoofing, now 2.0.8). High: `@modelcontextprotocol/sdk` (now 1.32.1), `@grpc/grpc-js`, axios, undici, hono, form-data, path-to-regexp, fast-uri, ip-address, linkify-it, nanoid, postcss, source-map-js, vite and vitest (dev), brace-expansion. No direct dependency changed. `npm audit`: 0 vulnerabilities; 177 tests pass; the server starts with its 84 tools and a live `listDriveFiles` call answers.
- npm 10.9.4 crashed in `npm audit fix` on this lockfile ("Cannot read properties of null (reading 'edgesOut')"); the fix was resolved with npm 11 and reconciled with npm 10 so `npm ci` installs it.

## 1.9.0

### Added
- **Self-auditing sheet writes.** Every mutation to a spreadsheet (`values.update/append/batchUpdate/clear/batchClear` and structural `batchUpdate`, which includes row deletions) now appends a row — timestamp, action, detail, `claude-mcp` — to a "Claude MCP Log" tab in the target spreadsheet, created automatically on first write. Enforced by wrapping the Sheets client at its single access point (`getSheetsClient`, both local and remote request paths), so no tool present or future can write unlogged. `deleteDimension`/`deleteRange` entries record exact row/column indexes. Audit rows land only after the mutation succeeds, audit failures never fail the user's edit, and writes to the log tab itself are excluded to prevent recursion. Born from the 26 Aug 2026 bill-tracker incident: rows deleted by a sheets session were indistinguishable from the account owner's own edits, because MCP OAuth acts as the owner in every record Google keeps.

## 1.8.0

### Added
- `listCharts` — lists every chart with its numeric chart ID and full spec: title, type, stacking, legend position, axis titles, domain range, and each series range in stacking order with its explicit color and target axis, plus anchor cell and pixel size. Charts are embedded objects rather than cell values, so `readSpreadsheet` returns nothing for a sheet holding only charts. Without this there was no way to tell an empty tab from one full of charts, and no way to obtain the chart ID that `deleteChart` requires.
- `updateChart` — edits an existing chart in place, keeping its ID and every property not passed. Covers title, subtitle, chart type, stacking, legend position, axis titles, domain range, the full series list including stacking order, per-series color and name, header count, and the chart's anchor cell and pixel size. It reads the current spec and patches it, because the underlying `updateChartSpec` request replaces the spec wholesale.

### Google API constraints found while testing this against the live API
- The Sheets API has no series-name field. A series is named by the first cell of its source range when `headerCount` is 1, so `updateChart` takes a `nameCell` and merges it into the range when it sits immediately before the values. A detached label has to be passed as a second source, which Google accepts only for column-oriented data listed in ascending sheet order, and rejects for row-oriented series with `ChartSourceRange ranges require all rows or all columns to have length of 1`. Both failure modes are now caught locally with an explanatory error instead of surfacing the API's generic contiguity complaint.
- `updateEmbeddedObjectPosition` rejects sub-field masks (`Invalid field: overlay_position`), so the position is replaced whole, merged from the existing overlay.
- The value axis number format is not exposed by the API. It is inferred from the source cells, so a single stray date-formatted cell anywhere in a chart's range renders the whole axis as dates.

## 1.1.0

### Fixed
- Fix orphaned process 100% CPU spin: `@modelcontextprotocol/sdk` `StdioServerTransport` doesn't listen for stdin `end`/`close` events, causing libuv to spin-poll a dead fd when the parent dies. Added stdin event handlers for immediate exit.
- Parent PID watchdog kept as fallback for macOS where stdin events are unreliable.

## 1.0.1

### Fixed
- Replace unreliable stdin `end`/`close` handlers with parent PID watchdog polling.

### Added
- `setDropdownValidation` now supports `sourceRange` parameter for `ONE_OF_RANGE` validation.

## 1.0.0

### Added
- Fork of [a-bonus/google-docs-mcp](https://github.com/a-bonus/google-docs-mcp) with custom extensions
- `getBoundScriptId` tool for finding bound Apps Script projects
- `getCellFormatting` tool for reading cell formatting details
- Removed upstream CI/release workflows
- Parent PID watchdog for orphan process cleanup
