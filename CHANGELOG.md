# Changelog

## v1.1.0 — 2026-03-04

### Fixed
- Fix orphaned process 100% CPU spin: `@modelcontextprotocol/sdk` `StdioServerTransport` doesn't listen for stdin `end`/`close` events, causing libuv to spin-poll a dead fd when the parent dies. Added stdin event handlers for immediate exit.
- Parent PID watchdog kept as fallback for macOS where stdin events are unreliable.

## v1.0.1 — 2026-02-28

### Fixed
- Replace unreliable stdin `end`/`close` handlers with parent PID watchdog polling.

### Added
- `setDropdownValidation` now supports `sourceRange` parameter for `ONE_OF_RANGE` validation.

## v1.0.0 — 2026-02-25

### Added
- Fork of [a-bonus/google-docs-mcp](https://github.com/a-bonus/google-docs-mcp) with custom extensions
- `getBoundScriptId` tool for finding bound Apps Script projects
- `getCellFormatting` tool for reading cell formatting details
- Removed upstream CI/release workflows
- Parent PID watchdog for orphan process cleanup
