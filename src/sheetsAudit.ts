// Self-auditing wrapper for the Sheets client — every mutation this server
// performs on any spreadsheet appends one row to a "Claude MCP Log" tab in
// that same spreadsheet (created on first write). Enforced HERE, at the only
// client access point, so no tool — present or future — can write unlogged.
//
// Why: on 26 Aug 2026 a sheets session deleted rows from the bill-tracker
// Transactions tab and Google's own history attributed it to the human
// account the OAuth token belongs to, making "a Claude session" and "the
// user by hand" indistinguishable. The log tab is the differentiator: an
// edit with a matching MCP log row was Claude; without one it was the user
// (or an add-on automation like Tiller acting under their grant).
//
// Rules:
// - Audit rows are appended AFTER the mutation succeeds; a failed mutation
//   logs nothing.
// - Audit failures never fail the user's edit (console.error only).
// - Writes targeting the log tab itself are not audited (no recursion).
// - deleteDimension requests record the exact row/column indexes — the
//   forensic detail that was missing from the original incident.

import type { sheets_v4 } from 'googleapis';

export const AUDIT_LOG_TAB = 'Claude MCP Log';
const HEADER = ['Timestamp (UTC)', 'Action', 'Detail', 'Source'];

const wrappedClients = new WeakSet<object>();

function isLogRange(range: string | null | undefined): boolean {
  return !!range && range.includes(AUDIT_LOG_TAB);
}

// Human summary of a spreadsheets.batchUpdate request list, with exact
// indexes for destructive requests.
function summarizeRequests(requests: sheets_v4.Schema$Request[] | undefined): string {
  if (!requests?.length) return '';
  const parts: string[] = [];
  const counts = new Map<string, number>();
  for (const req of requests) {
    const kind = Object.keys(req)[0] || 'unknown';
    if (kind === 'deleteDimension') {
      const r = req.deleteDimension?.range;
      parts.push(
        `deleteDimension ${r?.dimension || '?'} ${(r?.startIndex ?? 0) + 1}-${r?.endIndex ?? '?'} (sheetId ${r?.sheetId ?? '?'})`,
      );
    } else if (kind === 'deleteRange') {
      const r = req.deleteRange?.range;
      parts.push(
        `deleteRange rows ${(r?.startRowIndex ?? 0) + 1}-${r?.endRowIndex ?? '?'} cols ${(r?.startColumnIndex ?? 0) + 1}-${r?.endColumnIndex ?? '?'} (sheetId ${r?.sheetId ?? '?'})`,
      );
    } else if (kind === 'deleteSheet') {
      parts.push(`deleteSheet sheetId ${req.deleteSheet?.sheetId ?? '?'}`);
    } else {
      counts.set(kind, (counts.get(kind) || 0) + 1);
    }
  }
  for (const [kind, n] of counts) parts.push(n > 1 ? `${kind} x${n}` : kind);
  return parts.join(', ');
}

// True when a batchUpdate is (only) the audit's own log-tab creation.
function isAuditTabCreation(requests: sheets_v4.Schema$Request[] | undefined): boolean {
  return !!requests?.every((r) => r.addSheet?.properties?.title === AUDIT_LOG_TAB);
}

export function wrapSheetsClientWithAudit(sheets: sheets_v4.Sheets): sheets_v4.Sheets {
  if (wrappedClients.has(sheets)) return sheets;
  wrappedClients.add(sheets);

  const values = sheets.spreadsheets.values;
  const ss = sheets.spreadsheets;

  const origValuesUpdate = values.update.bind(values);
  const origValuesAppend = values.append.bind(values);
  const origValuesBatchUpdate = values.batchUpdate.bind(values);
  const origValuesClear = values.clear.bind(values);
  const origValuesBatchClear = values.batchClear.bind(values);
  const origBatchUpdate = ss.batchUpdate.bind(ss);

  async function audit(spreadsheetId: string | undefined, action: string, detail: string): Promise<void> {
    if (!spreadsheetId) return;
    const row = [new Date().toISOString(), action, detail, 'claude-mcp'];
    try {
      await origValuesAppend({
        spreadsheetId,
        range: `'${AUDIT_LOG_TAB}'!A1`,
        valueInputOption: 'RAW',
        insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] },
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (/unable to parse range|not found/i.test(msg)) {
        // Log tab doesn't exist yet in this spreadsheet — create it with a
        // header, then log. One retry only.
        try {
          await origBatchUpdate({
            spreadsheetId,
            requestBody: { requests: [{ addSheet: { properties: { title: AUDIT_LOG_TAB } } }] },
          });
          await origValuesAppend({
            spreadsheetId,
            range: `'${AUDIT_LOG_TAB}'!A1`,
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            requestBody: { values: [HEADER, row] },
          });
        } catch (err2) {
          console.error(`[sheets-audit] could not create/append log tab in ${spreadsheetId}: ${err2}`);
        }
      } else {
        console.error(`[sheets-audit] append failed for ${spreadsheetId}: ${msg}`);
      }
    }
  }

  values.update = (async (params: sheets_v4.Params$Resource$Spreadsheets$Values$Update, ...rest: unknown[]) => {
    const res = await (origValuesUpdate as Function)(params, ...rest);
    if (!isLogRange(params?.range)) await audit(params?.spreadsheetId, 'values.update', `range ${params?.range}`);
    return res;
  }) as typeof values.update;

  values.append = (async (params: sheets_v4.Params$Resource$Spreadsheets$Values$Append, ...rest: unknown[]) => {
    const res = await (origValuesAppend as Function)(params, ...rest);
    if (!isLogRange(params?.range)) {
      const n = params?.requestBody?.values?.length ?? 0;
      await audit(params?.spreadsheetId, 'values.append', `${n} row(s) at ${params?.range}`);
    }
    return res;
  }) as typeof values.append;

  values.batchUpdate = (async (params: sheets_v4.Params$Resource$Spreadsheets$Values$Batchupdate, ...rest: unknown[]) => {
    const res = await (origValuesBatchUpdate as Function)(params, ...rest);
    const ranges = (params?.requestBody?.data || []).map((d) => d.range).filter((r) => !isLogRange(r));
    if (ranges.length) await audit(params?.spreadsheetId, 'values.batchUpdate', `ranges ${ranges.join('; ')}`);
    return res;
  }) as typeof values.batchUpdate;

  values.clear = (async (params: sheets_v4.Params$Resource$Spreadsheets$Values$Clear, ...rest: unknown[]) => {
    const res = await (origValuesClear as Function)(params, ...rest);
    if (!isLogRange(params?.range)) await audit(params?.spreadsheetId, 'values.clear', `range ${params?.range}`);
    return res;
  }) as typeof values.clear;

  values.batchClear = (async (params: sheets_v4.Params$Resource$Spreadsheets$Values$Batchclear, ...rest: unknown[]) => {
    const res = await (origValuesBatchClear as Function)(params, ...rest);
    const ranges = (params?.requestBody?.ranges || []).filter((r) => !isLogRange(r));
    if (ranges.length) await audit(params?.spreadsheetId, 'values.batchClear', `ranges ${ranges.join('; ')}`);
    return res;
  }) as typeof values.batchClear;

  ss.batchUpdate = (async (params: sheets_v4.Params$Resource$Spreadsheets$Batchupdate, ...rest: unknown[]) => {
    const res = await (origBatchUpdate as Function)(params, ...rest);
    const requests = params?.requestBody?.requests ?? undefined;
    if (!isAuditTabCreation(requests)) {
      const summary = summarizeRequests(requests);
      if (summary) await audit(params?.spreadsheetId, 'batchUpdate', summary);
    }
    return res;
  }) as typeof ss.batchUpdate;

  return sheets;
}
