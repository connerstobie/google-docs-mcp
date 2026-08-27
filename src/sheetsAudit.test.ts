import { describe, it, expect, vi } from 'vitest';
import { wrapSheetsClientWithAudit, AUDIT_LOG_TAB } from './sheetsAudit.js';
import type { sheets_v4 } from 'googleapis';

function fakeClient() {
  const calls: { method: string; params: any }[] = [];
  const record = (method: string) =>
    vi.fn(async (params: any) => {
      calls.push({ method, params });
      return { data: {} };
    });
  const client = {
    spreadsheets: {
      batchUpdate: record('batchUpdate'),
      values: {
        update: record('values.update'),
        append: record('values.append'),
        batchUpdate: record('values.batchUpdate'),
        clear: record('values.clear'),
        batchClear: record('values.batchClear'),
      },
    },
  } as unknown as sheets_v4.Sheets;
  return { client, calls };
}

describe('wrapSheetsClientWithAudit', () => {
  it('appends an audit row to the log tab after a values.update', async () => {
    const { client, calls } = fakeClient();
    const wrapped = wrapSheetsClientWithAudit(client);
    await wrapped.spreadsheets.values.update({ spreadsheetId: 'S1', range: 'Transactions!B5', requestBody: { values: [['x']] } } as any);
    const audit = calls.filter((c) => c.method === 'values.append');
    expect(audit).toHaveLength(1);
    expect(audit[0].params.range).toContain(AUDIT_LOG_TAB);
    const row = audit[0].params.requestBody.values[0];
    expect(row[1]).toBe('values.update');
    expect(row[2]).toContain('Transactions!B5');
    expect(row[3]).toBe('claude-mcp');
  });

  it('does not audit writes to the log tab itself (no recursion)', async () => {
    const { client, calls } = fakeClient();
    const wrapped = wrapSheetsClientWithAudit(client);
    await wrapped.spreadsheets.values.append({ spreadsheetId: 'S1', range: `'${AUDIT_LOG_TAB}'!A1`, requestBody: { values: [['x']] } } as any);
    // exactly the one user call, no audit append on top
    expect(calls.filter((c) => c.method === 'values.append')).toHaveLength(1);
  });

  it('records exact row indexes for deleteDimension batchUpdates', async () => {
    const { client, calls } = fakeClient();
    const wrapped = wrapSheetsClientWithAudit(client);
    await wrapped.spreadsheets.batchUpdate({
      spreadsheetId: 'S1',
      requestBody: { requests: [{ deleteDimension: { range: { sheetId: 7, dimension: 'ROWS', startIndex: 3520, endIndex: 3522 } } }] },
    } as any);
    const audit = calls.filter((c) => c.method === 'values.append');
    expect(audit).toHaveLength(1);
    const row = audit[0].params.requestBody.values[0];
    expect(row[1]).toBe('batchUpdate');
    expect(row[2]).toContain('deleteDimension ROWS 3521-3522');
    expect(row[2]).toContain('sheetId 7');
  });

  it('creates the log tab on first write when missing, then logs', async () => {
    const { client, calls } = fakeClient();
    // First append to the log tab fails like the real API does for a
    // missing tab; subsequent appends succeed.
    let appendCalls = 0;
    (client.spreadsheets.values.append as any) = vi.fn(async (params: any) => {
      calls.push({ method: 'values.append', params });
      appendCalls++;
      if (appendCalls === 1 && String(params.range).includes(AUDIT_LOG_TAB)) {
        throw new Error('Unable to parse range: Claude MCP Log!A1');
      }
      return { data: {} };
    });
    const wrapped = wrapSheetsClientWithAudit(client);
    await wrapped.spreadsheets.values.update({ spreadsheetId: 'S1', range: 'A1', requestBody: { values: [['x']] } } as any);
    const addSheet = calls.filter((c) => c.method === 'batchUpdate');
    expect(addSheet).toHaveLength(1);
    expect(addSheet[0].params.requestBody.requests[0].addSheet.properties.title).toBe(AUDIT_LOG_TAB);
    // header + row landed on the retry
    const retry = calls.filter((c) => c.method === 'values.append').at(-1)!;
    expect(retry.params.requestBody.values[0][0]).toBe('Timestamp (UTC)');
  });

  it('does not audit the log-tab creation batchUpdate itself', async () => {
    const { client, calls } = fakeClient();
    const wrapped = wrapSheetsClientWithAudit(client);
    await wrapped.spreadsheets.batchUpdate({
      spreadsheetId: 'S1',
      requestBody: { requests: [{ addSheet: { properties: { title: AUDIT_LOG_TAB } } }] },
    } as any);
    expect(calls.filter((c) => c.method === 'values.append')).toHaveLength(0);
  });
});
