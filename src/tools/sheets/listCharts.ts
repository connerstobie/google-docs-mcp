import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import { summarizeChart } from '../../googleSheetsChartHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'listCharts',
    description:
      'Lists every chart in a spreadsheet with its numeric chart ID and full specification: title, chart type, ' +
      'stacking, legend position, axis titles, the domain (x-axis) range, and each series range in stacking order ' +
      'with its explicit color and target axis. Also reports the anchor cell and pixel size. ' +
      'IMPORTANT: charts are embedded objects, not cell values, so readSpreadsheet returns nothing for a sheet ' +
      'that contains only charts — never conclude such a tab is empty without calling this first. ' +
      'Use this to get a chart ID before calling updateChart or deleteChart.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe(
          'The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'
        ),
      sheetName: z
        .string()
        .optional()
        .describe('Only list charts anchored on this sheet/tab. Omit to list charts on all sheets.'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(`Listing charts in spreadsheet ${args.spreadsheetId}`);

      try {
        const res = await sheets.spreadsheets.get({
          spreadsheetId: args.spreadsheetId,
          fields: 'sheets(properties(sheetId,title),charts)',
        });

        const titleById = new Map<number, string>();
        for (const sheet of res.data.sheets ?? []) {
          const id = sheet.properties?.sheetId;
          const title = sheet.properties?.title;
          if (id !== undefined && id !== null && title) titleById.set(id, title);
        }

        if (args.sheetName && ![...titleById.values()].includes(args.sheetName)) {
          throw new UserError(
            `Sheet "${args.sheetName}" not found. Available: ${[...titleById.values()].join(', ')}`
          );
        }

        const charts = [];
        for (const sheet of res.data.sheets ?? []) {
          const sheetTitle = sheet.properties?.title ?? null;
          for (const chart of sheet.charts ?? []) {
            const summary = summarizeChart(chart, titleById);
            // A chart lives in the sheet it is anchored on; fall back to the
            // containing sheet when the anchor cell omits a sheetId.
            if (!summary.anchoredOnSheet) summary.anchoredOnSheet = sheetTitle;
            if (args.sheetName && summary.anchoredOnSheet !== args.sheetName) continue;
            charts.push(summary);
          }
        }

        if (charts.length === 0) {
          return args.sheetName
            ? `No charts found on sheet "${args.sheetName}".`
            : 'No charts found in this spreadsheet.';
        }

        return JSON.stringify({ count: charts.length, charts }, null, 2);
      } catch (error: any) {
        log.error(`Error listing charts: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to list charts: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
