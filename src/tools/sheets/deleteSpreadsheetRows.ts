import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'deleteSpreadsheetRows',
    description:
      'Deletes rows from a Google Spreadsheet. Rows are 1-based and inclusive. WARNING: This permanently removes rows and shifts all rows below up. When deleting multiple non-contiguous rows, delete from bottom to top to avoid index shifting.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z
        .string()
        .describe('The name of the sheet/tab containing the rows to delete.'),
      startRow: z
        .number()
        .int()
        .min(1)
        .describe('The starting row number to delete (1-based, inclusive).'),
      endRow: z
        .number()
        .int()
        .min(1)
        .describe('The ending row number to delete (1-based, inclusive).'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Deleting rows ${args.startRow}-${args.endRow} from sheet "${args.sheetName}" in spreadsheet ${args.spreadsheetId}`
      );

      try {
        await SheetsHelpers.deleteRows(
          sheets,
          args.spreadsheetId,
          args.sheetName,
          args.startRow,
          args.endRow
        );

        const count = args.endRow - args.startRow + 1;
        return `Successfully deleted ${count} row(s) (rows ${args.startRow}-${args.endRow}) from sheet "${args.sheetName}".`;
      } catch (error: any) {
        log.error(`Error deleting rows from spreadsheet: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to delete rows: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
