import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'clearConditionalFormatRules',
    description: 'Removes all conditional formatting rules from a specific sheet.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z.string().describe('The name of the sheet to clear rules from.'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Clearing all conditional format rules from spreadsheet ${args.spreadsheetId}, sheet: ${args.sheetName}`
      );

      try {
        const metadata = await SheetsHelpers.getSpreadsheetMetadata(sheets, args.spreadsheetId);
        const sheet = metadata.sheets?.find((s) => s.properties?.title === args.sheetName);

        if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
          throw new UserError(`Sheet "${args.sheetName}" not found in spreadsheet.`);
        }

        const sheetId = sheet.properties.sheetId!;

        await SheetsHelpers.clearConditionalFormatRules(sheets, args.spreadsheetId, sheetId);

        return `Successfully cleared all conditional formatting rules from sheet "${args.sheetName}".`;
      } catch (error: any) {
        log.error(`Error clearing conditional format rules: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to clear conditional format rules: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
