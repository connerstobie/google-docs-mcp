import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'deleteConditionalFormatRule',
    description: 'Deletes a specific conditional formatting rule by its index from a sheet.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z.string().describe('The name of the sheet containing the rule.'),
      ruleIndex: z
        .number()
        .int()
        .min(0)
        .describe(
          'The index of the rule to delete (0-based). Use getConditionalFormatRules to see rule indices.'
        ),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Deleting conditional format rule ${args.ruleIndex} from spreadsheet ${args.spreadsheetId}, sheet: ${args.sheetName}`
      );

      try {
        const metadata = await SheetsHelpers.getSpreadsheetMetadata(sheets, args.spreadsheetId);
        const sheet = metadata.sheets?.find((s) => s.properties?.title === args.sheetName);

        if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
          throw new UserError(`Sheet "${args.sheetName}" not found in spreadsheet.`);
        }

        const sheetId = sheet.properties.sheetId!;

        await SheetsHelpers.deleteConditionalFormatRule(
          sheets,
          args.spreadsheetId,
          sheetId,
          args.ruleIndex
        );

        return `Successfully deleted conditional formatting rule at index ${args.ruleIndex} from sheet "${args.sheetName}".`;
      } catch (error: any) {
        log.error(`Error deleting conditional format rule: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to delete conditional format rule: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
