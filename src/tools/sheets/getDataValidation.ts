import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'getDataValidation',
    description:
      'Gets data validation rules (dropdowns, list constraints, etc.) for cells in a Google Spreadsheet range. Returns the validation type, allowed values, and configuration for each cell that has validation.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      range: z
        .string()
        .describe(
          'A1 notation range to check for data validation (e.g., "Sheet1!A1:B5" or "\'Monthly Budget\'!H2").'
        ),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Getting data validation for range: ${args.range} in spreadsheet: ${args.spreadsheetId}`
      );

      try {
        const result = await SheetsHelpers.getDataValidation(
          sheets,
          args.spreadsheetId,
          args.range
        );

        if (result.validations.length === 0) {
          return `No data validation rules found in range ${args.range}.`;
        }

        let output = `**Data Validation Rules** for ${args.range}:\n\n`;

        for (const v of result.validations) {
          output += `**Cell ${v.cell}:**\n`;
          if (v.condition) {
            output += `- Type: ${v.condition.type}\n`;
            if (v.condition.values.length > 0) {
              output += `- Values: ${v.condition.values.map((val) => `\`${val}\``).join(', ')}\n`;
            }
          }
          if (v.inputMessage) {
            output += `- Input Message: ${v.inputMessage}\n`;
          }
          if (v.strict !== undefined) {
            output += `- Strict: ${v.strict}\n`;
          }
          if (v.showCustomUi !== undefined) {
            output += `- Show Dropdown: ${v.showCustomUi}\n`;
          }
          output += '\n';
        }

        return output;
      } catch (error: any) {
        log.error(`Error getting data validation for ${args.range}: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to get data validation: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
