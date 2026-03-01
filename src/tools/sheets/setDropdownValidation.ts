import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'setDropdownValidation',
    description:
      'Adds or removes a dropdown list on a range of cells. Provide values to create a dropdown restricting input to those options, or sourceRange to populate options from a cell range. Omit both to remove dropdown validation.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe(
          'The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'
        ),
      range: z
        .string()
        .describe(
          'A1 notation range to apply the dropdown to (e.g., "Sheet1!B2:B100" or "C2:C50").'
        ),
      values: z
        .array(z.string())
        .optional()
        .describe(
          'The allowed dropdown options (e.g., ["Open", "In Progress", "Done"]). Omit to remove existing dropdown validation from the range.'
        ),
      sourceRange: z
        .string()
        .optional()
        .describe(
          'A1 notation range whose values populate the dropdown (e.g., "Sheet1!$W$1:$W$7"). Creates a ONE_OF_RANGE validation that auto-updates when the source cells change. Takes precedence over values if both are provided.'
        ),
      strict: z
        .boolean()
        .optional()
        .default(true)
        .describe('If true, reject input that does not match one of the dropdown values.'),
      inputMessage: z
        .string()
        .optional()
        .describe('Help text shown when a cell with the dropdown is selected.'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      const isClearing = !args.sourceRange && (!args.values || args.values.length === 0);
      const isRange = !!args.sourceRange;
      log.info(
        isClearing
          ? `Clearing dropdown validation on "${args.range}" in spreadsheet ${args.spreadsheetId}`
          : isRange
            ? `Setting range-based dropdown on "${args.range}" from "${args.sourceRange}" in spreadsheet ${args.spreadsheetId}`
            : `Setting dropdown validation on "${args.range}" with ${args.values!.length} options in spreadsheet ${args.spreadsheetId}`
      );

      try {
        await SheetsHelpers.setDropdownValidation(
          sheets,
          args.spreadsheetId,
          args.range,
          args.sourceRange ? undefined : args.values,
          args.strict,
          args.inputMessage,
          args.sourceRange
        );

        if (isClearing) {
          return `Successfully removed dropdown validation from range "${args.range}".`;
        }
        if (isRange) {
          return `Successfully set range-based dropdown on "${args.range}" referencing "${args.sourceRange}".`;
        }
        return `Successfully added dropdown validation to range "${args.range}" with ${args.values!.length} options: ${args.values!.join(', ')}.`;
      } catch (error: any) {
        log.error(`Error setting dropdown validation: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to set dropdown validation: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
