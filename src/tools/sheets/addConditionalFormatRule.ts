import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'addConditionalFormatRule',
    description:
      'Adds a conditional formatting rule to a range in a Google Spreadsheet. Rules can highlight cells based on their values.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z.string().describe('The name of the sheet/tab to apply the rule to.'),
      range: z
        .string()
        .describe('A1 notation range to apply formatting to (e.g., "A1:D10" or "D3:F9").'),
      conditionType: z
        .enum([
          'NUMBER_GREATER',
          'NUMBER_LESS',
          'NUMBER_EQ',
          'NUMBER_GREATER_THAN_EQ',
          'NUMBER_LESS_THAN_EQ',
          'TEXT_CONTAINS',
          'TEXT_NOT_CONTAINS',
          'BLANK',
          'NOT_BLANK',
          'CUSTOM_FORMULA',
        ])
        .describe('The type of condition to check.'),
      conditionValue: z
        .string()
        .optional()
        .describe(
          'The value to compare against (e.g., "0" for NUMBER_GREATER than 0, or a formula for CUSTOM_FORMULA like "=$A1>0").'
        ),
      backgroundColor: z
        .string()
        .optional()
        .describe('Background color in hex format (e.g., "#f4cccc" for light red, "#d9ead3" for light green).'),
      textColor: z.string().optional().describe('Text color in hex format (e.g., "#FF0000" for red).'),
      bold: z.boolean().optional().describe('Whether to make the text bold.'),
      italic: z.boolean().optional().describe('Whether to make the text italic.'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Adding conditional format rule to spreadsheet ${args.spreadsheetId}, sheet: ${args.sheetName}, range: ${args.range}`
      );

      try {
        const metadata = await SheetsHelpers.getSpreadsheetMetadata(sheets, args.spreadsheetId);
        const sheet = metadata.sheets?.find((s) => s.properties?.title === args.sheetName);

        if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
          throw new UserError(`Sheet "${args.sheetName}" not found in spreadsheet.`);
        }

        const sheetId = sheet.properties.sheetId!;

        // Parse the A1 range
        const rangeMatch = args.range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
        if (!rangeMatch) {
          throw new UserError(`Invalid range format: ${args.range}. Expected format like "A1:D10"`);
        }

        function colToIndex(col: string): number {
          let index = 0;
          col = col.toUpperCase();
          for (let i = 0; i < col.length; i++) {
            index = index * 26 + (col.charCodeAt(i) - 64);
          }
          return index - 1;
        }

        const startCol = colToIndex(rangeMatch[1]);
        const startRow = parseInt(rangeMatch[2], 10) - 1;
        const endCol = colToIndex(rangeMatch[3]) + 1;
        const endRow = parseInt(rangeMatch[4], 10);

        const rule: {
          type: 'NUMBER_GREATER' | 'NUMBER_LESS' | 'NUMBER_EQ' | 'NUMBER_GREATER_THAN_EQ' | 'NUMBER_LESS_THAN_EQ' | 'TEXT_CONTAINS' | 'TEXT_NOT_CONTAINS' | 'BLANK' | 'NOT_BLANK' | 'CUSTOM_FORMULA';
          values?: string[];
          backgroundColor?: { red: number; green: number; blue: number };
          textColor?: { red: number; green: number; blue: number };
          bold?: boolean;
          italic?: boolean;
        } = {
          type: args.conditionType,
        };

        if (args.conditionValue) {
          rule.values = [args.conditionValue];
        }

        if (args.backgroundColor) {
          const bg = SheetsHelpers.hexToRgb(args.backgroundColor);
          if (bg) rule.backgroundColor = bg;
        }

        if (args.textColor) {
          const tc = SheetsHelpers.hexToRgb(args.textColor);
          if (tc) rule.textColor = tc;
        }

        if (args.bold !== undefined) rule.bold = args.bold;
        if (args.italic !== undefined) rule.italic = args.italic;

        await SheetsHelpers.addConditionalFormatRule(
          sheets,
          args.spreadsheetId,
          sheetId,
          {
            startRowIndex: startRow,
            endRowIndex: endRow,
            startColumnIndex: startCol,
            endColumnIndex: endCol,
          },
          rule
        );

        return `Successfully added conditional formatting rule to range ${args.range} on sheet "${args.sheetName}". Condition: ${args.conditionType}${args.conditionValue ? ` (value: ${args.conditionValue})` : ''}`;
      } catch (error: any) {
        log.error(`Error adding conditional format rule: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to add conditional format rule: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
