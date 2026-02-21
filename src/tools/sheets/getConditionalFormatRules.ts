import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'getConditionalFormatRules',
    description:
      'Gets all conditional formatting rules for a Google Spreadsheet or a specific sheet.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z
        .string()
        .optional()
        .describe(
          'Optional: The name of a specific sheet to get rules for. If not provided, returns rules for all sheets.'
        ),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(
        `Getting conditional format rules for spreadsheet ${args.spreadsheetId}${args.sheetName ? `, sheet: ${args.sheetName}` : ''}`
      );

      try {
        let sheetId: number | undefined;

        if (args.sheetName) {
          const metadata = await SheetsHelpers.getSpreadsheetMetadata(sheets, args.spreadsheetId);
          const sheet = metadata.sheets?.find((s) => s.properties?.title === args.sheetName);
          if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
            throw new UserError(`Sheet "${args.sheetName}" not found in spreadsheet.`);
          }
          sheetId = sheet.properties.sheetId!;
        }

        const results = await SheetsHelpers.getConditionalFormatRules(
          sheets,
          args.spreadsheetId,
          sheetId
        );

        if (results.length === 0 || results.every((r) => r.rules.length === 0)) {
          return 'No conditional formatting rules found.';
        }

        let output = '**Conditional Formatting Rules:**\n\n';

        for (const sheetRules of results) {
          if (sheetRules.rules.length === 0) continue;

          output += `### Sheet: ${sheetRules.sheetName} (ID: ${sheetRules.sheetId})\n\n`;

          sheetRules.rules.forEach((rule, index) => {
            output += `**Rule ${index + 1}:**\n`;

            if (rule.ranges && rule.ranges.length > 0) {
              const range = rule.ranges[0];
              output += `- Range: Row ${(range.startRowIndex || 0) + 1} to ${range.endRowIndex || 'end'}, Col ${(range.startColumnIndex || 0) + 1} to ${range.endColumnIndex || 'end'}\n`;
            }

            if (rule.booleanRule) {
              const cond = rule.booleanRule.condition;
              output += `- Condition: ${cond?.type || 'Unknown'}`;
              if (cond?.values && cond.values.length > 0) {
                output += ` (${cond.values.map((v) => v.userEnteredValue || v.relativeDate || 'value').join(', ')})`;
              }
              output += '\n';

              const format = rule.booleanRule.format;
              if (format?.backgroundColor) {
                const bg = format.backgroundColor;
                output += `- Background: rgb(${Math.round((bg.red || 0) * 255)}, ${Math.round((bg.green || 0) * 255)}, ${Math.round((bg.blue || 0) * 255)})\n`;
              }
              if (format?.textFormat) {
                if (format.textFormat.foregroundColor) {
                  const fg = format.textFormat.foregroundColor;
                  output += `- Text Color: rgb(${Math.round((fg.red || 0) * 255)}, ${Math.round((fg.green || 0) * 255)}, ${Math.round((fg.blue || 0) * 255)})\n`;
                }
                if (format.textFormat.bold) output += `- Bold: true\n`;
                if (format.textFormat.italic) output += `- Italic: true\n`;
              }
            }

            if (rule.gradientRule) {
              output += `- Type: Gradient (color scale)\n`;
            }

            output += '\n';
          });
        }

        return output;
      } catch (error: any) {
        log.error(`Error getting conditional format rules: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to get conditional format rules: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
