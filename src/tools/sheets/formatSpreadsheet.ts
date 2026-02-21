import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'formatSpreadsheet',
    description:
      'Applies visual formatting to a Google Spreadsheet in a single batch operation. Supports: cell formatting (font size, bold, colors, number format, alignment), column widths, row heights, hiding columns, merging cells, and freezing rows/columns. All operations are batched into one API call for efficiency.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      sheetName: z.string().describe('The name of the sheet/tab to format.'),
      cellFormats: z
        .array(
          z.object({
            range: z.string().describe('A1 notation range (e.g., "A1:B2" or "A1").'),
            fontSize: z.number().optional().describe('Font size in points.'),
            bold: z.boolean().optional().describe('Whether to make text bold.'),
            italic: z.boolean().optional().describe('Whether to make text italic.'),
            fontFamily: z.string().optional().describe('Font family (e.g., "Arial", "Courier New").'),
            fontColor: z.string().optional().describe('Text color in hex (e.g., "#FF0000").'),
            backgroundColor: z
              .string()
              .optional()
              .describe('Background color in hex (e.g., "#d9ead3").'),
            horizontalAlignment: z
              .enum(['LEFT', 'CENTER', 'RIGHT'])
              .optional()
              .describe('Horizontal text alignment.'),
            verticalAlignment: z
              .enum(['TOP', 'MIDDLE', 'BOTTOM'])
              .optional()
              .describe('Vertical text alignment.'),
            numberFormat: z
              .string()
              .optional()
              .describe(
                'Number format pattern (e.g., "$#,##0.00" for currency, "0%" for percentage, "#,##0" for integers).'
              ),
            wrapStrategy: z
              .enum(['OVERFLOW_CELL', 'CLIP', 'WRAP'])
              .optional()
              .describe('Text wrap strategy.'),
          })
        )
        .optional()
        .describe('Array of cell formatting operations.'),
      columnWidths: z
        .array(
          z.object({
            column: z.string().describe('Column letter (e.g., "A", "B", "AA").'),
            width: z.number().describe('Width in pixels.'),
          })
        )
        .optional()
        .describe('Array of column width settings.'),
      rowHeights: z
        .array(
          z.object({
            row: z.number().int().min(1).describe('Row number (1-based).'),
            height: z.number().describe('Height in pixels.'),
          })
        )
        .optional()
        .describe('Array of row height settings.'),
      hideColumns: z
        .object({
          startColumn: z.string().describe('First column letter to hide (e.g., "H").'),
          endColumn: z.string().describe('Last column letter to hide, inclusive (e.g., "Z").'),
        })
        .optional()
        .describe('Range of columns to hide.'),
      merges: z
        .array(z.string())
        .optional()
        .describe('Array of A1 ranges to merge (e.g., ["A7:E7", "B17:C17"]).'),
      freezeRows: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('Number of rows to freeze at top (0 to unfreeze).'),
      freezeColumns: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('Number of columns to freeze at left (0 to unfreeze).'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(`Batch formatting sheet "${args.sheetName}" in spreadsheet ${args.spreadsheetId}`);

      try {
        await SheetsHelpers.batchFormat(sheets, args.spreadsheetId, args.sheetName, {
          cellFormats: args.cellFormats,
          columnWidths: args.columnWidths,
          rowHeights: args.rowHeights,
          hideColumns: args.hideColumns,
          merges: args.merges,
          freezeRows: args.freezeRows,
          freezeColumns: args.freezeColumns,
        });

        const parts: string[] = [];
        if (args.cellFormats?.length) parts.push(`${args.cellFormats.length} cell format(s)`);
        if (args.columnWidths?.length) parts.push(`${args.columnWidths.length} column width(s)`);
        if (args.rowHeights?.length) parts.push(`${args.rowHeights.length} row height(s)`);
        if (args.hideColumns)
          parts.push(`hidden columns ${args.hideColumns.startColumn}-${args.hideColumns.endColumn}`);
        if (args.merges?.length) parts.push(`${args.merges.length} merge(s)`);
        if (args.freezeRows !== undefined) parts.push(`${args.freezeRows} frozen row(s)`);
        if (args.freezeColumns !== undefined) parts.push(`${args.freezeColumns} frozen column(s)`);

        return `Successfully applied formatting to sheet "${args.sheetName}": ${parts.join(', ')}.`;
      } catch (error: any) {
        log.error(`Error formatting spreadsheet: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to format spreadsheet: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
