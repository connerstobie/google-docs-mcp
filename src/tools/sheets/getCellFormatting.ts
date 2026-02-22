import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getSheetsClient } from '../../clients.js';
import * as SheetsHelpers from '../../googleSheetsApiHelpers.js';

function rgbToHex(color: { red?: number | null; green?: number | null; blue?: number | null } | null | undefined): string {
  if (!color) return '#000000';
  const r = Math.round((color.red || 0) * 255);
  const g = Math.round((color.green || 0) * 255);
  const b = Math.round((color.blue || 0) * 255);
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`;
}

export function register(server: FastMCP) {
  server.addTool({
    name: 'getCellFormatting',
    description:
      'Reads the effective cell formatting (font, colors, backgrounds, number format, alignment, etc.) for a range of cells. Returns formatting details per cell so you can inspect or replicate existing styles.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe('The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'),
      range: z
        .string()
        .describe(
          'A1 notation range to read formatting from (e.g., "Sheet1!A1:C10" or "Schedule!A10:P16").'
        ),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(`Getting cell formatting for ${args.range} in spreadsheet ${args.spreadsheetId}`);

      try {
        const data = await SheetsHelpers.getCellFormatting(
          sheets,
          args.spreadsheetId,
          args.range
        );

        const sheetData = data.sheets?.[0];
        if (!sheetData?.data?.[0]?.rowData) {
          return 'No formatting data found for the specified range.';
        }

        const gridData = sheetData.data[0];
        const startRow = (gridData.startRow || 0) + 1; // Convert to 1-based
        const startCol = gridData.startColumn || 0;
        const sheetName = sheetData.properties?.title || 'Unknown';

        let output = `**Cell Formatting: ${sheetName}!${args.range.includes('!') ? args.range.split('!')[1] : args.range}**\n\n`;

        for (let r = 0; r < gridData.rowData!.length; r++) {
          const row = gridData.rowData![r];
          if (!row.values || row.values.length === 0) continue;

          const rowNum = startRow + r;

          for (let c = 0; c < row.values.length; c++) {
            const cell = row.values[c];
            const fmt = cell.effectiveFormat;
            if (!fmt) continue;

            const colLetter = SheetsHelpers.rowColToA1(0, startCol + c).replace(/\d+/, '');
            const cellRef = `${colLetter}${rowNum}`;

            const parts: string[] = [];

            // Background color
            if (fmt.backgroundColor) {
              const hex = rgbToHex(fmt.backgroundColor);
              if (hex !== '#ffffff') {
                parts.push(`bg: ${hex}`);
              }
            }

            // Text format
            const tf = fmt.textFormat;
            if (tf) {
              if (tf.foregroundColor) {
                const hex = rgbToHex(tf.foregroundColor);
                if (hex !== '#000000') {
                  parts.push(`color: ${hex}`);
                }
              }
              if (tf.fontFamily) parts.push(`font: ${tf.fontFamily}`);
              if (tf.fontSize) parts.push(`size: ${tf.fontSize}pt`);
              if (tf.bold) parts.push('bold');
              if (tf.italic) parts.push('italic');
              if (tf.strikethrough) parts.push('strikethrough');
              if (tf.underline) parts.push('underline');
            }

            // Alignment
            if (fmt.horizontalAlignment && fmt.horizontalAlignment !== 'LEFT') {
              parts.push(`align: ${fmt.horizontalAlignment}`);
            }

            // Number format
            if (fmt.numberFormat?.pattern) {
              parts.push(`numFmt: ${fmt.numberFormat.pattern}`);
            }

            // Only output if there's something notable
            if (parts.length > 0) {
              output += `${cellRef}: ${parts.join(' | ')}\n`;
            }
          }
        }

        if (output.split('\n').length <= 2) {
          output += 'All cells have default formatting (Arial 10pt, black text, white background).\n';
        }

        return output;
      } catch (error: any) {
        log.error(`Error getting cell formatting: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(
          `Failed to get cell formatting: ${error.message || 'Unknown error'}`
        );
      }
    },
  });
}
