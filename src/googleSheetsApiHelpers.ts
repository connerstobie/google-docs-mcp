// src/googleSheetsApiHelpers.ts
import { google, sheets_v4 } from 'googleapis';
import { UserError } from 'fastmcp';

type Sheets = sheets_v4.Sheets; // Alias for convenience

// --- Core Helper Functions ---

/**
 * Converts A1 notation to row/column indices (0-based)
 * Example: "A1" -> {row: 0, col: 0}, "B2" -> {row: 1, col: 1}
 */
export function a1ToRowCol(a1: string): { row: number; col: number } {
  const match = a1.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new UserError(`Invalid A1 notation: ${a1}. Expected format like "A1" or "B2"`);
  }

  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10) - 1; // Convert to 0-based

  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1; // Convert to 0-based

  return { row, col };
}

/**
 * Converts row/column indices (0-based) to A1 notation
 * Example: {row: 0, col: 0} -> "A1", {row: 1, col: 1} -> "B2"
 */
export function rowColToA1(row: number, col: number): string {
  if (row < 0 || col < 0) {
    throw new UserError(
      `Row and column indices must be non-negative. Got row: ${row}, col: ${col}`
    );
  }

  let colStr = '';
  let colNum = col + 1; // Convert to 1-based for calculation
  while (colNum > 0) {
    colNum -= 1;
    colStr = String.fromCharCode(65 + (colNum % 26)) + colStr;
    colNum = Math.floor(colNum / 26);
  }

  return `${colStr}${row + 1}`;
}

/**
 * Validates and normalizes a range string
 * Examples: "A1" -> "Sheet1!A1", "A1:B2" -> "Sheet1!A1:B2"
 */
export function normalizeRange(range: string, sheetName?: string): string {
  // If range already contains '!', assume it's already normalized
  if (range.includes('!')) {
    return range;
  }

  // If sheetName is provided, prepend it
  if (sheetName) {
    return `${sheetName}!${range}`;
  }

  // Default to Sheet1 if no sheet name provided
  return `Sheet1!${range}`;
}

/**
 * Reads values from a spreadsheet range
 */
export async function readRange(
  sheets: Sheets,
  spreadsheetId: string,
  range: string,
  valueRenderOption: 'FORMATTED_VALUE' | 'UNFORMATTED_VALUE' | 'FORMULA' = 'FORMATTED_VALUE'
): Promise<sheets_v4.Schema$ValueRange> {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
      valueRenderOption,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have read access.`
      );
    }
    throw new UserError(`Failed to read range: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Writes values to a spreadsheet range
 */
export async function writeRange(
  sheets: Sheets,
  spreadsheetId: string,
  range: string,
  values: any[][],
  valueInputOption: 'RAW' | 'USER_ENTERED' = 'USER_ENTERED'
): Promise<sheets_v4.Schema$UpdateValuesResponse> {
  try {
    const response = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption,
      requestBody: {
        values,
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    throw new UserError(`Failed to write range: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Appends values to the end of a sheet
 */
export async function appendValues(
  sheets: Sheets,
  spreadsheetId: string,
  range: string,
  values: any[][],
  valueInputOption: 'RAW' | 'USER_ENTERED' = 'USER_ENTERED'
): Promise<sheets_v4.Schema$AppendValuesResponse> {
  try {
    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption,
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values,
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    throw new UserError(`Failed to append values: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Clears values from a range
 */
export async function clearRange(
  sheets: Sheets,
  spreadsheetId: string,
  range: string
): Promise<sheets_v4.Schema$ClearValuesResponse> {
  try {
    const response = await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    throw new UserError(`Failed to clear range: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Gets spreadsheet metadata including sheet information
 */
export async function getSpreadsheetMetadata(
  sheets: Sheets,
  spreadsheetId: string
): Promise<sheets_v4.Schema$Spreadsheet> {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: false,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have read access.`
      );
    }
    throw new UserError(`Failed to get spreadsheet metadata: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Gets cell formatting data for a range using includeGridData
 */
export async function getCellFormatting(
  sheets: Sheets,
  spreadsheetId: string,
  range: string
): Promise<sheets_v4.Schema$Spreadsheet> {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: [range],
      includeGridData: true,
      fields:
        'sheets.data.rowData.values.effectiveFormat,sheets.data.startRow,sheets.data.startColumn,sheets.properties.title,sheets.properties.sheetId',
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have read access.`
      );
    }
    throw new UserError(`Failed to get cell formatting: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Creates a new sheet/tab in a spreadsheet
 */
export async function addSheet(
  sheets: Sheets,
  spreadsheetId: string,
  sheetTitle: string
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addSheet: {
              properties: {
                title: sheetTitle,
              },
            },
          },
        ],
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    throw new UserError(`Failed to add sheet: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Parses A1 notation range to extract sheet name and cell range
 * Returns {sheetName, a1Range} where a1Range is just the cell part (e.g., "A1:B2")
 */
function parseRange(range: string): { sheetName: string | null; a1Range: string } {
  if (range.includes('!')) {
    const parts = range.split('!');
    return {
      sheetName: parts[0].replace(/^'|'$/g, ''), // Remove quotes if present
      a1Range: parts[1],
    };
  }
  return {
    sheetName: null,
    a1Range: range,
  };
}

/**
 * Resolves a sheet name to a numeric sheet ID.
 * If sheetName is null/undefined, returns the first sheet's ID.
 */
async function resolveSheetId(
  sheets: Sheets,
  spreadsheetId: string,
  sheetName?: string | null
): Promise<number> {
  const metadata = await getSpreadsheetMetadata(sheets, spreadsheetId);

  if (sheetName) {
    const sheet = metadata.sheets?.find((s) => s.properties?.title === sheetName);
    if (!sheet || sheet.properties?.sheetId === undefined || sheet.properties?.sheetId === null) {
      throw new UserError(`Sheet "${sheetName}" not found in spreadsheet.`);
    }
    return sheet.properties.sheetId;
  }

  const firstSheet = metadata.sheets?.[0];
  if (firstSheet?.properties?.sheetId === undefined || firstSheet?.properties?.sheetId === null) {
    throw new UserError('Spreadsheet has no sheets.');
  }
  return firstSheet.properties.sheetId;
}

/**
 * Converts column letters to a 0-based column index.
 * Example: "A" -> 0, "B" -> 1, "Z" -> 25, "AA" -> 26
 */
function colLettersToIndex(col: string): number {
  let index = 0;
  const upper = col.toUpperCase();
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index - 1;
}

/**
 * Parses an A1-notation cell range string into a Google Sheets GridRange object.
 * Supports:
 *   - Standard: "A1", "A1:B2"
 *   - Whole rows: "1:1", "1:3"
 *   - Whole columns: "A:A", "A:C"
 * When a component is omitted (whole row/column), the corresponding
 * start/end index is left out of the GridRange, which the Sheets API
 * interprets as "unbounded" (i.e., the entire row or column).
 */
function parseA1ToGridRange(a1Range: string, sheetId: number): sheets_v4.Schema$GridRange {
  // Whole-row pattern: "1:3" or "1"
  const rowOnlyMatch = a1Range.match(/^(\d+)(?::(\d+))?$/);
  if (rowOnlyMatch) {
    const startRow = parseInt(rowOnlyMatch[1], 10) - 1;
    const endRow = rowOnlyMatch[2] ? parseInt(rowOnlyMatch[2], 10) : startRow + 1;
    return {
      sheetId,
      startRowIndex: startRow,
      endRowIndex: endRow,
      // no column indices → entire row
    };
  }

  // Whole-column pattern: "A:C" or "A"
  const colOnlyMatch = a1Range.match(/^([A-Z]+)(?::([A-Z]+))?$/i);
  if (colOnlyMatch && !/\d/.test(a1Range)) {
    const startCol = colLettersToIndex(colOnlyMatch[1]);
    const endCol = colOnlyMatch[2] ? colLettersToIndex(colOnlyMatch[2]) + 1 : startCol + 1;
    return {
      sheetId,
      startColumnIndex: startCol,
      endColumnIndex: endCol,
      // no row indices → entire column
    };
  }

  // Standard A1 pattern: "A1" or "A1:B2"
  const standardMatch = a1Range.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!standardMatch) {
    throw new UserError(
      `Invalid range format: "${a1Range}". Expected "A1:B2", "1:1" (whole row), or "A:A" (whole column).`
    );
  }

  const startCol = colLettersToIndex(standardMatch[1]);
  const startRow = parseInt(standardMatch[2], 10) - 1;
  const endCol = standardMatch[3] ? colLettersToIndex(standardMatch[3]) + 1 : startCol + 1;
  const endRow = standardMatch[4] ? parseInt(standardMatch[4], 10) : startRow + 1;

  return {
    sheetId,
    startRowIndex: startRow,
    endRowIndex: endRow,
    startColumnIndex: startCol,
    endColumnIndex: endCol,
  };
}

/**
 * Formats cells in a range.
 * Supports standard A1 ranges ("A1:D1"), whole-row ("1:1"), and whole-column ("A:A") notation.
 */
export async function formatCells(
  sheets: Sheets,
  spreadsheetId: string,
  range: string,
  format: {
    backgroundColor?: { red: number; green: number; blue: number };
    textFormat?: {
      foregroundColor?: { red: number; green: number; blue: number };
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
    };
    horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT';
    verticalAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
  }
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    // Parse the range to get sheet name and cell range
    const { sheetName, a1Range } = parseRange(range);
    const sheetId = await resolveSheetId(sheets, spreadsheetId, sheetName);

    // Parse A1 range to get row/column indices
    // Supports: "A1:B2" (standard), "1:3" (whole rows), "A:C" (whole columns)
    const gridRange = parseA1ToGridRange(a1Range, sheetId);

    const userEnteredFormat: sheets_v4.Schema$CellFormat = {};

    if (format.backgroundColor) {
      userEnteredFormat.backgroundColor = {
        red: format.backgroundColor.red,
        green: format.backgroundColor.green,
        blue: format.backgroundColor.blue,
        alpha: 1,
      };
    }

    if (format.textFormat) {
      userEnteredFormat.textFormat = {};
      if (format.textFormat.foregroundColor) {
        userEnteredFormat.textFormat.foregroundColor = {
          red: format.textFormat.foregroundColor.red,
          green: format.textFormat.foregroundColor.green,
          blue: format.textFormat.foregroundColor.blue,
          alpha: 1,
        };
      }
      if (format.textFormat.fontSize !== undefined) {
        userEnteredFormat.textFormat.fontSize = format.textFormat.fontSize;
      }
      if (format.textFormat.bold !== undefined) {
        userEnteredFormat.textFormat.bold = format.textFormat.bold;
      }
      if (format.textFormat.italic !== undefined) {
        userEnteredFormat.textFormat.italic = format.textFormat.italic;
      }
    }

    if (format.horizontalAlignment) {
      userEnteredFormat.horizontalAlignment = format.horizontalAlignment;
    }

    if (format.verticalAlignment) {
      userEnteredFormat.verticalAlignment = format.verticalAlignment;
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: gridRange,
              cell: {
                userEnteredFormat,
              },
              fields:
                'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)',
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to format cells: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Freezes rows and/or columns in a sheet so they remain visible when scrolling.
 */
export async function freezeRowsAndColumns(
  sheets: Sheets,
  spreadsheetId: string,
  sheetName?: string | null,
  frozenRows?: number,
  frozenColumns?: number
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const sheetId = await resolveSheetId(sheets, spreadsheetId, sheetName);

    const gridProperties: sheets_v4.Schema$GridProperties = {};
    const fieldParts: string[] = [];

    if (frozenRows !== undefined) {
      gridProperties.frozenRowCount = frozenRows;
      fieldParts.push('gridProperties.frozenRowCount');
    }
    if (frozenColumns !== undefined) {
      gridProperties.frozenColumnCount = frozenColumns;
      fieldParts.push('gridProperties.frozenColumnCount');
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId,
                gridProperties,
              },
              fields: fieldParts.join(','),
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to freeze rows/columns: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Sets or clears dropdown data validation on a range of cells.
 * When values are provided, creates a ONE_OF_LIST validation rule.
 * When sourceRange is provided, creates a ONE_OF_RANGE validation rule that
 * auto-updates when the source cells change.
 * When neither is provided, clears any existing validation from the range.
 */
export async function setDropdownValidation(
  sheets: Sheets,
  spreadsheetId: string,
  range: string,
  values?: string[],
  strict: boolean = true,
  inputMessage?: string,
  sourceRange?: string
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const { sheetName, a1Range } = parseRange(range);
    const sheetId = await resolveSheetId(sheets, spreadsheetId, sheetName);
    const gridRange = parseA1ToGridRange(a1Range, sheetId);

    let rule: sheets_v4.Schema$DataValidationRule | undefined;

    if (sourceRange) {
      // ONE_OF_RANGE: dropdown populated from a cell range
      const { sheetName: srcSheet, a1Range: srcA1 } = parseRange(sourceRange);
      const srcSheetId = await resolveSheetId(sheets, spreadsheetId, srcSheet);
      const srcGridRange = parseA1ToGridRange(srcA1, srcSheetId);
      rule = {
        condition: {
          type: 'ONE_OF_RANGE' as const,
          values: [
            {
              userEnteredValue: `=${srcSheet ? `'${srcSheet}'!` : ''}${srcA1}`,
            },
          ],
        },
        showCustomUi: true,
        strict,
        inputMessage: inputMessage || null,
      };
    } else if (values && values.length > 0) {
      // ONE_OF_LIST: dropdown with hardcoded values
      rule = {
        condition: {
          type: 'ONE_OF_LIST' as const,
          values: values.map((v) => ({ userEnteredValue: v })),
        },
        showCustomUi: true,
        strict,
        inputMessage: inputMessage || null,
      };
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            setDataValidation: {
              range: gridRange,
              rule,
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(
        `Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`
      );
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to set dropdown validation: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Adds a conditional formatting rule to a spreadsheet range
 */
export async function addConditionalFormatRule(
  sheets: Sheets,
  spreadsheetId: string,
  sheetId: number,
  range: {
    startRowIndex: number;
    endRowIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
  },
  rule: {
    type: 'NUMBER_GREATER' | 'NUMBER_LESS' | 'NUMBER_EQ' | 'NUMBER_GREATER_THAN_EQ' | 'NUMBER_LESS_THAN_EQ' | 'TEXT_CONTAINS' | 'TEXT_NOT_CONTAINS' | 'BLANK' | 'NOT_BLANK' | 'CUSTOM_FORMULA';
    values?: string[];
    backgroundColor?: { red: number; green: number; blue: number };
    textColor?: { red: number; green: number; blue: number };
    bold?: boolean;
    italic?: boolean;
  },
  index: number = 0
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const format: sheets_v4.Schema$CellFormat = {};

    if (rule.backgroundColor) {
      format.backgroundColor = {
        red: rule.backgroundColor.red,
        green: rule.backgroundColor.green,
        blue: rule.backgroundColor.blue,
        alpha: 1,
      };
    }

    if (rule.textColor || rule.bold !== undefined || rule.italic !== undefined) {
      format.textFormat = {};
      if (rule.textColor) {
        format.textFormat.foregroundColor = {
          red: rule.textColor.red,
          green: rule.textColor.green,
          blue: rule.textColor.blue,
          alpha: 1,
        };
      }
      if (rule.bold !== undefined) {
        format.textFormat.bold = rule.bold;
      }
      if (rule.italic !== undefined) {
        format.textFormat.italic = rule.italic;
      }
    }

    const booleanCondition: sheets_v4.Schema$BooleanCondition = {
      type: rule.type,
    };

    if (rule.values && rule.values.length > 0) {
      booleanCondition.values = rule.values.map(v => ({ userEnteredValue: v }));
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            addConditionalFormatRule: {
              rule: {
                ranges: [
                  {
                    sheetId,
                    startRowIndex: range.startRowIndex,
                    endRowIndex: range.endRowIndex,
                    startColumnIndex: range.startColumnIndex,
                    endColumnIndex: range.endColumnIndex,
                  },
                ],
                booleanRule: {
                  condition: booleanCondition,
                  format,
                },
              },
              index,
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`);
    }
    throw new UserError(`Failed to add conditional format rule: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Gets all conditional formatting rules for a spreadsheet
 */
export async function getConditionalFormatRules(
  sheets: Sheets,
  spreadsheetId: string,
  sheetId?: number
): Promise<{ sheetId: number; sheetName: string; rules: sheets_v4.Schema$ConditionalFormatRule[] }[]> {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      includeGridData: false,
      fields: 'sheets(properties(sheetId,title),conditionalFormats)',
    });

    const results: { sheetId: number; sheetName: string; rules: sheets_v4.Schema$ConditionalFormatRule[] }[] = [];

    for (const sheet of response.data.sheets || []) {
      const currentSheetId = sheet.properties?.sheetId;
      const sheetName = sheet.properties?.title || 'Unknown';

      if (sheetId !== undefined && currentSheetId !== sheetId) {
        continue;
      }

      if (currentSheetId != null) {
        results.push({
          sheetId: currentSheetId,
          sheetName,
          rules: sheet.conditionalFormats || [],
        });
      }
    }

    return results;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have read access.`);
    }
    throw new UserError(`Failed to get conditional format rules: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Deletes a conditional formatting rule by index
 */
export async function deleteConditionalFormatRule(
  sheets: Sheets,
  spreadsheetId: string,
  sheetId: number,
  index: number
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            deleteConditionalFormatRule: {
              sheetId,
              index,
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`);
    }
    throw new UserError(`Failed to delete conditional format rule: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Clears all conditional formatting rules from a sheet
 */
export async function clearConditionalFormatRules(
  sheets: Sheets,
  spreadsheetId: string,
  sheetId: number
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const existingRules = await getConditionalFormatRules(sheets, spreadsheetId, sheetId);
    const sheetRules = existingRules.find(s => s.sheetId === sheetId);

    if (!sheetRules || sheetRules.rules.length === 0) {
      return { spreadsheetId };
    }

    const requests = [];
    for (let i = sheetRules.rules.length - 1; i >= 0; i--) {
      requests.push({
        deleteConditionalFormatRule: {
          sheetId,
          index: i,
        },
      });
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`);
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to clear conditional format rules: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Batch format a spreadsheet with multiple formatting operations in a single API call.
 */
export async function batchFormat(
  sheets: Sheets,
  spreadsheetId: string,
  sheetName: string,
  operations: {
    cellFormats?: Array<{
      range: string;
      fontSize?: number;
      bold?: boolean;
      italic?: boolean;
      fontFamily?: string;
      fontColor?: string;
      backgroundColor?: string;
      horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT';
      verticalAlignment?: 'TOP' | 'MIDDLE' | 'BOTTOM';
      numberFormat?: string;
      wrapStrategy?: 'OVERFLOW_CELL' | 'CLIP' | 'WRAP';
    }>;
    columnWidths?: Array<{ column: string; width: number }>;
    rowHeights?: Array<{ row: number; height: number }>;
    hideColumns?: { startColumn: string; endColumn: string };
    merges?: string[];
    freezeRows?: number;
    freezeColumns?: number;
  }
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const metadata = await getSpreadsheetMetadata(sheets, spreadsheetId);
    const sheet = metadata.sheets?.find(s => s.properties?.title === sheetName);
    if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
      throw new UserError(`Sheet "${sheetName}" not found in spreadsheet.`);
    }
    const sheetId = sheet.properties.sheetId!;

    function colLetterToIndex(col: string): number {
      let idx = 0;
      col = col.toUpperCase();
      for (let i = 0; i < col.length; i++) {
        idx = idx * 26 + (col.charCodeAt(i) - 64);
      }
      return idx - 1;
    }

    function parseA1Range(a1: string): { startRow: number; endRow: number; startCol: number; endCol: number } {
      const match = a1.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
      if (!match) throw new UserError(`Invalid A1 range: ${a1}`);
      const startCol = colLetterToIndex(match[1]);
      const startRow = parseInt(match[2], 10) - 1;
      const endCol = match[3] ? colLetterToIndex(match[3]) : startCol;
      const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;
      return { startRow, endRow, startCol, endCol };
    }

    const requests: sheets_v4.Schema$Request[] = [];

    if (operations.cellFormats) {
      for (const fmt of operations.cellFormats) {
        const { startRow, endRow, startCol, endCol } = parseA1Range(fmt.range);
        const userEnteredFormat: sheets_v4.Schema$CellFormat = {};
        const fields: string[] = [];

        if (fmt.backgroundColor) {
          const bg = hexToRgb(fmt.backgroundColor);
          if (bg) {
            userEnteredFormat.backgroundColor = { ...bg, alpha: 1 };
            fields.push('userEnteredFormat.backgroundColor');
          }
        }

        const textFormat: sheets_v4.Schema$TextFormat = {};
        let hasTextFormat = false;

        if (fmt.fontColor) {
          const fc = hexToRgb(fmt.fontColor);
          if (fc) {
            textFormat.foregroundColor = { ...fc, alpha: 1 };
            hasTextFormat = true;
            fields.push('userEnteredFormat.textFormat.foregroundColor');
          }
        }
        if (fmt.fontSize !== undefined) {
          textFormat.fontSize = fmt.fontSize;
          hasTextFormat = true;
          fields.push('userEnteredFormat.textFormat.fontSize');
        }
        if (fmt.bold !== undefined) {
          textFormat.bold = fmt.bold;
          hasTextFormat = true;
          fields.push('userEnteredFormat.textFormat.bold');
        }
        if (fmt.italic !== undefined) {
          textFormat.italic = fmt.italic;
          hasTextFormat = true;
          fields.push('userEnteredFormat.textFormat.italic');
        }
        if (fmt.fontFamily) {
          textFormat.fontFamily = fmt.fontFamily;
          hasTextFormat = true;
          fields.push('userEnteredFormat.textFormat.fontFamily');
        }

        if (hasTextFormat) {
          userEnteredFormat.textFormat = textFormat;
        }

        if (fmt.horizontalAlignment) {
          userEnteredFormat.horizontalAlignment = fmt.horizontalAlignment;
          fields.push('userEnteredFormat.horizontalAlignment');
        }
        if (fmt.verticalAlignment) {
          userEnteredFormat.verticalAlignment = fmt.verticalAlignment;
          fields.push('userEnteredFormat.verticalAlignment');
        }
        if (fmt.numberFormat) {
          userEnteredFormat.numberFormat = { type: 'NUMBER', pattern: fmt.numberFormat };
          fields.push('userEnteredFormat.numberFormat');
        }
        if (fmt.wrapStrategy) {
          userEnteredFormat.wrapStrategy = fmt.wrapStrategy;
          fields.push('userEnteredFormat.wrapStrategy');
        }

        if (fields.length > 0) {
          requests.push({
            repeatCell: {
              range: {
                sheetId,
                startRowIndex: startRow,
                endRowIndex: endRow + 1,
                startColumnIndex: startCol,
                endColumnIndex: endCol + 1,
              },
              cell: { userEnteredFormat },
              fields: fields.join(','),
            },
          });
        }
      }
    }

    if (operations.columnWidths) {
      for (const cw of operations.columnWidths) {
        const colIndex = colLetterToIndex(cw.column);
        requests.push({
          updateDimensionProperties: {
            range: { sheetId, dimension: 'COLUMNS', startIndex: colIndex, endIndex: colIndex + 1 },
            properties: { pixelSize: cw.width },
            fields: 'pixelSize',
          },
        });
      }
    }

    if (operations.rowHeights) {
      for (const rh of operations.rowHeights) {
        requests.push({
          updateDimensionProperties: {
            range: { sheetId, dimension: 'ROWS', startIndex: rh.row - 1, endIndex: rh.row },
            properties: { pixelSize: rh.height },
            fields: 'pixelSize',
          },
        });
      }
    }

    if (operations.hideColumns) {
      const startCol = colLetterToIndex(operations.hideColumns.startColumn);
      const endCol = colLetterToIndex(operations.hideColumns.endColumn);
      requests.push({
        updateDimensionProperties: {
          range: { sheetId, dimension: 'COLUMNS', startIndex: startCol, endIndex: endCol + 1 },
          properties: { hiddenByUser: true },
          fields: 'hiddenByUser',
        },
      });
    }

    if (operations.merges) {
      for (const mergeRange of operations.merges) {
        const { startRow, endRow, startCol, endCol } = parseA1Range(mergeRange);
        requests.push({
          mergeCells: {
            range: {
              sheetId,
              startRowIndex: startRow,
              endRowIndex: endRow + 1,
              startColumnIndex: startCol,
              endColumnIndex: endCol + 1,
            },
            mergeType: 'MERGE_ALL',
          },
        });
      }
    }

    if (operations.freezeRows !== undefined || operations.freezeColumns !== undefined) {
      const gridProperties: sheets_v4.Schema$GridProperties = {};
      const fields: string[] = [];
      if (operations.freezeRows !== undefined) {
        gridProperties.frozenRowCount = operations.freezeRows;
        fields.push('gridProperties.frozenRowCount');
      }
      if (operations.freezeColumns !== undefined) {
        gridProperties.frozenColumnCount = operations.freezeColumns;
        fields.push('gridProperties.frozenColumnCount');
      }
      requests.push({
        updateSheetProperties: {
          properties: { sheetId, gridProperties },
          fields: fields.join(','),
        },
      });
    }

    if (requests.length === 0) {
      return { spreadsheetId };
    }

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`);
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to batch format: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Deletes rows from a sheet
 */
export async function deleteRows(
  sheets: Sheets,
  spreadsheetId: string,
  sheetName: string,
  startRow: number,
  endRow: number
): Promise<sheets_v4.Schema$BatchUpdateSpreadsheetResponse> {
  try {
    const metadata = await getSpreadsheetMetadata(sheets, spreadsheetId);
    const sheet = metadata.sheets?.find(s => s.properties?.title === sheetName);

    if (sheet?.properties?.sheetId == null) {
      throw new UserError(`Sheet "${sheetName}" not found in spreadsheet.`);
    }

    const sheetId = sheet.properties.sheetId;

    if (startRow < 1 || endRow < 1) {
      throw new UserError('Row numbers must be 1-based (minimum 1).');
    }

    if (startRow > endRow) {
      throw new UserError('startRow cannot be greater than endRow.');
    }

    const startIndex = startRow - 1;
    const endIndex = endRow;

    const response = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: { sheetId, dimension: 'ROWS', startIndex, endIndex },
            },
          },
        ],
      },
    });

    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have write access.`);
    }
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to delete rows: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Gets data validation rules for a range of cells.
 */
export async function getDataValidation(
  sheets: Sheets,
  spreadsheetId: string,
  range: string
): Promise<{
  range: string;
  validations: Array<{
    cell: string;
    condition?: { type: string; values: string[] };
    inputMessage?: string;
    strict?: boolean;
    showCustomUi?: boolean;
  }>;
}> {
  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: [range],
      includeGridData: true,
      fields: 'sheets(properties(title,sheetId),data(startRow,startColumn,rowData(values(dataValidation,formattedValue))))',
    });

    const validations: Array<{
      cell: string;
      condition?: { type: string; values: string[] };
      inputMessage?: string;
      strict?: boolean;
      showCustomUi?: boolean;
    }> = [];

    const sheetData = response.data.sheets?.[0];
    const gridData = sheetData?.data?.[0];

    if (!gridData?.rowData) {
      return { range, validations };
    }

    const startRow = gridData.startRow ?? 0;
    const startCol = gridData.startColumn ?? 0;

    for (let rowIdx = 0; rowIdx < gridData.rowData.length; rowIdx++) {
      const row = gridData.rowData[rowIdx];
      if (!row.values) continue;

      for (let colIdx = 0; colIdx < row.values.length; colIdx++) {
        const cellData = row.values[colIdx];
        const cellRef = rowColToA1(startRow + rowIdx, startCol + colIdx);

        if (cellData.dataValidation) {
          const dv = cellData.dataValidation;
          const conditionValues = dv.condition?.values?.map(
            (v) => v.userEnteredValue || v.relativeDate || ''
          ) ?? [];

          validations.push({
            cell: cellRef,
            condition: dv.condition ? {
              type: dv.condition.type || 'UNKNOWN',
              values: conditionValues,
            } : undefined,
            inputMessage: dv.inputMessage || undefined,
            strict: dv.strict ?? undefined,
            showCustomUi: dv.showCustomUi ?? undefined,
          });
        }
      }
    }

    return { range, validations };
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Spreadsheet not found (ID: ${spreadsheetId}). Check the ID.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for spreadsheet (ID: ${spreadsheetId}). Ensure you have read access.`);
    }
    throw new UserError(`Failed to get data validation: ${error.message || 'Unknown error'}`);
  }
}

/**
 * Helper to convert hex color to RGB (0-1 range)
 */
export function hexToRgb(hex: string): { red: number; green: number; blue: number } | null {
  if (!hex) return null;
  let hexClean = hex.startsWith('#') ? hex.slice(1) : hex;

  if (hexClean.length === 3) {
    hexClean = hexClean[0] + hexClean[0] + hexClean[1] + hexClean[1] + hexClean[2] + hexClean[2];
  }
  if (hexClean.length !== 6) return null;
  const bigint = parseInt(hexClean, 16);
  if (isNaN(bigint)) return null;

  return {
    red: ((bigint >> 16) & 255) / 255,
    green: ((bigint >> 8) & 255) / 255,
    blue: (bigint & 255) / 255,
  };
}
