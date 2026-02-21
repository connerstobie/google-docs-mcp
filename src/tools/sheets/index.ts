import type { FastMCP } from 'fastmcp';
import { register as readSpreadsheet } from './readSpreadsheet.js';
import { register as writeSpreadsheet } from './writeSpreadsheet.js';
import { register as appendSpreadsheetRows } from './appendSpreadsheetRows.js';
import { register as clearSpreadsheetRange } from './clearSpreadsheetRange.js';
import { register as getSpreadsheetInfo } from './getSpreadsheetInfo.js';
import { register as addSpreadsheetSheet } from './addSpreadsheetSheet.js';
import { register as createSpreadsheet } from './createSpreadsheet.js';
import { register as listGoogleSheets } from './listGoogleSheets.js';

// Formatting & validation
import { register as formatCells } from './formatCells.js';
import { register as freezeRowsAndColumns } from './freezeRowsAndColumns.js';
import { register as setDropdownValidation } from './setDropdownValidation.js';
import { register as formatSpreadsheet } from './formatSpreadsheet.js';

// Conditional formatting
import { register as addConditionalFormatRule } from './addConditionalFormatRule.js';
import { register as getConditionalFormatRules } from './getConditionalFormatRules.js';
import { register as deleteConditionalFormatRule } from './deleteConditionalFormatRule.js';
import { register as clearConditionalFormatRules } from './clearConditionalFormatRules.js';

// Row operations
import { register as deleteSpreadsheetRows } from './deleteSpreadsheetRows.js';

// Data validation reading
import { register as getDataValidation } from './getDataValidation.js';

export function registerSheetsTools(server: FastMCP) {
  readSpreadsheet(server);
  writeSpreadsheet(server);
  appendSpreadsheetRows(server);
  clearSpreadsheetRange(server);
  getSpreadsheetInfo(server);
  addSpreadsheetSheet(server);
  createSpreadsheet(server);
  listGoogleSheets(server);

  // Formatting & validation
  formatCells(server);
  freezeRowsAndColumns(server);
  setDropdownValidation(server);
  formatSpreadsheet(server);

  // Conditional formatting
  addConditionalFormatRule(server);
  getConditionalFormatRules(server);
  deleteConditionalFormatRule(server);
  clearConditionalFormatRules(server);

  // Row operations
  deleteSpreadsheetRows(server);

  // Data validation reading
  getDataValidation(server);
}
