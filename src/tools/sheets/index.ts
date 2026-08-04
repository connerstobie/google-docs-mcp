import type { FastMCP } from 'fastmcp';
import { register as readSpreadsheet } from './readSpreadsheet.js';
import { register as writeSpreadsheet } from './writeSpreadsheet.js';
import { register as batchWrite } from './batchWrite.js';
import { register as appendSpreadsheetRows } from './appendSpreadsheetRows.js';
import { register as clearSpreadsheetRange } from './clearSpreadsheetRange.js';
import { register as getSpreadsheetInfo } from './getSpreadsheetInfo.js';
import { register as addSpreadsheetSheet } from './addSpreadsheetSheet.js';
import { register as createSpreadsheet } from './createSpreadsheet.js';
import { register as listGoogleSheets } from './listGoogleSheets.js';
import { register as deleteSheet } from './deleteSheet.js';
import { register as renameSheet } from './renameSheet.js';
import { register as duplicateSheet } from './duplicateSheet.js';

// Formatting & validation
import { register as formatCells } from './formatCells.js';
import { register as readCellFormat } from './readCellFormat.js';
import { register as copyFormatting } from './copyFormatting.js';
import { register as freezeRowsAndColumns } from './freezeRowsAndColumns.js';
import { register as setColumnWidths } from './setColumnWidths.js';
import { register as autoResizeColumns } from './autoResizeColumns.js';
import { register as setDropdownValidation } from './setDropdownValidation.js';
import { register as formatSpreadsheet } from './formatSpreadsheet.js';
import { register as addConditionalFormatting } from './addConditionalFormatting.js';
import { register as groupRows } from './groupRows.js';
import { register as ungroupAllRows } from './ungroupAllRows.js';

// Conditional formatting (local tools)
import { register as addConditionalFormatRule } from './addConditionalFormatRule.js';
import { register as getConditionalFormatRules } from './getConditionalFormatRules.js';
import { register as deleteConditionalFormatRule } from './deleteConditionalFormatRule.js';
import { register as clearConditionalFormatRules } from './clearConditionalFormatRules.js';

// Row operations
import { register as deleteSpreadsheetRows } from './deleteSpreadsheetRows.js';

// Data validation reading
import { register as getDataValidation } from './getDataValidation.js';

// Cell formatting reading
import { register as getCellFormatting } from './getCellFormatting.js';

// Tables (upstream)
import { register as createTable } from './createTable.js';
import { register as listTables } from './listTables.js';
import { register as getTable } from './getTable.js';
import { register as deleteTable } from './deleteTable.js';
import { register as updateTableRange } from './updateTableRange.js';
import { register as appendTableRows } from './appendTableRows.js';
import { register as insertChart } from './insertChart.js';
import { register as deleteChart } from './deleteChart.js';
import { register as listCharts } from './listCharts.js';
import { register as updateChart } from './updateChart.js';

export function registerSheetsTools(server: FastMCP) {
  readSpreadsheet(server);
  writeSpreadsheet(server);
  batchWrite(server);
  appendSpreadsheetRows(server);
  clearSpreadsheetRange(server);
  getSpreadsheetInfo(server);
  addSpreadsheetSheet(server);
  createSpreadsheet(server);
  listGoogleSheets(server);
  deleteSheet(server);
  renameSheet(server);
  duplicateSheet(server);

  // Formatting & validation
  formatCells(server);
  readCellFormat(server);
  copyFormatting(server);
  freezeRowsAndColumns(server);
  setColumnWidths(server);
  autoResizeColumns(server);
  setDropdownValidation(server);
  formatSpreadsheet(server);
  addConditionalFormatting(server);
  groupRows(server);
  ungroupAllRows(server);

  // Conditional formatting (local tools)
  addConditionalFormatRule(server);
  getConditionalFormatRules(server);
  deleteConditionalFormatRule(server);
  clearConditionalFormatRules(server);

  // Row operations
  deleteSpreadsheetRows(server);

  // Data validation reading
  getDataValidation(server);

  // Cell formatting reading
  getCellFormatting(server);

  // Tables (upstream)
  createTable(server);
  listTables(server);
  getTable(server);
  deleteTable(server);
  updateTableRange(server);
  appendTableRows(server);
  insertChart(server);
  deleteChart(server);
  listCharts(server);
  updateChart(server);
}
