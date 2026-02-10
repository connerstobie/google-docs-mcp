// tests/conditionalFormatting.test.js
import {
  addConditionalFormatRule,
  getConditionalFormatRules,
  deleteConditionalFormatRule,
  clearConditionalFormatRules,
} from '../dist/googleSheetsApiHelpers.js';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

describe('Conditional Formatting', () => {
  describe('addConditionalFormatRule', () => {
    it('should add a conditional format rule with backgroundColor', async () => {
      const mockSheets = {
        spreadsheets: {
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}] },
          })),
        },
      };

      const result = await addConditionalFormatRule(
        mockSheets,
        'sheet123',
        0,
        { startRowIndex: 2, endRowIndex: 8, startColumnIndex: 3, endColumnIndex: 6 },
        {
          type: 'NUMBER_GREATER',
          values: ['0'],
          backgroundColor: { red: 0.96, green: 0.8, blue: 0.8 },
        }
      );

      assert.strictEqual(result.spreadsheetId, 'sheet123');
      assert.strictEqual(mockSheets.spreadsheets.batchUpdate.mock.calls.length, 1);

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      const rule = requestBody.requests[0].addConditionalFormatRule.rule;

      assert.strictEqual(rule.booleanRule.condition.type, 'NUMBER_GREATER');
      assert.deepStrictEqual(rule.booleanRule.condition.values, [{ userEnteredValue: '0' }]);
      assert.deepStrictEqual(rule.booleanRule.format.backgroundColor, {
        red: 0.96,
        green: 0.8,
        blue: 0.8,
        alpha: 1,
      });
      assert.deepStrictEqual(rule.ranges[0], {
        sheetId: 0,
        startRowIndex: 2,
        endRowIndex: 8,
        startColumnIndex: 3,
        endColumnIndex: 6,
      });
    });

    it('should add a rule with text formatting (bold, italic, textColor)', async () => {
      const mockSheets = {
        spreadsheets: {
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}] },
          })),
        },
      };

      await addConditionalFormatRule(
        mockSheets,
        'sheet123',
        0,
        { startRowIndex: 0, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 1 },
        {
          type: 'TEXT_CONTAINS',
          values: ['error'],
          textColor: { red: 1, green: 0, blue: 0 },
          bold: true,
          italic: false,
        }
      );

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      const format = requestBody.requests[0].addConditionalFormatRule.rule.booleanRule.format;

      assert.strictEqual(format.textFormat.bold, true);
      assert.strictEqual(format.textFormat.italic, false);
      assert.deepStrictEqual(format.textFormat.foregroundColor, {
        red: 1,
        green: 0,
        blue: 0,
        alpha: 1,
      });
    });

    it('should add a rule with custom formula type', async () => {
      const mockSheets = {
        spreadsheets: {
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}] },
          })),
        },
      };

      await addConditionalFormatRule(
        mockSheets,
        'sheet123',
        0,
        { startRowIndex: 0, endRowIndex: 100, startColumnIndex: 0, endColumnIndex: 10 },
        {
          type: 'CUSTOM_FORMULA',
          values: ['=A1>B1'],
          backgroundColor: { red: 0.8, green: 0.96, blue: 0.8 },
        }
      );

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      const condition = requestBody.requests[0].addConditionalFormatRule.rule.booleanRule.condition;

      assert.strictEqual(condition.type, 'CUSTOM_FORMULA');
      assert.deepStrictEqual(condition.values, [{ userEnteredValue: '=A1>B1' }]);
    });

    it('should use the provided rule index', async () => {
      const mockSheets = {
        spreadsheets: {
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}] },
          })),
        },
      };

      await addConditionalFormatRule(
        mockSheets,
        'sheet123',
        0,
        { startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
        { type: 'BLANK', backgroundColor: { red: 1, green: 1, blue: 0 } },
        5 // index parameter
      );

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      assert.strictEqual(requestBody.requests[0].addConditionalFormatRule.index, 5);
    });
  });

  describe('getConditionalFormatRules', () => {
    it('should return all conditional format rules for all sheets', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [
                {
                  properties: { sheetId: 0, title: 'Sheet1' },
                  conditionalFormats: [
                    {
                      ranges: [{ sheetId: 0, startRowIndex: 0, endRowIndex: 10 }],
                      booleanRule: { condition: { type: 'NUMBER_GREATER', values: [{ userEnteredValue: '0' }] } },
                    },
                  ],
                },
                {
                  properties: { sheetId: 1, title: 'Sheet2' },
                  conditionalFormats: [],
                },
              ],
            },
          })),
        },
      };

      const result = await getConditionalFormatRules(mockSheets, 'sheet123');

      assert.strictEqual(result.length, 2);
      assert.strictEqual(result[0].sheetId, 0);
      assert.strictEqual(result[0].sheetName, 'Sheet1');
      assert.strictEqual(result[0].rules.length, 1);
      assert.strictEqual(result[1].sheetId, 1);
      assert.strictEqual(result[1].sheetName, 'Sheet2');
      assert.strictEqual(result[1].rules.length, 0);
    });

    it('should filter by sheetId when provided', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [
                { properties: { sheetId: 0, title: 'Sheet1' }, conditionalFormats: [] },
                {
                  properties: { sheetId: 1, title: 'Sheet2' },
                  conditionalFormats: [{ booleanRule: {} }],
                },
                { properties: { sheetId: 2, title: 'Sheet3' }, conditionalFormats: [] },
              ],
            },
          })),
        },
      };

      const result = await getConditionalFormatRules(mockSheets, 'sheet123', 1);

      assert.strictEqual(result.length, 1);
      assert.strictEqual(result[0].sheetId, 1);
      assert.strictEqual(result[0].sheetName, 'Sheet2');
    });

    it('should handle sheets without conditionalFormats property', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
            },
          })),
        },
      };

      const result = await getConditionalFormatRules(mockSheets, 'sheet123');

      assert.strictEqual(result.length, 1);
      assert.deepStrictEqual(result[0].rules, []);
    });
  });

  describe('deleteConditionalFormatRule', () => {
    it('should delete a conditional format rule by index', async () => {
      const mockSheets = {
        spreadsheets: {
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}] },
          })),
        },
      };

      const result = await deleteConditionalFormatRule(mockSheets, 'sheet123', 0, 2);

      assert.strictEqual(result.spreadsheetId, 'sheet123');
      assert.strictEqual(mockSheets.spreadsheets.batchUpdate.mock.calls.length, 1);

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      assert.deepStrictEqual(requestBody.requests[0].deleteConditionalFormatRule, {
        sheetId: 0,
        index: 2,
      });
    });
  });

  describe('clearConditionalFormatRules', () => {
    it('should delete all conditional format rules from a sheet', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [
                {
                  properties: { sheetId: 0, title: 'Sheet1' },
                  conditionalFormats: [{}, {}, {}], // 3 rules
                },
              ],
            },
          })),
          batchUpdate: mock.fn(async () => ({
            data: { spreadsheetId: 'sheet123', replies: [{}, {}, {}] },
          })),
        },
      };

      const result = await clearConditionalFormatRules(mockSheets, 'sheet123', 0);

      assert.strictEqual(result.spreadsheetId, 'sheet123');
      assert.strictEqual(mockSheets.spreadsheets.batchUpdate.mock.calls.length, 1);

      const requestBody = mockSheets.spreadsheets.batchUpdate.mock.calls[0].arguments[0].requestBody;
      // Should delete in reverse order (2, 1, 0)
      assert.strictEqual(requestBody.requests.length, 3);
      assert.strictEqual(requestBody.requests[0].deleteConditionalFormatRule.index, 2);
      assert.strictEqual(requestBody.requests[1].deleteConditionalFormatRule.index, 1);
      assert.strictEqual(requestBody.requests[2].deleteConditionalFormatRule.index, 0);
    });

    it('should return early if no rules exist', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [
                {
                  properties: { sheetId: 0, title: 'Sheet1' },
                  conditionalFormats: [],
                },
              ],
            },
          })),
          batchUpdate: mock.fn(async () => ({})),
        },
      };

      const result = await clearConditionalFormatRules(mockSheets, 'sheet123', 0);

      assert.strictEqual(result.spreadsheetId, 'sheet123');
      // batchUpdate should NOT be called since there are no rules
      assert.strictEqual(mockSheets.spreadsheets.batchUpdate.mock.calls.length, 0);
    });

    it('should return early if sheet has no rules property', async () => {
      const mockSheets = {
        spreadsheets: {
          get: mock.fn(async () => ({
            data: {
              sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
            },
          })),
          batchUpdate: mock.fn(async () => ({})),
        },
      };

      const result = await clearConditionalFormatRules(mockSheets, 'sheet123', 0);

      assert.strictEqual(result.spreadsheetId, 'sheet123');
      assert.strictEqual(mockSheets.spreadsheets.batchUpdate.mock.calls.length, 0);
    });
  });
});
