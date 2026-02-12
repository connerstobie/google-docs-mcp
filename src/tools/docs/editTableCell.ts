import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import type { docs_v1 } from 'googleapis';
import { getDocsClient } from '../../clients.js';
import * as GDocsHelpers from '../../googleDocsApiHelpers.js';
import {
  DocumentIdParameter,
  TextStyleParameters,
  ParagraphStyleParameters,
  TextStyleArgs,
  ParagraphStyleArgs,
} from '../../types.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'editTableCell',
    description:
      'Edits the content and/or basic style of a specific table cell. Requires knowing table start index.',
    parameters: DocumentIdParameter.extend({
      tableStartIndex: z
        .number()
        .int()
        .min(1)
        .describe(
          'The starting index of the TABLE element itself (tricky to find, may require reading structure first).'
        ),
      rowIndex: z.number().int().min(0).describe('Row index (0-based).'),
      columnIndex: z.number().int().min(0).describe('Column index (0-based).'),
      textContent: z
        .string()
        .optional()
        .describe('Optional: New text content for the cell. Replaces existing content.'),
      textStyle: TextStyleParameters.optional().describe('Optional: Text styles to apply.'),
      paragraphStyle: ParagraphStyleParameters.optional().describe(
        'Optional: Paragraph styles (like alignment) to apply.'
      ),
      tabId: z.string().optional().describe('Optional: Target a specific document tab.'),
    }),
    execute: async (args, { log }) => {
      const docs = await getDocsClient();
      log.info(
        `Editing cell (${args.rowIndex}, ${args.columnIndex}) in table starting at ${args.tableStartIndex}, doc ${args.documentId}`
      );

      try {
        const cellRange = await GDocsHelpers.getTableCellRange(
          docs,
          args.documentId,
          args.tableStartIndex,
          args.rowIndex,
          args.columnIndex,
          args.tabId
        );
        log.info(`Cell content range: ${cellRange.startIndex}-${cellRange.endIndex}`);

        const requests: docs_v1.Schema$Request[] = [];
        let newTextStart = cellRange.startIndex;
        let newTextEnd = cellRange.startIndex;

        if (args.textContent !== undefined) {
          if (cellRange.endIndex > cellRange.startIndex) {
            requests.push({
              deleteContentRange: {
                range: {
                  startIndex: cellRange.startIndex,
                  endIndex: cellRange.endIndex,
                },
              },
            });
          }
          if (args.textContent.length > 0) {
            requests.push({
              insertText: {
                location: { index: cellRange.startIndex },
                text: args.textContent,
              },
            });
          }
          newTextEnd = cellRange.startIndex + args.textContent.length;
        } else {
          newTextEnd = cellRange.endIndex;
        }

        if (args.textStyle && newTextEnd > newTextStart) {
          const styleResult = GDocsHelpers.buildUpdateTextStyleRequest(
            newTextStart,
            newTextEnd,
            args.textStyle as TextStyleArgs,
            args.tabId
          );
          if (styleResult) {
            requests.push(styleResult.request);
          }
        }

        if (args.paragraphStyle && newTextEnd >= newTextStart) {
          const paraEnd = args.textContent !== undefined ? newTextEnd + 1 : cellRange.endIndex + 1;
          const paraResult = GDocsHelpers.buildUpdateParagraphStyleRequest(
            newTextStart,
            paraEnd,
            args.paragraphStyle as ParagraphStyleArgs,
            args.tabId
          );
          if (paraResult) {
            requests.push(paraResult.request);
          }
        }

        if (requests.length === 0) {
          return `No changes specified for cell (${args.rowIndex}, ${args.columnIndex}). Provide textContent, textStyle, or paragraphStyle.`;
        }

        await GDocsHelpers.executeBatchUpdateWithSplitting(docs, args.documentId, requests, log);

        const actions: string[] = [];
        if (args.textContent !== undefined) actions.push(`text set to "${args.textContent}"`);
        if (args.textStyle) actions.push('text style applied');
        if (args.paragraphStyle) actions.push('paragraph style applied');
        return `Successfully edited cell (${args.rowIndex}, ${args.columnIndex}): ${actions.join(', ')}.`;
      } catch (error: unknown) {
        const err = error as { message?: string };
        log.error(`Error editing table cell: ${err?.message ?? error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to edit table cell: ${err?.message ?? 'Unknown error'}`);
      }
    },
  });
}
