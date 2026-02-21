import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getAppsScriptClient } from '../../clients.js';
import * as AppsScriptHelpers from '../../googleAppsScriptApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'updateAppsScriptFile',
    description: 'Updates the source code of a specific file in a Google Apps Script project.',
    parameters: z.object({
      scriptId: z
        .string()
        .describe('The Script ID (found in Project Settings in the Apps Script editor).'),
      fileName: z
        .string()
        .describe(
          'The name of the file to update (without extension, e.g., "Code" not "Code.gs").'
        ),
      source: z.string().describe('The new source code for the file.'),
      fileType: z
        .enum(['SERVER_JS', 'HTML', 'JSON'])
        .optional()
        .default('SERVER_JS')
        .describe(
          'The file type (SERVER_JS for .gs files, HTML for .html files, JSON for appsscript.json).'
        ),
    }),
    execute: async (args, { log }) => {
      const script = await getAppsScriptClient();
      log.info(`Updating file "${args.fileName}" in Apps Script project: ${args.scriptId}`);

      try {
        await AppsScriptHelpers.updateScriptFile(
          script,
          args.scriptId,
          args.fileName,
          args.source,
          args.fileType
        );

        return `Successfully updated file "${args.fileName}" in Apps Script project.`;
      } catch (error: any) {
        log.error(`Error updating script file: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to update script file: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
