import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getAppsScriptClient } from '../../clients.js';
import * as AppsScriptHelpers from '../../googleAppsScriptApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'readAppsScriptFile',
    description: 'Reads the source code of a specific file in a Google Apps Script project.',
    parameters: z.object({
      scriptId: z
        .string()
        .describe('The Script ID (found in Project Settings in the Apps Script editor).'),
      fileName: z
        .string()
        .describe(
          'The name of the file to read (without extension, e.g., "Code" not "Code.gs").'
        ),
    }),
    execute: async (args, { log }) => {
      const script = await getAppsScriptClient();
      log.info(`Reading file "${args.fileName}" from Apps Script project: ${args.scriptId}`);

      try {
        const file = await AppsScriptHelpers.getScriptFile(script, args.scriptId, args.fileName);

        if (!file) {
          throw new UserError(`File "${args.fileName}" not found in script project.`);
        }

        return `**File: ${file.name}** (${file.type})\n\n\`\`\`javascript\n${file.source}\n\`\`\``;
      } catch (error: any) {
        log.error(`Error reading script file: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to read script file: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
