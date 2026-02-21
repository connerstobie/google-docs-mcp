import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getAppsScriptClient } from '../../clients.js';
import * as AppsScriptHelpers from '../../googleAppsScriptApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'listAppsScriptFiles',
    description: 'Lists all files in a Google Apps Script project.',
    parameters: z.object({
      scriptId: z
        .string()
        .describe('The Script ID (found in Project Settings in the Apps Script editor).'),
    }),
    execute: async (args, { log }) => {
      const script = await getAppsScriptClient();
      log.info(`Listing files in Apps Script project: ${args.scriptId}`);

      try {
        const files = await AppsScriptHelpers.listScriptFiles(script, args.scriptId);

        const fileList = files.map((f) => `- ${f.name} (${f.type})`).join('\n');
        return `**Apps Script Project Files:**\n${fileList}`;
      } catch (error: any) {
        log.error(`Error listing script files: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to list script files: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
