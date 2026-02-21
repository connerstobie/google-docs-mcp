import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getAppsScriptClient } from '../../clients.js';
import * as AppsScriptHelpers from '../../googleAppsScriptApiHelpers.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'getAppsScriptMetadata',
    description:
      'Gets metadata about a Google Apps Script project (title, create time, update time, etc.).',
    parameters: z.object({
      scriptId: z
        .string()
        .describe('The Script ID (found in Project Settings in the Apps Script editor).'),
    }),
    execute: async (args, { log }) => {
      const script = await getAppsScriptClient();
      log.info(`Getting metadata for Apps Script project: ${args.scriptId}`);

      try {
        const metadata = await AppsScriptHelpers.getScriptMetadata(script, args.scriptId);

        return `**Apps Script Project Metadata:**
- **Title:** ${metadata.title || 'Unknown'}
- **Script ID:** ${metadata.scriptId || args.scriptId}
- **Parent ID:** ${metadata.parentId || 'None (standalone script)'}
- **Create Time:** ${metadata.createTime || 'Unknown'}
- **Update Time:** ${metadata.updateTime || 'Unknown'}`;
      } catch (error: any) {
        log.error(`Error getting script metadata: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to get script metadata: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
