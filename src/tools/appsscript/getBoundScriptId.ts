import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getDriveClient } from '../../clients.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'getBoundScriptId',
    description:
      'Finds the Apps Script project ID bound to a Google Spreadsheet (or Doc). Bound scripts are container-bound projects attached to a specific file.',
    parameters: z.object({
      fileId: z
        .string()
        .describe(
          'The ID of the Google Spreadsheet or Document that has a bound Apps Script project.'
        ),
    }),
    execute: async (args, { log }) => {
      const drive = await getDriveClient();
      log.info(`Looking for bound Apps Script project on file: ${args.fileId}`);

      try {
        const response = await drive.files.list({
          q: `'${args.fileId}' in parents and mimeType='application/vnd.google-apps.script' and trashed=false`,
          fields: 'files(id,name,createdTime,modifiedTime)',
          supportsAllDrives: true,
          includeItemsFromAllDrives: true,
        });

        const files = response.data.files || [];

        if (files.length === 0) {
          return `No bound Apps Script project found for file ${args.fileId}. The file may not have any scripts attached, or you may not have access.`;
        }

        const scripts = files.map((f) => ({
          scriptId: f.id,
          name: f.name,
          createdTime: f.createdTime,
          modifiedTime: f.modifiedTime,
        }));

        return JSON.stringify({ scripts }, null, 2);
      } catch (error: any) {
        log.error(`Error finding bound script: ${error.message || error}`);
        if (error.code === 404) {
          throw new UserError(`File not found (ID: ${args.fileId}). Check the ID.`);
        }
        if (error.code === 403) {
          throw new UserError(
            `Permission denied for file (ID: ${args.fileId}). Ensure you have access.`
          );
        }
        throw new UserError(`Failed to find bound script: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
