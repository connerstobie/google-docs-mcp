import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { getDriveClient, getAppsScriptClient } from '../../clients.js';

export function register(server: FastMCP) {
  server.addTool({
    name: 'getBoundScriptId',
    description:
      'Finds all Apps Script projects bound to a Google Spreadsheet (or Doc). Also works without a fileId to list all accessible Apps Script projects.',
    parameters: z.object({
      fileId: z
        .string()
        .optional()
        .describe(
          'The ID of the Google Spreadsheet or Document to find bound scripts for. If omitted, lists all accessible Apps Script projects.'
        ),
    }),
    execute: async (args, { log }) => {
      const drive = await getDriveClient();
      const script = await getAppsScriptClient();

      log.info(
        args.fileId
          ? `Looking for bound Apps Script projects on file: ${args.fileId}`
          : 'Listing all accessible Apps Script projects'
      );

      try {
        // Search Drive for all Apps Script projects the user can access
        const allScripts: { id: string; name: string; createdTime: string; modifiedTime: string }[] = [];
        let pageToken: string | undefined;

        do {
          const response = await drive.files.list({
            q: `mimeType='application/vnd.google-apps.script' and trashed=false`,
            fields: 'nextPageToken,files(id,name,createdTime,modifiedTime)',
            pageSize: 100,
            supportsAllDrives: true,
            includeItemsFromAllDrives: true,
            pageToken,
          });

          for (const f of response.data.files || []) {
            if (f.id && f.name) {
              allScripts.push({
                id: f.id,
                name: f.name,
                createdTime: f.createdTime || '',
                modifiedTime: f.modifiedTime || '',
              });
            }
          }
          pageToken = response.data.nextPageToken || undefined;
        } while (pageToken);

        log.info(`Found ${allScripts.length} total Apps Script projects`);

        // If no fileId specified, return all scripts
        if (!args.fileId) {
          if (allScripts.length === 0) {
            return 'No Apps Script projects found. You may not have access to any.';
          }
          return JSON.stringify({ scripts: allScripts }, null, 2);
        }

        // For each script, check if it's bound to the target file using the Apps Script API
        const boundScripts: { scriptId: string; name: string; parentId?: string; createdTime: string; modifiedTime: string }[] = [];

        for (const s of allScripts) {
          try {
            const meta = await script.projects.get({ scriptId: s.id });
            const parentId = meta.data.parentId;

            if (parentId === args.fileId) {
              boundScripts.push({
                scriptId: s.id,
                name: s.name,
                parentId: parentId || undefined,
                createdTime: s.createdTime,
                modifiedTime: s.modifiedTime,
              });
            }
          } catch {
            // Skip scripts we can't read metadata for (permission issues)
            continue;
          }
        }

        if (boundScripts.length === 0) {
          return `No bound Apps Script projects found for file ${args.fileId}. Found ${allScripts.length} total scripts but none are bound to this file.`;
        }

        return JSON.stringify({ scripts: boundScripts }, null, 2);
      } catch (error: any) {
        log.error(`Error finding scripts: ${error.message || error}`);
        if (error.code === 404) {
          throw new UserError(`File not found (ID: ${args.fileId}). Check the ID.`);
        }
        if (error.code === 403) {
          throw new UserError(
            `Permission denied for file (ID: ${args.fileId}). Ensure you have access.`
          );
        }
        throw new UserError(`Failed to find scripts: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
