// src/googleAppsScriptApiHelpers.ts
import { script_v1 } from 'googleapis';
import { UserError } from 'fastmcp';

type Script = script_v1.Script;

/**
 * Gets the content of an Apps Script project (all files)
 */
export async function getScriptContent(
  script: Script,
  scriptId: string
): Promise<script_v1.Schema$Content> {
  try {
    const response = await script.projects.getContent({
      scriptId: scriptId,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Script not found with ID: ${scriptId}. Make sure the Script ID is correct and you have access to it.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for script: ${scriptId}. Make sure you have edit access to this script.`);
    }
    throw new UserError(`Failed to get script content: ${error.message}`);
  }
}

/**
 * Updates the content of an Apps Script project
 */
export async function updateScriptContent(
  script: Script,
  scriptId: string,
  files: script_v1.Schema$File[]
): Promise<script_v1.Schema$Content> {
  try {
    const response = await script.projects.updateContent({
      scriptId: scriptId,
      requestBody: {
        files: files,
      },
    });
    return response.data;
  } catch (error: any) {
    const errorData = error.response?.data || {};
    const errorDetails = errorData.error || {};
    const errorMessage = errorDetails.message || error.message || 'Unknown error';
    const errorStatus = errorDetails.status || error.status || error.code;
    const errors = error.errors || errorDetails.details || [];

    let fullError = `API Error (${errorStatus}): ${errorMessage}`;
    if (errors.length > 0) {
      fullError += ` | Details: ${JSON.stringify(errors)}`;
    }

    if (error.code === 404 || errorStatus === 'NOT_FOUND') {
      throw new UserError(`Script not found with ID: ${scriptId}. Make sure the Script ID is correct.`);
    }
    if (error.code === 403 || errorStatus === 'PERMISSION_DENIED') {
      throw new UserError(`Permission denied for script: ${scriptId}. ${fullError}.

Possible fixes:
1. In Google Cloud Console, ensure the Apps Script API is enabled
2. In the Apps Script editor, go to Project Settings and enable "Google Apps Script API"
3. Make sure you have Editor (not Viewer) access to the script
4. Try re-authenticating: delete token.json and restart the MCP server`);
    }
    if (error.code === 400 || errorStatus === 'INVALID_ARGUMENT') {
      throw new UserError(`Invalid script content: ${fullError}`);
    }
    throw new UserError(`Failed to update script content: ${fullError}`);
  }
}

/**
 * Gets metadata about an Apps Script project
 */
export async function getScriptMetadata(
  script: Script,
  scriptId: string
): Promise<script_v1.Schema$Project> {
  try {
    const response = await script.projects.get({
      scriptId: scriptId,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Script not found with ID: ${scriptId}.`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied for script: ${scriptId}.`);
    }
    throw new UserError(`Failed to get script metadata: ${error.message}`);
  }
}

/**
 * Updates a single file within a script project
 */
export async function updateScriptFile(
  script: Script,
  scriptId: string,
  fileName: string,
  newSource: string,
  fileType: 'SERVER_JS' | 'HTML' | 'JSON' = 'SERVER_JS'
): Promise<script_v1.Schema$Content> {
  const content = await getScriptContent(script, scriptId);
  const files = content.files || [];

  let fileFound = false;
  const updatedFiles = files.map(file => {
    if (file.name === fileName) {
      fileFound = true;
      return { ...file, source: newSource };
    }
    return file;
  });

  if (!fileFound) {
    updatedFiles.push({
      name: fileName,
      type: fileType,
      source: newSource,
    });
  }

  return updateScriptContent(script, scriptId, updatedFiles);
}

/**
 * Lists all files in a script project with their names and types
 */
export async function listScriptFiles(
  script: Script,
  scriptId: string
): Promise<{ name: string; type: string; source?: string }[]> {
  const content = await getScriptContent(script, scriptId);
  const files = content.files || [];

  return files.map(file => ({
    name: file.name || 'Unknown',
    type: file.type || 'Unknown',
    source: file.source || undefined,
  }));
}

/**
 * Gets a single file's source code from a script project
 */
export async function getScriptFile(
  script: Script,
  scriptId: string,
  fileName: string
): Promise<{ name: string; type: string; source: string } | null> {
  const content = await getScriptContent(script, scriptId);
  const files = content.files || [];

  const file = files.find(f => f.name === fileName);
  if (!file) {
    return null;
  }

  return {
    name: file.name || 'Unknown',
    type: file.type || 'Unknown',
    source: file.source || '',
  };
}
