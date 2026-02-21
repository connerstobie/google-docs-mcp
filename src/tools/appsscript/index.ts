import type { FastMCP } from 'fastmcp';
import { register as getBoundScriptId } from './getBoundScriptId.js';
import { register as listAppsScriptFiles } from './listAppsScriptFiles.js';
import { register as readAppsScriptFile } from './readAppsScriptFile.js';
import { register as updateAppsScriptFile } from './updateAppsScriptFile.js';
import { register as getAppsScriptMetadata } from './getAppsScriptMetadata.js';

export function registerAppsScriptTools(server: FastMCP) {
  getBoundScriptId(server);
  listAppsScriptFiles(server);
  readAppsScriptFile(server);
  updateAppsScriptFile(server);
  getAppsScriptMetadata(server);
}
