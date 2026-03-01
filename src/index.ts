#!/usr/bin/env node

// src/index.ts
//
// Single entry point for the Google Docs MCP Server.
//
// Usage:
//   google-docs-mcp          Start the MCP server (default)
//   google-docs-mcp auth     Run the interactive OAuth flow

import { FastMCP } from 'fastmcp';
import { initializeGoogleClient } from './clients.js';
import { registerAllTools } from './tools/index.js';
import { logger } from './logger.js';

// --- Auth subcommand ---
if (process.argv[2] === 'auth') {
  const { runAuthFlow } = await import('./auth.js');
  try {
    await runAuthFlow();
    logger.info('Authorization complete. You can now start the MCP server.');
    process.exit(0);
  } catch (error: any) {
    logger.error('Authorization failed:', error.message || error);
    process.exit(1);
  }
}

// --- Server startup ---

// Exit cleanly when the parent process (Claude/VSCode) dies.
// The stdin 'end'/'close' approach is unreliable on macOS — events often don't
// fire when the parent is killed. Instead, poll the parent PID directly.
// See: https://github.com/anthropics/claude-code/issues/1935
const parentPid = process.ppid;
const parentCheckInterval = setInterval(() => {
  try {
    process.kill(parentPid, 0); // signal 0 = existence check, no actual signal sent
  } catch {
    logger.info(`Parent process ${parentPid} is gone. Exiting.`);
    clearInterval(parentCheckInterval);
    process.exit(0);
  }
}, 2000);

// Set up process-level unhandled error/rejection handlers to prevent crashes
process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception:', error);
});

process.on('unhandledRejection', (reason, _promise) => {
  logger.error('Unhandled Promise Rejection:', reason);
});

const server = new FastMCP({
  name: 'Ultimate Google Docs & Sheets MCP Server',
  version: '1.0.0',
});

registerAllTools(server);

try {
  await initializeGoogleClient();
  logger.info('Starting Ultimate Google Docs & Sheets MCP server...');

  server.start({ transportType: 'stdio' as const });
  logger.info('MCP Server running using stdio. Awaiting client connection...');
  logger.info('Process-level error handling configured to prevent crashes from timeout errors.');
} catch (startError: any) {
  logger.error('FATAL: Server failed to start:', startError.message || startError);
  process.exit(1);
}
