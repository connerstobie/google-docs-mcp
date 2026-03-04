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

// Fix for 100% CPU on parent death: the @modelcontextprotocol/sdk
// StdioServerTransport listens for 'data' and 'error' on stdin but NOT
// 'end' or 'close'. When the parent dies and the stdin pipe breaks, libuv
// keeps polling the dead fd at full speed → 100% CPU spin.
// Note: on macOS, stdin 'end'/'close' events are unreliable (sometimes they
// never fire), so the parent PID watchdog below is the reliable fallback.
process.stdin.on('end', () => { logger.info('stdin ended. Exiting.'); process.exit(0); });
process.stdin.on('close', () => { logger.info('stdin closed. Exiting.'); process.exit(0); });

// Fallback: poll parent PID every 2 seconds. Handles the macOS case where
// stdin events don't fire. See: https://github.com/anthropics/claude-code/issues/1935
const parentPid = process.ppid;
setInterval(() => {
  try {
    process.kill(parentPid, 0);
  } catch {
    logger.info(`Parent process ${parentPid} gone. Exiting.`);
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
