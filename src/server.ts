// src/server.ts
import { FastMCP } from 'fastmcp';
import { initializeGoogleClient } from './clients.js';
import { registerAllTools } from './tools/index.js';
import { logger } from './logger.js';

// Set up process-level unhandled error/rejection handlers to prevent crashes
process.on('uncaughtException', (error) => {
  logger.error('Uncaught Exception:', error);
  // Don't exit process, just log the error and continue
  // This will catch timeout errors that might otherwise crash the server
});

process.on('unhandledRejection', (reason, promise) => {
  logger.error('Unhandled Promise Rejection:', reason);
  // Don't exit process, just log the error and continue
});

const server = new FastMCP({
  name: 'Ultimate Google Docs & Sheets MCP Server',
  version: '1.0.0',
});

// Register all tools from individual modules
registerAllTools(server);

// --- Server Startup ---
async function startServer() {
  try {
    await initializeGoogleClient(); // Authorize BEFORE starting listeners
    logger.info('Starting Ultimate Google Docs & Sheets MCP server...');

    // Using stdio as before
    const configToUse = {
      transportType: 'stdio' as const,
    };

    // Start the server with proper error handling
    server.start(configToUse);
    logger.info(
      `MCP Server running using ${configToUse.transportType}. Awaiting client connection...`
    );

    logger.info(
      'Process-level error handling configured to prevent crashes from timeout errors.'
    );
  } catch (startError: any) {
    logger.error('FATAL: Server failed to start:', startError.message || startError);
    process.exit(1);
  }
}

startServer(); // Removed .catch here, let errors propagate if startup fails critically
