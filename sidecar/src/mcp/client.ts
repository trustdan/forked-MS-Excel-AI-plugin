import type { McpServerConfig } from '@excelai/shared';
import type { McpClient } from './types.js';
import { SqliteConnector } from './connectors/sqliteConnector.js';
import { CsvDirectoryConnector } from './connectors/csvDirectoryConnector.js';
import { FilesystemConnector } from './connectors/filesystemConnector.js';
import { StdioMcpClient } from './transports/stdioTransport.js';

export function createMcpClient(config: McpServerConfig): McpClient {
  if (config.transport === 'stdio') {
    return new StdioMcpClient(config);
  }

  if (config.transport === 'builtin') {
    switch (config.type) {
      case 'sqlite':
        return new SqliteConnector(config);
      case 'csv-directory':
        return new CsvDirectoryConnector(config);
      case 'filesystem':
        return new FilesystemConnector(config);
      default:
        throw new Error(`Unsupported builtin MCP server type: ${String(config.type)}`);
    }
  }

  throw new Error(`Unsupported MCP transport: ${String(config.transport)}`);
}
