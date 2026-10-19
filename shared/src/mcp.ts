import type { ChangeOp } from './changeset.js';

export type McpServerType = 'sqlite' | 'csv-directory' | 'filesystem';
export type McpTransport = 'builtin' | 'stdio';

export interface McpServerConfig {
  id: string;
  name: string;
  enabled: boolean;
  transport: McpTransport;
  type?: McpServerType;
  command?: string;
  args?: string[];
  env?: Record<string, string>;
  allowedPaths: string[];
  options?: Record<string, unknown>;
}

export interface McpServerStatus {
  id: string;
  name: string;
  enabled: boolean;
  transport: McpTransport;
  type?: McpServerType;
  status: 'connected' | 'disconnected' | 'error';
  error?: string;
  toolCount: number;
  tools?: Array<{ name: string; description?: string }>;
}

export interface McpToolDefinition {
  name: string;
  description?: string;
  inputSchema: Record<string, unknown>;
}

export interface McpToolResult {
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
  changeOps?: ChangeOp[];
}
