import type {
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';

export interface McpClient {
  readonly id: string;
  readonly name: string;
  readonly config: McpServerConfig;
  initialize(): Promise<void>;
  listTools(): Promise<McpToolDefinition[]>;
  callTool(
    toolName: string,
    args: Record<string, unknown>,
    options?: { abortSignal?: AbortSignal },
  ): Promise<McpToolResult>;
  close(): Promise<void>;
  getStatus(): McpServerStatus;
}
