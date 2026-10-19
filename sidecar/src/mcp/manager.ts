import type {
  ChangeOp,
  McpServerConfig,
  McpServerStatus,
  McpToolDefinition,
  McpToolResult,
} from '@excelai/shared';
import type { ToolSpec } from '../providers/types.js';
import type { McpClient } from './types.js';
import { createMcpClient } from './client.js';

export interface DispatchedToolInfo {
  client: McpClient;
  originalName: string;
  definition: McpToolDefinition;
}

export class McpManager {
  private clients = new Map<string, McpClient>();
  private toolMap = new Map<string, DispatchedToolInfo>();

  /**
   * Synchronize active MCP clients with the given server configurations and global approved paths.
   */
  async sync(configs: McpServerConfig[], globalApprovedPaths: string[] = []): Promise<void> {
    const desiredIds = new Set(configs.map((c) => c.id));

    // Close clients no longer present
    for (const [id, client] of this.clients.entries()) {
      if (!desiredIds.has(id)) {
        await client.close().catch(() => {});
        this.clients.delete(id);
      }
    }

    // Initialize or update clients
    for (const config of configs) {
      if (!config.enabled) {
        if (this.clients.has(config.id)) {
          await this.clients.get(config.id)!.close().catch(() => {});
          this.clients.delete(config.id);
        }
        continue;
      }

      // Merge global approved paths into server allowed paths
      const effectiveAllowedPaths = Array.from(
        new Set([...(config.allowedPaths ?? []), ...globalApprovedPaths]),
      );
      const effectiveConfig: McpServerConfig = {
        ...config,
        allowedPaths: effectiveAllowedPaths,
      };

      // If client exists and config changed, restart it
      const existing = this.clients.get(config.id);
      if (existing) {
        if (JSON.stringify(existing.config) !== JSON.stringify(effectiveConfig)) {
          await existing.close().catch(() => {});
          this.clients.delete(config.id);
        } else {
          continue; // unchanged
        }
      }

      try {
        const client = createMcpClient(effectiveConfig);
        await client.initialize();
        this.clients.set(config.id, client);
      } catch (err) {
        console.error(`[mcp] failed to initialize client "${config.name}":`, err);
      }
    }

    await this.refreshToolMap();
  }

  /**
   * Rebuild the tool mapping across all active and connected MCP clients.
   */
  async refreshToolMap(): Promise<void> {
    this.toolMap.clear();

    for (const client of this.clients.values()) {
      try {
        const tools = await client.listTools();
        for (const t of tools) {
          // If tool name is already taken, prefix with server id to prevent collisions
          let registeredName = t.name;
          if (this.toolMap.has(registeredName)) {
            registeredName = `mcp__${client.id}__${t.name}`;
          }
          this.toolMap.set(registeredName, {
            client,
            originalName: t.name,
            definition: t,
          });
        }
      } catch (err) {
        console.error(`[mcp] failed to list tools for client "${client.name}":`, err);
      }
    }
  }

  /**
   * Returns tool specs compatible with LLM providers.
   */
  getToolSpecs(): ToolSpec[] {
    const specs: ToolSpec[] = [];
    for (const [name, info] of this.toolMap.entries()) {
      specs.push({
        name,
        description: `[MCP: ${info.client.name}] ${info.definition.description ?? ''}`,
        inputSchema: info.definition.inputSchema,
      });
    }
    return specs;
  }

  /**
   * Check if a tool name belongs to a registered MCP connector.
   */
  hasTool(toolName: string): boolean {
    return this.toolMap.has(toolName);
  }

  /**
   * Execute an MCP tool with prompt injection defenses and untrusted data tagging.
   */
  async callTool(
    toolName: string,
    args: Record<string, unknown>,
    options?: { abortSignal?: AbortSignal },
  ): Promise<{ resultText: string; isError?: boolean; changeOps?: ChangeOp[] }> {
    const info = this.toolMap.get(toolName);
    if (!info) {
      return {
        resultText: `Error: MCP tool "${toolName}" is not registered.`,
        isError: true,
      };
    }

    try {
      const res: McpToolResult = await info.client.callTool(info.originalName, args, options);
      const rawText = res.content.map((c) => c.text).join('\n');

      // Wrap in untrusted data marker to protect against prompt injection (CWE-1236 & prompt injection defense)
      const sanitizedText =
        `[UNTRUSTED EXTERNAL DATA FROM MCP SERVER: ${info.client.name}]\n` +
        `${rawText}\n` +
        `[/UNTRUSTED EXTERNAL DATA]`;

      return {
        resultText: sanitizedText,
        isError: res.isError,
        changeOps: res.changeOps,
      };
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      return {
        resultText: `MCP execution error in "${toolName}": ${msg}`,
        isError: true,
      };
    }
  }

  /**
   * Probe / test an MCP server configuration without saving it permanently.
   */
  async testServer(
    config: McpServerConfig,
    globalApprovedPaths: string[] = [],
  ): Promise<{ ok: boolean; tools: McpToolDefinition[]; error?: string }> {
    const effectiveAllowed = Array.from(
      new Set([...(config.allowedPaths ?? []), ...globalApprovedPaths]),
    );
    const testConfig: McpServerConfig = {
      ...config,
      allowedPaths: effectiveAllowed,
    };

    let client: McpClient | null = null;
    try {
      client = createMcpClient(testConfig);
      await client.initialize();
      const tools = await client.listTools();
      return { ok: true, tools };
    } catch (err) {
      return {
        ok: false,
        tools: [],
        error: err instanceof Error ? err.message : String(err),
      };
    } finally {
      if (client) {
        await client.close().catch(() => {});
      }
    }
  }

  /**
   * Get statuses for all configured servers.
   */
  getServerStatuses(configuredServers: McpServerConfig[]): McpServerStatus[] {
    return configuredServers.map((cfg) => {
      const active = this.clients.get(cfg.id);
      if (active) {
        return active.getStatus();
      }
      return {
        id: cfg.id,
        name: cfg.name,
        enabled: cfg.enabled,
        transport: cfg.transport,
        type: cfg.type,
        status: cfg.enabled ? 'disconnected' : 'disconnected',
        toolCount: 0,
      };
    });
  }

  /**
   * Close all active MCP connections on shutdown.
   */
  async closeAll(): Promise<void> {
    for (const client of this.clients.values()) {
      try {
        await client.close();
      } catch {}
    }
    this.clients.clear();
    this.toolMap.clear();
  }
}
