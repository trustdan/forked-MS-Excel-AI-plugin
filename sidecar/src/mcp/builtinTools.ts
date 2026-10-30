import type { z } from 'zod';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { McpToolDefinition } from '@excelai/shared';

/**
 * A builtin connector tool. Its zod schema is both the JSON schema the model
 * sees and the runtime check on what the model sends, so the two cannot drift.
 */
export interface BuiltinTool<S extends z.ZodTypeAny = z.ZodTypeAny> {
  description: string;
  parameters: S;
}

export function toolDefinitions(tools: Record<string, BuiltinTool>): McpToolDefinition[] {
  return Object.entries(tools).map(([name, tool]) => ({
    name,
    description: tool.description,
    inputSchema: zodToJsonSchema(tool.parameters, { $refStrategy: 'none' }) as Record<string, unknown>,
  }));
}

/** Validates model-supplied arguments; malformed values are rejected, never coerced. */
export function parseArgs<S extends z.ZodTypeAny>(tool: BuiltinTool<S>, args: unknown): z.infer<S> {
  const parsed = tool.parameters.safeParse(args);
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`)
      .join('; ');
    throw new Error(`Invalid arguments: ${issues}`);
  }
  return parsed.data;
}
