import type { ChangePreviewItem } from '@excelai/shared';

export type ActivityStep =
  | { kind: 'thinking'; text: string; durationMs?: number }
  | { kind: 'tool'; text: string };

export type ChatItem =
  | { kind: 'user'; text: string }
  | { kind: 'assistant'; text: string; streaming?: boolean }
  | { kind: 'tool'; text: string }
  | { kind: 'thinking'; text: string; durationMs?: number; active?: boolean }
  | {
      kind: 'activity';
      steps: ActivityStep[];
      totalDurationMs: number;
      active: boolean;
    }
  | { kind: 'error'; text: string }
  | { kind: 'stopped'; text: string }
  | {
      kind: 'changeset';
      id: string;
      preview: ChangePreviewItem[];
      status: 'staged' | 'applied' | 'rejected' | 'undone';
    };

export type ProviderKind = 'anthropic' | 'openai' | 'openai-compatible' | 'embedded';

export interface SavedProvider {
  id: string;
  kind: ProviderKind;
  model: string;
  baseUrl?: string;
  /** masked, presence means a key is stored */
  apiKey?: string;
  modelPath?: string;
}

export interface ProviderList {
  providers: SavedProvider[];
  activeProviderId: string | null;
}

export const KIND_LABELS: Record<ProviderKind, string> = {
  anthropic: 'Anthropic',
  openai: 'OpenAI',
  'openai-compatible': 'OpenAI-compatible (Ollama, LM Studio, …)',
  embedded: 'Embedded GGUF (In-Process node-llama-cpp)',
};
