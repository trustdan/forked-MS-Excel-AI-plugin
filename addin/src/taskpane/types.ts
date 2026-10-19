import type { ChangePreviewItem } from '@excelai/shared';

export type ChatItem =
  | { kind: 'user'; text: string }
  | { kind: 'assistant'; text: string }
  | { kind: 'tool'; text: string }
  | { kind: 'error'; text: string }
  | { kind: 'stopped'; text: string }
  | {
      kind: 'changeset';
      id: string;
      preview: ChangePreviewItem[];
      status: 'staged' | 'applied' | 'rejected' | 'undone';
    };

export type ProviderKind = 'anthropic' | 'openai' | 'openai-compatible';

export interface SavedProvider {
  id: string;
  kind: ProviderKind;
  model: string;
  baseUrl?: string;
  /** masked, presence means a key is stored */
  apiKey?: string;
}

export interface ProviderList {
  providers: SavedProvider[];
  activeProviderId: string | null;
}

export const KIND_LABELS: Record<ProviderKind, string> = {
  anthropic: 'Anthropic',
  openai: 'OpenAI',
  'openai-compatible': 'OpenAI-compatible (Ollama, LM Studio, …)',
};
