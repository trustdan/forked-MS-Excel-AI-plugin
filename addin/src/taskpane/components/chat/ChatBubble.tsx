import { Markdown } from '../../Markdown.js';
import type { ChatItem } from '../../types.js';
import { ChangeSetCard } from '../changeset/ChangeSetCard.js';
import { ThinkingBlock } from './ThinkingBlock.js';

export interface ChatBubbleProps {
  item: ChatItem;
  onApprove: (id: string) => void;
  onReject: (id: string) => void;
  onUndo: (id: string) => void;
}

export function ChatBubble({
  item,
  onApprove,
  onReject,
  onUndo,
}: ChatBubbleProps): JSX.Element {
  switch (item.kind) {
    case 'user':
      return <div className="bubble user">{item.text}</div>;
    case 'assistant':
      return (
        <div className="bubble assistant">
          <Markdown content={item.text} />
        </div>
      );
    case 'thinking':
      return (
        <ThinkingBlock
          steps={[{ kind: 'thinking', text: item.text, durationMs: item.durationMs }]}
          totalDurationMs={item.durationMs ?? 0}
          active={item.active ?? false}
        />
      );
    case 'activity':
      return (
        <ThinkingBlock
          steps={item.steps}
          totalDurationMs={item.totalDurationMs}
          active={item.active}
        />
      );
    case 'tool':
      return <div className="toolnote">🔧 {item.text}</div>;
    case 'error':
      return <div className="bubble error">{item.text}</div>;
    case 'stopped':
      return (
        <div className="bubble stopped">
          <span className="stopped-icon">■</span>
          <span>{item.text}</span>
        </div>
      );
    case 'changeset':
      return (
        <ChangeSetCard
          id={item.id}
          preview={item.preview}
          status={item.status}
          onApprove={onApprove}
          onReject={onReject}
          onUndo={onUndo}
        />
      );
  }
}
