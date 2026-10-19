import type { SelectionInfo } from '../../../officeBridge.js';

export interface ComposerProps {
  input: string;
  setInput: (value: string | ((prev: string) => string)) => void;
  busy: boolean;
  stopping: boolean;
  status: 'connecting' | 'connected' | 'disconnected' | 'error';
  selection: SelectionInfo | null;
  attachSelection: boolean;
  setAttachSelection: (value: boolean | ((prev: boolean) => boolean)) => void;
  onSendMessage: () => void;
  onStopGeneration: () => void;
}

export function Composer({
  input,
  setInput,
  busy,
  stopping,
  status,
  selection,
  attachSelection,
  setAttachSelection,
  onSendMessage,
  onStopGeneration,
}: ComposerProps): JSX.Element {
  return (
    <footer>
      {selection && (
        <div className="selection-bar">
          <button
            type="button"
            className={`selection-chip ${attachSelection ? 'attached' : ''}`}
            onClick={() => setAttachSelection((prev) => !prev)}
            title={
              attachSelection
                ? 'Selection attached. Click to detach from next message.'
                : 'Click to attach this selection to your next message.'
            }
          >
            <span className="selection-icon">{attachSelection ? '✓' : '⌗'}</span>
            <span className="selection-label">
              {selection.address} ({selection.cellCount} cell{selection.cellCount === 1 ? '' : 's'})
            </span>
            <span className="selection-action">{attachSelection ? 'Attached' : '+ Attach'}</span>
          </button>
          <button
            type="button"
            className="ghost selection-insert-btn"
            title="Insert address into text"
            onClick={() => {
              setInput((prev) => (prev ? `${prev} ${selection.address}` : selection.address));
            }}
          >
            Insert
          </button>
        </div>
      )}
      <div className="composer">
        <textarea
          value={input}
          rows={1}
          placeholder="Ask about this sheet… (Shift+Enter for a new line)"
          aria-label="Prompt input"
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              onSendMessage();
            } else if (e.key === 'Escape' && busy) {
              e.preventDefault();
              onStopGeneration();
            }
          }}
        />
        {busy ? (
          <button
            type="button"
            className={`stop-btn ${stopping ? 'stopping' : ''}`}
            aria-label={stopping ? 'Stopping…' : 'Stop generation'}
            title={stopping ? 'Stopping…' : 'Stop generation (Esc)'}
            onClick={onStopGeneration}
            disabled={stopping}
          >
            <span className="stop-square" />
          </button>
        ) : (
          <button
            type="button"
            className="send"
            aria-label="Send"
            onClick={onSendMessage}
            disabled={busy || status !== 'connected' || !input.trim()}
          >
            ↑
          </button>
        )}
      </div>
    </footer>
  );
}
