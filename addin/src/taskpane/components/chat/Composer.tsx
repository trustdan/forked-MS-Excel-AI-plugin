import { useRef } from 'react';
import type { SelectionInfo } from '../../../officeBridge.js';

export interface AttachedFile {
  name: string;
  size: number;
  content: string;
}

export interface ComposerProps {
  input: string;
  setInput: (value: string | ((prev: string) => string)) => void;
  busy: boolean;
  stopping: boolean;
  status: 'connecting' | 'connected' | 'disconnected' | 'error';
  selection: SelectionInfo | null;
  attachSelection: boolean;
  setAttachSelection: (value: boolean | ((prev: boolean) => boolean)) => void;
  attachedFile?: AttachedFile | null;
  setAttachedFile?: (file: AttachedFile | null | ((prev: AttachedFile | null) => AttachedFile | null)) => void;
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
  attachedFile,
  setAttachedFile,
  onSendMessage,
  onStopGeneration,
}: ComposerProps): JSX.Element {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file || !setAttachedFile) return;
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result ?? '');
      setAttachedFile({
        name: file.name,
        size: file.size,
        content: text,
      });
      if (!input.trim()) {
        setInput(`Import ${file.name} starting at A1`);
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  return (
    <footer>
      <div className="selection-bar">
        {selection && (
          <>
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
          </>
        )}

        {attachedFile ? (
          <div className="file-chip" title={`${attachedFile.name} (${Math.round(attachedFile.size / 1024) || 1} KB)`}>
            <span className="file-icon">📄</span>
            <span className="file-name">{attachedFile.name}</span>
            <button
              type="button"
              className="chip-remove-btn"
              title="Detach file"
              aria-label="Detach file"
              onClick={() => setAttachedFile?.(null)}
            >
              ✕
            </button>
          </div>
        ) : (
          setAttachedFile && (
            <button
              type="button"
              className="ghost file-attach-btn"
              title="Attach local CSV, TSV, Markdown, or JSON file"
              onClick={() => fileInputRef.current?.click()}
            >
              📎 Attach File
            </button>
          )
        )}
        <input
          ref={fileInputRef}
          type="file"
          accept=".csv,.tsv,.txt,.md,.markdown,.json"
          style={{ display: 'none' }}
          data-testid="file-upload-input"
          onChange={handleFileChange}
        />
      </div>
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
