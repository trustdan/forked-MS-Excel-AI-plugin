import { useCallback, useEffect, useRef, useState } from 'react';
import {
  api,
  checkAuth,
  clearToken,
  connect,
  getToken,
  type AgentEventPayload,
  type SidecarConnection,
} from '../connection.js';
import {
  getWorkbookIdentity,
  getSelectedRangeInfo,
  registerSelectionChangeListener,
  type SelectionInfo,
} from '../officeBridge.js';
import {
  type ThemePreference,
  applyTheme,
  getStoredThemePreference,
  resolveIsDark,
  setStoredThemePreference,
} from './theme.js';
import type { ChatItem, ProviderList } from './types.js';
import { PairingScreen } from './components/pairing/PairingScreen.js';
import { ModelsScreen } from './components/settings/ModelsScreen.js';
import { ChatBubble } from './components/chat/ChatBubble.js';
import { Composer } from './components/chat/Composer.js';

export function App(): JSX.Element {
  const [paired, setPaired] = useState<boolean>(() => getToken() !== null);
  const [status, setStatus] = useState<'connecting' | 'connected' | 'disconnected' | 'error'>('connecting');
  const [statusDetail, setStatusDetail] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [stopping, setStopping] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [connectAttempt, setConnectAttempt] = useState(0);
  const [activeModel, setActiveModel] = useState<string | null>(null);
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [attachSelection, setAttachSelection] = useState<boolean>(true);
  const [themePref, setThemePref] = useState<ThemePreference>(() => getStoredThemePreference());
  const connection = useRef<SidecarConnection | null>(null);
  const chatRef = useRef<HTMLElement | null>(null);

  // Sync theme attribute whenever themePref changes or system preference updates
  useEffect(() => {
    applyTheme(resolveIsDark(themePref));
    if (themePref !== 'auto') return;

    if (typeof window !== 'undefined' && window.matchMedia) {
      const media = window.matchMedia('(prefers-color-scheme: dark)');
      const listener = (): void => {
        applyTheme(resolveIsDark('auto'));
      };
      media.addEventListener('change', listener);
      return () => media.removeEventListener('change', listener);
    }
  }, [themePref]);

  const cycleTheme = useCallback((): void => {
    setThemePref((prev) => {
      const next: ThemePreference = prev === 'auto' ? 'dark' : prev === 'dark' ? 'light' : 'auto';
      setStoredThemePreference(next);
      applyTheme(resolveIsDark(next));
      return next;
    });
  }, []);

  // Monitor Excel active selection changes
  useEffect(() => {
    void getSelectedRangeInfo().then((info) => {
      if (info) setSelection(info);
    });
    const unregister = registerSelectionChangeListener((info) => {
      setSelection(info);
    });
    return () => unregister();
  }, []);

  // Keep the latest message in view.
  useEffect(() => {
    chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight });
  }, [items, busy]);

  const unpair = useCallback((): void => {
    clearToken();
    connection.current?.close();
    setPaired(false);
  }, []);

  // Stale token (sidecar data reset) -> back to pairing automatically.
  useEffect(() => {
    if (!paired) return;
    void checkAuth().then((ok) => {
      if (!ok) unpair();
    });
  }, [paired, unpair]);

  const refreshActiveModel = useCallback((): void => {
    void api<ProviderList>('/api/settings/providers').then(
      (list) => {
        const active = list.providers.find((p) => p.id === list.activeProviderId);
        setActiveModel(active ? active.model : null);
        if (!active) setShowSettings(true); // first run: configure a model
      },
      () => {},
    );
  }, []);

  useEffect(() => {
    if (paired) refreshActiveModel();
  }, [paired, refreshActiveModel]);

  // Auto-reconnect with a fixed backoff while the sidecar is unreachable.
  useEffect(() => {
    if (!paired || status === 'connected' || status === 'connecting') return;
    const timer = setTimeout(() => {
      setStatus('connecting');
      setConnectAttempt((n) => n + 1);
    }, 8000);
    return () => clearTimeout(timer);
  }, [paired, status]);

  const handleEvent = useCallback((ev: AgentEventPayload) => {
    switch (ev.kind) {
      case 'text':
        setItems((prev) => [...prev, { kind: 'assistant', text: ev.text }]);
        break;
      case 'tool_use':
        setItems((prev) => [...prev, { kind: 'tool', text: `${ev.tool} ${ev.summary}`.trim() }]);
        break;
      case 'changeset_staged':
        setItems((prev) => [
          ...prev,
          { kind: 'changeset', id: ev.changeSetId, preview: ev.preview, status: 'staged' },
        ]);
        break;
      case 'changeset_applied':
        setItems((prev) =>
          prev.map((it) =>
            it.kind === 'changeset' && it.id === ev.changeSetId ? { ...it, status: 'applied' } : it,
          ),
        );
        break;
      case 'changeset_undone':
        setItems((prev) =>
          prev.map((it) =>
            it.kind === 'changeset' && it.id === ev.changeSetId ? { ...it, status: 'undone' } : it,
          ),
        );
        break;
      case 'error':
        setItems((prev) => [...prev, { kind: 'error', text: ev.message }]);
        break;
      case 'stopped':
        setItems((prev) => [
          ...prev,
          { kind: 'stopped', text: ev.reason ?? 'Generation stopped by user.' },
        ]);
        break;
      case 'done':
        setBusy(false);
        setStopping(false);
        break;
    }
  }, []);

  useEffect(() => {
    if (!paired) return;
    let conn: SidecarConnection | null = null;
    void (async () => {
      try {
        const { id, name } = await getWorkbookIdentity();
        conn = connect({
          workbookId: id,
          workbookName: name,
          onEvent: handleEvent,
          onStatus: (s, detail) => {
            setStatus(s === 'connected' ? 'connected' : s);
            setStatusDetail(detail ?? '');
          },
        });
        connection.current = conn;
      } catch (e) {
        setStatus('error');
        setStatusDetail(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => conn?.close();
  }, [paired, connectAttempt, handleEvent]);

  const stopGeneration = useCallback((): void => {
    setStopping(true);
    connection.current?.abort('user_requested');
  }, []);

  const sendMessage = (): void => {
    const text = input.trim();
    if (!text || busy || status !== 'connected') return;

    const hasAttachment = attachSelection && selection;
    const displayText = hasAttachment
      ? `[📎 Selection: ${selection.address}]\n${text}`
      : text;
    const payloadText = hasAttachment
      ? `[Context: Active Selection is ${selection.address} (${selection.cellCount} cells)]\n${text}`
      : text;

    setItems((prev) => [...prev, { kind: 'user', text: displayText }]);
    setInput('');
    setBusy(true);
    setStopping(false);
    connection.current?.sendChat(payloadText);
  };

  if (!paired) {
    return <PairingScreen onPaired={() => setPaired(true)} />;
  }

  return (
    <div className="app">
      <header>
        <span className={`dot ${status}`} title={statusDetail} />
        <strong>Excel AI</strong>
        {activeModel && !showSettings && (
          <button className="chip" title="Change model" onClick={() => setShowSettings(true)}>
            {activeModel}
          </button>
        )}
        {status !== 'connected' && (
          <button
            className="ghost"
            onClick={() => {
              setStatus('connecting');
              setConnectAttempt((n) => n + 1);
            }}
          >
            Reconnect
          </button>
        )}
        <button
          type="button"
          className="ghost theme-toggle-btn"
          title={`Theme: ${themePref} (click to cycle Auto / Dark / Light)`}
          onClick={cycleTheme}
          aria-label={`Current theme ${themePref}. Click to cycle.`}
        >
          {themePref === 'auto' ? '🌓' : themePref === 'dark' ? '🌙' : '☀️'}
        </button>
        <button
          className="ghost"
          onClick={() => {
            if (showSettings) refreshActiveModel();
            setShowSettings((s) => !s);
          }}
        >
          {showSettings ? 'Chat' : 'Settings'}
        </button>
      </header>
      {status !== 'connected' && (
        <p className="hint banner">
          {status === 'connecting'
            ? 'Connecting to the sidecar…'
            : `Not connected${statusDetail ? ` — ${statusDetail}` : ''}. Is the sidecar running? (pnpm dev:sidecar)`}{' '}
          <button className="link" onClick={unpair}>
            Re-pair
          </button>
        </p>
      )}

      {showSettings ? (
        <ModelsScreen onChanged={refreshActiveModel} />
      ) : (
        <>
          <main className="chat" ref={chatRef}>
            {items.length === 0 && (
              <div className="empty">
                <div className="empty-title">Ask anything about this workbook</div>
                <div className="empty-sub">Edits are proposed first — nothing changes until you approve.</div>
              </div>
            )}
            {items.map((item, i) => (
              <ChatBubble
                key={i}
                item={item}
                onApprove={(id) => connection.current?.approve(id)}
                onReject={(id) => {
                  connection.current?.reject(id);
                  setItems((prev) =>
                    prev.map((it) =>
                      it.kind === 'changeset' && it.id === id ? { ...it, status: 'rejected' } : it,
                    ),
                  );
                }}
                onUndo={(id) => connection.current?.undo(id)}
              />
            ))}
            {busy && (
              <div className="typing" aria-label="Working">
                <span /><span /><span />
              </div>
            )}
          </main>
          <Composer
            input={input}
            setInput={setInput}
            busy={busy}
            stopping={stopping}
            status={status}
            selection={selection}
            attachSelection={attachSelection}
            setAttachSelection={setAttachSelection}
            onSendMessage={sendMessage}
            onStopGeneration={stopGeneration}
          />
        </>
      )}
    </div>
  );
}
