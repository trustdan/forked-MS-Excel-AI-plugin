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
import type { ActivityStep, ChatItem, ProviderList } from './types.js';
import { PairingScreen } from './components/pairing/PairingScreen.js';
import { ModelsScreen } from './components/settings/ModelsScreen.js';
import { McpScreen } from './components/settings/McpScreen.js';
import { ChatBubble } from './components/chat/ChatBubble.js';
import { Composer, type AttachedFile } from './components/chat/Composer.js';

function closeActiveActivities(items: ChatItem[]): ChatItem[] {
  let changed = false;
  const next = items.map((it) => {
    if (it.kind === 'activity' && it.active) {
      changed = true;
      return { ...it, active: false };
    }
    if (it.kind === 'assistant' && it.streaming) {
      changed = true;
      return { kind: 'assistant' as const, text: it.text };
    }
    return it;
  });
  return changed ? next : items;
}

function appendThinkingDelta(items: ChatItem[], delta: string): ChatItem[] {
  if (!delta) return items;
  const rawLast = items[items.length - 1];
  // If a prior step streamed an unfinalized assistant draft that was rejected/retried
  // before a finalized `text` event arrived, discard the draft bubble and resume activity.
  const baseItems =
    rawLast && rawLast.kind === 'assistant' && rawLast.streaming ? items.slice(0, -1) : items;
  const last = baseItems[baseItems.length - 1];
  if (last && last.kind === 'activity') {
    const lastStep = last.steps[last.steps.length - 1];
    if (lastStep && lastStep.kind === 'thinking' && lastStep.durationMs === undefined) {
      return [
        ...baseItems.slice(0, -1),
        {
          ...last,
          active: true,
          steps: [
            ...last.steps.slice(0, -1),
            { ...lastStep, text: lastStep.text + delta },
          ],
        },
      ];
    }
    return [
      ...baseItems.slice(0, -1),
      {
        ...last,
        active: true,
        steps: [...last.steps, { kind: 'thinking', text: delta }],
      },
    ];
  }
  return [
    ...baseItems,
    {
      kind: 'activity',
      steps: [{ kind: 'thinking', text: delta }],
      totalDurationMs: 0,
      active: true,
    },
  ];
}

function appendTextDelta(items: ChatItem[], delta: string): ChatItem[] {
  if (!delta) return items;
  const last = items[items.length - 1];
  if (last && last.kind === 'assistant' && last.streaming) {
    return [
      ...items.slice(0, -1),
      { kind: 'assistant', text: last.text + delta, streaming: true },
    ];
  }
  return [...closeActiveActivities(items), { kind: 'assistant', text: delta, streaming: true }];
}

function appendActivityStep(items: ChatItem[], step: ActivityStep): ChatItem[] {
  const addDuration = step.kind === 'thinking' ? (step.durationMs ?? 0) : 0;
  const last = items[items.length - 1];

  // When text_delta has already started a streaming assistant bubble at the end of items,
  // attach finalized thinking/tool events to the activity block preceding the streaming bubble
  // so the streaming bubble remains last and is replaced (not duplicated) by the final text event.
  if (last && last.kind === 'assistant' && last.streaming) {
    const beforeStreaming = items.slice(0, -1);
    const prevItem = beforeStreaming[beforeStreaming.length - 1];
    if (prevItem && prevItem.kind === 'activity') {
      const prevStep = prevItem.steps[prevItem.steps.length - 1];
      if (
        step.kind === 'thinking' &&
        prevStep &&
        prevStep.kind === 'thinking' &&
        prevStep.durationMs === undefined
      ) {
        return [
          ...beforeStreaming.slice(0, -1),
          {
            ...prevItem,
            steps: [...prevItem.steps.slice(0, -1), step],
            totalDurationMs: prevItem.totalDurationMs + addDuration,
          },
          last,
        ];
      }
      return [
        ...beforeStreaming.slice(0, -1),
        {
          ...prevItem,
          steps: [...prevItem.steps, step],
          totalDurationMs: prevItem.totalDurationMs + addDuration,
        },
        last,
      ];
    }
    return [
      ...beforeStreaming,
      {
        kind: 'activity',
        steps: [step],
        totalDurationMs: addDuration,
        active: false,
      },
      last,
    ];
  }

  if (last && last.kind === 'activity' && last.active) {
    const lastStep = last.steps[last.steps.length - 1];
    // Finalize an in-flight streamed thinking step when the completed thinking event arrives
    if (step.kind === 'thinking' && lastStep && lastStep.kind === 'thinking' && lastStep.durationMs === undefined) {
      return [
        ...items.slice(0, -1),
        {
          ...last,
          steps: [...last.steps.slice(0, -1), step],
          totalDurationMs: last.totalDurationMs + addDuration,
        },
      ];
    }
    return [
      ...items.slice(0, -1),
      {
        ...last,
        steps: [...last.steps, step],
        totalDurationMs: last.totalDurationMs + addDuration,
      },
    ];
  }
  return [
    ...items,
    {
      kind: 'activity',
      steps: [step],
      totalDurationMs: addDuration,
      active: true,
    },
  ];
}

export function applyAgentEventToItems(prev: ChatItem[], ev: AgentEventPayload): ChatItem[] {
  switch (ev.kind) {
    case 'thinking_delta':
      return appendThinkingDelta(prev, ev.text);
    case 'text_delta':
      return appendTextDelta(prev, ev.text);
    case 'thinking':
      return appendActivityStep(prev, {
        kind: 'thinking',
        text: ev.text,
        durationMs: ev.durationMs,
      });
    case 'tool_use':
      return appendActivityStep(prev, {
        kind: 'tool',
        text: `${ev.tool} ${ev.summary}`.trim(),
      });
    case 'text': {
      const last = prev[prev.length - 1];
      if (last && last.kind === 'assistant' && last.streaming) {
        return [...closeActiveActivities(prev.slice(0, -1)), { kind: 'assistant', text: ev.text }];
      }
      return [...closeActiveActivities(prev), { kind: 'assistant', text: ev.text }];
    }
    case 'changeset_staged':
      return [
        ...closeActiveActivities(prev),
        {
          kind: 'changeset',
          id: ev.changeSetId,
          preview: ev.preview,
          undoFidelity: ev.undoFidelity,
          status: 'staged',
        },
      ];
    case 'changeset_status':
      return prev.map((it) =>
        it.kind === 'changeset' && it.id === ev.changeSetId ? { ...it, status: ev.status } : it,
      );
    case 'error':
      return [...closeActiveActivities(prev), { kind: 'error', text: ev.message }];
    case 'stopped':
      return [
        ...closeActiveActivities(prev),
        { kind: 'stopped', text: ev.reason ?? 'Generation stopped by user.' },
      ];
    case 'done':
      return closeActiveActivities(prev);
  }
}

export function App(): JSX.Element {
  const [paired, setPaired] = useState<boolean>(() => getToken() !== null);
  const [status, setStatus] = useState<'connecting' | 'connected' | 'disconnected' | 'error'>('connecting');
  const [statusDetail, setStatusDetail] = useState('');
  const [items, setItems] = useState<ChatItem[]>([]);
  const [input, setInput] = useState('');
  const [busy, setBusy] = useState(false);
  const [stopping, setStopping] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [settingsTab, setSettingsTab] = useState<'models' | 'mcp'>('models');
  const [connectAttempt, setConnectAttempt] = useState(0);
  const [activeModel, setActiveModel] = useState<string | null>(null);
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [attachSelection, setAttachSelection] = useState<boolean>(true);
  const [attachedFile, setAttachedFile] = useState<AttachedFile | null>(null);
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

  // Monitor Excel active selection changes and refresh on window focus
  useEffect(() => {
    const refreshSelection = (): void => {
      void getSelectedRangeInfo().then((info) => {
        if (info) setSelection(info);
      });
    };
    refreshSelection();
    const unregister = registerSelectionChangeListener((info) => {
      setSelection(info);
    });
    window.addEventListener('focus', refreshSelection);
    return () => {
      unregister();
      window.removeEventListener('focus', refreshSelection);
    };
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
    setItems((prev) => applyAgentEventToItems(prev, ev));
    if (ev.kind === 'done') {
      setBusy(false);
      setStopping(false);
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
            if (s === 'disconnected' || s === 'error') {
              setBusy(false);
              setStopping(false);
              setItems((prev) => closeActiveActivities(prev));
            }
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
    const rawText = input.trim();
    if ((!rawText && !attachedFile) || busy || status !== 'connected') return;

    setInput('');
    const currentFile = attachedFile;
    setAttachedFile(null);
    setBusy(true);
    setStopping(false);

    const text =
      rawText ||
      (currentFile && /\.(csv|tsv|json)$/i.test(currentFile.name)
        ? `Import ${currentFile.name} starting at A1`
        : currentFile
          ? `Help me create the spreadsheet described in ${currentFile.name}`
          : '');

    void (async () => {
      let currentSelection = selection;
      if (attachSelection && !currentSelection) {
        currentSelection = await getSelectedRangeInfo();
        if (currentSelection) setSelection(currentSelection);
      }

      const hasSelection = attachSelection && currentSelection;
      let displayText = text;
      let payloadText = text;

      if (currentFile) {
        displayText = `[📄 Attached File: ${currentFile.name}]\n${displayText}`;
        payloadText = `[Attached Document: ${currentFile.name} (full contents provided inline below — read directly from this message; do NOT search the filesystem for this file)]\n\`\`\`\n${currentFile.content}\n\`\`\`\n\n${payloadText}`;
      }

      if (hasSelection && currentSelection) {
        displayText = `[📎 Selection: ${currentSelection.address}]\n${displayText}`;
        payloadText = `[Context: Active Selection is ${currentSelection.address} (${currentSelection.cellCount} cells)]\n${payloadText}`;
      }

      setItems((prev) => [...prev, { kind: 'user', text: displayText }]);
      connection.current?.sendChat(payloadText);
    })();
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
        <div className="settings-container">
          <div
            className="settings-tabs"
            style={{
              display: 'flex',
              gap: '4px',
              padding: '8px 16px',
              borderBottom: '1px solid var(--border)',
              background: 'var(--bg-card)',
            }}
          >
            <button
              type="button"
              className={settingsTab === 'models' ? 'primary' : 'ghost'}
              onClick={() => setSettingsTab('models')}
              style={{ fontSize: '0.85em', padding: '4px 10px' }}
            >
              Models
            </button>
            <button
              type="button"
              className={settingsTab === 'mcp' ? 'primary' : 'ghost'}
              onClick={() => setSettingsTab('mcp')}
              style={{ fontSize: '0.85em', padding: '4px 10px' }}
            >
              MCP Connectors
            </button>
          </div>
          {settingsTab === 'models' ? (
            <ModelsScreen onChanged={refreshActiveModel} />
          ) : (
            <McpScreen />
          )}
        </div>
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
            attachedFile={attachedFile}
            setAttachedFile={setAttachedFile}
            onSendMessage={sendMessage}
            onStopGeneration={stopGeneration}
          />
        </>
      )}
    </div>
  );
}
