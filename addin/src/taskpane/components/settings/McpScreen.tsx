import { useCallback, useEffect, useState } from 'react';
import { api } from '../../../connection.js';
import type { McpServerConfig, McpServerStatus, McpServerType, McpTransport } from '@excelai/shared';

interface AddServerForm {
  name: string;
  transport: McpTransport;
  type: McpServerType;
  path: string;
  command: string;
  args: string;
}

const EMPTY_FORM: AddServerForm = {
  name: '',
  transport: 'builtin',
  type: 'sqlite',
  path: '',
  command: '',
  args: '',
};

export function McpScreen(): JSX.Element {
  const [approvedPaths, setApprovedPaths] = useState<string[]>([]);
  const [newPath, setNewPath] = useState('');
  const [servers, setServers] = useState<McpServerStatus[]>([]);
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState<AddServerForm>(EMPTY_FORM);
  const [message, setMessage] = useState('');
  const [busyAction, setBusyAction] = useState('');

  const loadPaths = useCallback((): void => {
    void api<{ paths: string[] }>('/api/mcp/paths')
      .then((r) => setApprovedPaths(r.paths ?? []))
      .catch((e: unknown) => setMessage(String(e instanceof Error ? e.message : e)));
  }, []);

  const loadServers = useCallback((): void => {
    void api<{ servers: McpServerStatus[] }>('/api/mcp/servers')
      .then((r) => setServers(r.servers ?? []))
      .catch((e: unknown) => setMessage(String(e instanceof Error ? e.message : e)));
  }, []);

  useEffect(() => {
    loadPaths();
    loadServers();
  }, [loadPaths, loadServers]);

  const addPath = (): void => {
    const trimmed = newPath.trim();
    if (!trimmed) return;
    setBusyAction('add-path');
    void api<{ ok: boolean; paths: string[] }>('/api/mcp/paths', {
      method: 'POST',
      body: JSON.stringify({ path: trimmed }),
    })
      .then((r) => {
        setApprovedPaths(r.paths ?? []);
        setNewPath('');
        setMessage(`Path approved: ${trimmed}`);
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const removePath = (target: string): void => {
    setBusyAction(`rm-path-${target}`);
    void api<{ ok: boolean; paths: string[] }>('/api/mcp/paths', {
      method: 'DELETE',
      body: JSON.stringify({ path: target }),
    })
      .then((r) => {
        setApprovedPaths(r.paths ?? []);
        setMessage(`Path removed: ${target}`);
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const toggleServer = (id: string): void => {
    setBusyAction(`toggle-${id}`);
    void api<{ ok: boolean; enabled: boolean }>(`/api/mcp/servers/${id}/toggle`, {
      method: 'POST',
    })
      .then(() => {
        loadServers();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const deleteServer = (id: string): void => {
    setBusyAction(`del-${id}`);
    void api<{ ok: boolean }>(`/api/mcp/servers/${id}`, {
      method: 'DELETE',
    })
      .then(() => {
        setMessage('Connector removed.');
        loadServers();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const testServerConfig = (cfg: Partial<McpServerConfig>): void => {
    setBusyAction('test-server');
    setMessage('Testing connector…');
    void api<{ ok: boolean; tools: any[]; error?: string }>('/api/mcp/servers/test', {
      method: 'POST',
      body: JSON.stringify(cfg),
    })
      .then((r) => {
        if (r.ok) {
          setMessage(`✓ Connected successfully! Exposes ${r.tools.length} tool(s).`);
        } else {
          setMessage(`✗ ${r.error ?? 'Connection failed'}`);
        }
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  const buildPayloadFromForm = (): McpServerConfig => {
    const isBuiltin = form.transport === 'builtin';
    const cfg: McpServerConfig = {
      id: '',
      name: form.name.trim() || `${form.type} connector`,
      enabled: true,
      transport: form.transport,
      allowedPaths: approvedPaths,
    };

    if (isBuiltin) {
      cfg.type = form.type;
      if (form.type === 'sqlite') {
        cfg.options = { dbPath: form.path.trim() };
      } else if (form.type === 'csv-directory') {
        cfg.options = { dirPath: form.path.trim() };
      } else if (form.type === 'filesystem') {
        cfg.options = { rootPath: form.path.trim() };
      }
    } else {
      cfg.command = form.command.trim();
      cfg.args = form.args.trim() ? form.args.split(/\s+/) : [];
    }
    return cfg;
  };

  const saveConnector = (): void => {
    const payload = buildPayloadFromForm();
    setBusyAction('save-connector');
    void api<{ ok: boolean }>('/api/mcp/servers', {
      method: 'POST',
      body: JSON.stringify(payload),
    })
      .then(() => {
        setMessage('Connector saved and initialized.');
        setShowAdd(false);
        setForm(EMPTY_FORM);
        loadServers();
      })
      .catch((e: unknown) => setMessage(`✗ ${e instanceof Error ? e.message : String(e)}`))
      .finally(() => setBusyAction(''));
  };

  return (
    <div className="models-screen mcp-screen">
      <div className="screen-header">
        <h3>Model Context Protocol (MCP) Connectors</h3>
        <p className="screen-desc">
          Connect local SQLite databases, CSV directories, and filesystems to the AI assistant.
        </p>
      </div>

      {message && <div className="banner info-banner">{message}</div>}

      {/* Approved Paths Section */}
      <section className="settings-section" style={{ marginBottom: '20px' }}>
        <h4>🛡️ Explicit User-Approved Paths</h4>
        <p className="field-hint" style={{ fontSize: '0.85em', color: 'var(--text-muted)' }}>
          To guarantee zero data exfiltration, local connectors are strictly sandboxed to these paths.
        </p>
        <div className="path-list" style={{ marginTop: '8px', marginBottom: '8px' }}>
          {approvedPaths.length === 0 ? (
            <div className="empty-sub" style={{ fontStyle: 'italic', fontSize: '0.9em' }}>
              No approved paths yet. Add a directory or file path below to authorize access.
            </div>
          ) : (
            approvedPaths.map((p) => (
              <div
                key={p}
                className="path-item"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  padding: '4px 8px',
                  background: 'var(--bg-secondary)',
                  borderRadius: '4px',
                  marginBottom: '4px',
                  fontSize: '0.85em',
                  fontFamily: 'monospace',
                }}
              >
                <span>{p}</span>
                <button
                  type="button"
                  className="ghost"
                  style={{ padding: '2px 6px', color: 'var(--danger)' }}
                  onClick={() => removePath(p)}
                  disabled={busyAction.startsWith('rm-path')}
                  title="Remove path"
                >
                  ×
                </button>
              </div>
            ))
          )}
        </div>
        <div style={{ display: 'flex', gap: '8px' }}>
          <input
            type="text"
            placeholder="e.g. C:\Users\User\Documents\Data or /home/user/data"
            value={newPath}
            onChange={(e) => setNewPath(e.target.value)}
            style={{ flex: 1 }}
          />
          <button
            type="button"
            className="primary"
            onClick={addPath}
            disabled={!newPath.trim() || busyAction === 'add-path'}
          >
            + Add Path
          </button>
        </div>
      </section>

      {/* Configured Connectors Section */}
      <section className="settings-section">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h4>Configured MCP Servers</h4>
          {!showAdd && (
            <button type="button" className="ghost" onClick={() => setShowAdd(true)}>
              + Add Connector
            </button>
          )}
        </div>

        {servers.length === 0 && !showAdd && (
          <div className="empty" style={{ padding: '16px 0' }}>
            <div className="empty-sub">No MCP connectors configured. Click "+ Add Connector" to attach local data.</div>
          </div>
        )}

        <div className="server-list" style={{ marginTop: '10px' }}>
          {servers.map((s) => (
            <div
              key={s.id}
              className="server-card"
              style={{
                border: '1px solid var(--border)',
                borderRadius: '6px',
                padding: '10px',
                marginBottom: '8px',
                background: 'var(--bg-card)',
              }}
            >
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <strong>{s.name}</strong>{' '}
                  <span
                    style={{
                      fontSize: '0.75em',
                      padding: '2px 6px',
                      borderRadius: '4px',
                      background: 'var(--bg-secondary)',
                      marginLeft: '6px',
                    }}
                  >
                    {s.type ?? s.transport}
                  </span>
                </div>
                <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                  <span
                    style={{
                      fontSize: '0.8em',
                      color: s.status === 'connected' ? 'var(--success, #22c55e)' : 'var(--text-muted)',
                    }}
                  >
                    {s.enabled ? (s.status === 'connected' ? `● ${s.toolCount} tools` : '○ disconnected') : 'Disabled'}
                  </span>
                  <button
                    type="button"
                    className="ghost"
                    style={{ padding: '2px 8px', fontSize: '0.8em' }}
                    onClick={() => toggleServer(s.id)}
                  >
                    {s.enabled ? 'Disable' : 'Enable'}
                  </button>
                  <button
                    type="button"
                    className="ghost"
                    style={{ padding: '2px 8px', fontSize: '0.8em', color: 'var(--danger)' }}
                    onClick={() => deleteServer(s.id)}
                  >
                    Delete
                  </button>
                </div>
              </div>
              {s.error && (
                <div style={{ color: 'var(--danger)', fontSize: '0.8em', marginTop: '4px' }}>
                  Error: {s.error}
                </div>
              )}
            </div>
          ))}
        </div>

        {/* Add Connector Form */}
        {showAdd && (
          <div
            className="add-server-form"
            style={{
              border: '1px solid var(--accent)',
              borderRadius: '6px',
              padding: '12px',
              marginTop: '12px',
              background: 'var(--bg-card)',
            }}
          >
            <h5>Add Local Connector</h5>

            <label style={{ display: 'block', marginTop: '8px', fontSize: '0.85em' }}>
              Connector Name:
              <input
                type="text"
                placeholder="e.g. Orders Database"
                value={form.name}
                onChange={(e) => setForm({ ...form, name: e.target.value })}
                style={{ width: '100%', marginTop: '4px' }}
              />
            </label>

            <label style={{ display: 'block', marginTop: '8px', fontSize: '0.85em' }}>
              Connector Type:
              <select
                value={form.transport === 'builtin' ? form.type : 'stdio'}
                onChange={(e) => {
                  const val = e.target.value;
                  if (val === 'stdio') {
                    setForm({ ...form, transport: 'stdio' });
                  } else {
                    setForm({ ...form, transport: 'builtin', type: val as McpServerType });
                  }
                }}
                style={{ width: '100%', marginTop: '4px' }}
              >
                <option value="sqlite">Local SQLite Database (.db)</option>
                <option value="csv-directory">Local CSV / TSV Directory</option>
                <option value="filesystem">Local Filesystem Folder</option>
                <option value="stdio">Custom stdio MCP Server (command/args)</option>
              </select>
            </label>

            {form.transport === 'builtin' ? (
              <label style={{ display: 'block', marginTop: '8px', fontSize: '0.85em' }}>
                {form.type === 'sqlite'
                  ? 'SQLite Database File Path:'
                  : form.type === 'csv-directory'
                    ? 'CSV Directory Path:'
                    : 'Root Directory Path:'}
                <input
                  type="text"
                  placeholder={
                    form.type === 'sqlite'
                      ? 'C:\\data\\sales.db'
                      : 'C:\\data\\csv_reports'
                  }
                  value={form.path}
                  onChange={(e) => setForm({ ...form, path: e.target.value })}
                  style={{ width: '100%', marginTop: '4px' }}
                />
              </label>
            ) : (
              <>
                <label style={{ display: 'block', marginTop: '8px', fontSize: '0.85em' }}>
                  Command:
                  <input
                    type="text"
                    placeholder="e.g. npx or python"
                    value={form.command}
                    onChange={(e) => setForm({ ...form, command: e.target.value })}
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </label>
                <label style={{ display: 'block', marginTop: '8px', fontSize: '0.85em' }}>
                  Arguments (space separated):
                  <input
                    type="text"
                    placeholder="e.g. -y @modelcontextprotocol/server-sqlite"
                    value={form.args}
                    onChange={(e) => setForm({ ...form, args: e.target.value })}
                    style={{ width: '100%', marginTop: '4px' }}
                  />
                </label>
              </>
            )}

            <div style={{ display: 'flex', gap: '8px', marginTop: '12px', justifyContent: 'flex-end' }}>
              <button
                type="button"
                className="ghost"
                onClick={() => {
                  setShowAdd(false);
                  setForm(EMPTY_FORM);
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="ghost"
                onClick={() => testServerConfig(buildPayloadFromForm())}
                disabled={busyAction === 'test-server'}
              >
                Test Connection
              </button>
              <button
                type="button"
                className="primary"
                onClick={saveConnector}
                disabled={busyAction === 'save-connector'}
              >
                Save Connector
              </button>
            </div>
          </div>
        )}
      </section>
    </div>
  );
}
