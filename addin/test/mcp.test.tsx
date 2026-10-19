import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as connection from '../src/connection.js';
import { McpScreen } from '../src/taskpane/components/settings/McpScreen.js';

describe('McpScreen Component', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('renders approved paths and MCP server list, and allows adding a new path', async () => {
    let paths = ['C:\\approved\\data'];
    const servers = [
      {
        id: 'srv-1',
        name: 'Sales SQLite DB',
        enabled: true,
        transport: 'builtin',
        type: 'sqlite',
        status: 'connected',
        toolCount: 4,
      },
    ];

    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/mcp/paths') {
        if (!init || init.method === undefined || init.method === 'GET') {
          return { paths };
        }
        if (init.method === 'POST') {
          const body = JSON.parse(init.body as string);
          paths = [...paths, body.path];
          return { ok: true, paths };
        }
        if (init.method === 'DELETE') {
          const body = JSON.parse(init.body as string);
          paths = paths.filter((p) => p !== body.path);
          return { ok: true, paths };
        }
      }
      if (path === '/api/mcp/servers') {
        return { servers };
      }
      return {};
    });

    render(<McpScreen />);

    // Wait for approved paths and server to load
    await waitFor(() => {
      expect(screen.getByText('C:\\approved\\data')).toBeDefined();
      expect(screen.getByText('Sales SQLite DB')).toBeDefined();
      expect(screen.getByText('● 4 tools')).toBeDefined();
    });

    // Add new approved path
    const input = screen.getByPlaceholderText(/e.g. C:\\Users\\User\\Documents\\Data/i);
    fireEvent.change(input, { target: { value: 'C:\\approved\\csvs' } });

    const addBtn = screen.getByRole('button', { name: /\+ Add Path/i });
    fireEvent.click(addBtn);

    await waitFor(() => {
      expect(screen.getByText('C:\\approved\\csvs')).toBeDefined();
      expect(screen.getByText(/Path approved: C:\\approved\\csvs/i)).toBeDefined();
    });
  });

  it('allows opening connector form and testing connection', async () => {
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/mcp/paths') {
        return { paths: ['C:\\data'] };
      }
      if (path === '/api/mcp/servers' && (!init || init.method === undefined)) {
        return { servers: [] };
      }
      if (path === '/api/mcp/servers/test' && init?.method === 'POST') {
        return { ok: true, tools: [{ name: 'csv_list_files' }, { name: 'csv_query' }] };
      }
      return {};
    });

    render(<McpScreen />);

    // Click "+ Add Connector"
    const addConnectorBtn = screen.getByRole('button', { name: /\+ Add Connector/i });
    fireEvent.click(addConnectorBtn);

    // Form should appear
    expect(screen.getByText('Add Local Connector')).toBeDefined();

    // Fill form name
    const nameInput = screen.getByPlaceholderText(/e.g. Orders Database/i);
    fireEvent.change(nameInput, { target: { value: 'Orders CSVs' } });

    // Test connection
    const testBtn = screen.getByRole('button', { name: /Test Connection/i });
    fireEvent.click(testBtn);

    await waitFor(() => {
      expect(screen.getByText(/Connected successfully! Exposes 2 tool\(s\)/i)).toBeDefined();
    });
  });
});
