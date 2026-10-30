import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as connection from '../src/connection.js';
import { ModelsScreen } from '../src/taskpane/components/settings/ModelsScreen.js';

describe('ModelsScreen Component', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it('renders provider list and allows adding an embedded GGUF model', async () => {
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/settings/providers' && (!init || init.method === undefined)) {
        return {
          providers: [
            {
              id: 'p1',
              kind: 'openai-compatible',
              model: 'local-model',
              baseUrl: 'http://127.0.0.1:1234/v1',
            },
          ],
          activeProviderId: 'p1',
        };
      }
      if (path === '/api/providers/models' && init?.method === 'POST') {
        return { models: ['qwen2.5-7b-instruct-q4_k_m.gguf'] };
      }
      if (path === '/api/settings/providers' && init?.method === 'POST') {
        return { id: 'p2', activeProviderId: 'p1' };
      }
      return {};
    });

    const onChanged = vi.fn();
    render(<ModelsScreen onChanged={onChanged} />);

    // Wait for initial list to load
    await waitFor(() => {
      expect(screen.getByText('local-model')).toBeDefined();
      expect(screen.getByText('active')).toBeDefined();
    });

    // Click "+ Add a model"
    const addBtn = screen.getByRole('button', { name: /\+ Add a model/i });
    fireEvent.click(addBtn);

    // Select "Embedded GGUF"
    const select = screen.getByLabelText(/Provider/i) as HTMLSelectElement;
    fireEvent.change(select, { target: { value: 'embedded' } });

    // Verify hint is displayed
    expect(screen.getByText(/Runs completely in-process/i)).toBeDefined();

    // Verify "Scan local models" button appears and Base URL input is hidden
    const scanBtn = screen.getByRole('button', { name: /Scan local models/i });
    expect(scanBtn).toBeDefined();
    expect(screen.queryByLabelText(/Base URL/i)).toBeNull();

    // Click "Scan local models"
    fireEvent.click(scanBtn);

    await waitFor(() => {
      expect(mockApi).toHaveBeenCalledWith(
        '/api/providers/models',
        expect.objectContaining({ method: 'POST' }),
      );
      expect(screen.getByText(/1 models available/i)).toBeDefined();
    });
  });

  it('shows a settings-file problem and dismisses it', async () => {
    let problem: object | undefined = {
      kind: 'unreadable',
      detail: 'not valid JSON',
      preservedAs: 'C:/Users/me/.excelai/settings.json.corrupt-2026-10-30T12-00-00-000Z',
      at: '2026-10-30T12:00:00.000Z',
    };
    const mockApi = vi.spyOn(connection, 'api').mockImplementation(async (path, init) => {
      if (path === '/api/settings/providers' && !init?.method) {
        return { providers: [], activeProviderId: null, ...(problem ? { settingsProblem: problem } : {}) };
      }
      if (path === '/api/settings/problem' && init?.method === 'DELETE') {
        problem = undefined;
        return { ok: true };
      }
      return {};
    });

    render(<ModelsScreen onChanged={vi.fn()} />);
    const alert = await screen.findByRole('alert');
    expect(alert.textContent).toContain('unreadable');
    expect(alert.textContent).toContain('settings.json.corrupt-2026-10-30T12-00-00-000Z');

    fireEvent.click(screen.getByRole('button', { name: /Dismiss/i }));
    await waitFor(() => expect(screen.queryByRole('alert')).toBeNull());
    expect(mockApi).toHaveBeenCalledWith('/api/settings/problem', { method: 'DELETE' });
  });
});
