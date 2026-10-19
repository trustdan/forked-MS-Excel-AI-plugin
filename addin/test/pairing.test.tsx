import { describe, expect, it, vi, beforeEach, afterEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { PairingScreen } from '../src/taskpane/components/pairing/PairingScreen.js';

describe('PairingScreen', () => {
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('renders pairing instructions and input', () => {
    (globalThis.fetch as any).mockResolvedValueOnce({ ok: true });
    render(<PairingScreen onPaired={() => {}} />);

    expect(screen.getByText('Pair with the companion app')).toBeDefined();
    expect(screen.getByPlaceholderText('ABC123')).toBeDefined();
    expect(screen.getByRole('button', { name: 'Pair' })).toBeDefined();
  });

  it('keeps Pair button disabled until a 6-character code is entered', () => {
    (globalThis.fetch as any).mockResolvedValueOnce({ ok: true });
    render(<PairingScreen onPaired={() => {}} />);

    const button = screen.getByRole('button', { name: 'Pair' }) as HTMLButtonElement;
    const input = screen.getByPlaceholderText('ABC123') as HTMLInputElement;

    expect(button.disabled).toBe(true);

    fireEvent.change(input, { target: { value: 'abc' } });
    expect(input.value).toBe('ABC');
    expect(button.disabled).toBe(true);

    fireEvent.change(input, { target: { value: 'abc123' } });
    expect(input.value).toBe('ABC123');
    expect(button.disabled).toBe(false);
  });

  it('calls onPaired when pairing succeeds', async () => {
    (globalThis.fetch as any).mockImplementation(async (url: string) => {
      if (url.includes('/api/health')) {
        return { ok: true };
      }
      if (url.includes('/api/pair')) {
        return {
          ok: true,
          json: async () => ({ token: 'mock-valid-token' }),
        };
      }
      return { ok: false };
    });

    const onPaired = vi.fn();
    render(<PairingScreen onPaired={onPaired} />);

    const input = screen.getByPlaceholderText('ABC123');
    const button = screen.getByRole('button', { name: 'Pair' });

    fireEvent.change(input, { target: { value: 'XYZ789' } });
    fireEvent.click(button);

    await waitFor(() => {
      expect(onPaired).toHaveBeenCalled();
    });
  });
});
