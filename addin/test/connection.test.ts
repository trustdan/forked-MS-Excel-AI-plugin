import { afterEach, describe, expect, it, vi } from 'vitest';
import { connect, type AgentEventPayload } from '../src/connection';

class FakeWebSocket {
  static OPEN = 1;
  static last: FakeWebSocket | null = null;
  readyState = FakeWebSocket.OPEN;
  sent: string[] = [];
  onopen: (() => void) | null = null;
  onclose: (() => void) | null = null;
  onerror: (() => void) | null = null;
  onmessage: ((event: { data: string }) => void) | null = null;
  constructor(readonly url: string) {
    FakeWebSocket.last = this;
  }
  send(data: string): void {
    this.sent.push(data);
  }
  close(): void {}
}

afterEach(() => {
  vi.unstubAllGlobals();
  localStorage.clear();
});

describe('sidecar connection', () => {
  it('surfaces a protocol_error from the sidecar as a chat error', () => {
    vi.stubGlobal('WebSocket', FakeWebSocket);
    localStorage.setItem('excelai-token', 'tok');
    const events: AgentEventPayload[] = [];
    connect({ workbookId: 'wb', workbookName: 'Wb.xlsx', onEvent: (ev) => events.push(ev), onStatus: () => {} });

    FakeWebSocket.last!.onmessage!({
      data: JSON.stringify({
        v: 1,
        id: 'e',
        type: 'protocol_error',
        payload: { message: 'Malformed frame', issues: ['payload.message: Expected string'], refId: 'c1' },
      }),
    });

    expect(events).toEqual([{ kind: 'error', message: 'Sidecar rejected a message: Malformed frame' }]);
  });
});
