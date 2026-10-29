import { describe, expect, it, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import { ChatBubble } from '../src/taskpane/components/chat/ChatBubble.js';
import { Composer } from '../src/taskpane/components/chat/Composer.js';
import { ThinkingBlock } from '../src/taskpane/components/chat/ThinkingBlock.js';
import { applyAgentEventToItems } from '../src/taskpane/App.js';
import type { ChatItem } from '../src/taskpane/types.js';

describe('Chat Components', () => {
  describe('ChatBubble', () => {
    it('renders user message', () => {
      render(
        <ChatBubble
          item={{ kind: 'user', text: 'Hello AI' }}
          onApprove={() => {}}
          onReject={() => {}}
          onUndo={() => {}}
        />,
      );
      expect(screen.getByText('Hello AI')).toBeDefined();
    });

    it('renders assistant message with markdown', () => {
      render(
        <ChatBubble
          item={{ kind: 'assistant', text: '**Bold text**' }}
          onApprove={() => {}}
          onReject={() => {}}
          onUndo={() => {}}
        />,
      );
      const strong = screen.getByText('Bold text');
      expect(strong.tagName.toLowerCase()).toBe('strong');
    });

    it('renders tool note', () => {
      render(
        <ChatBubble
          item={{ kind: 'tool', text: 'read_range Sheet1!A1:B2' }}
          onApprove={() => {}}
          onReject={() => {}}
          onUndo={() => {}}
        />,
      );
      expect(screen.getByText(/read_range Sheet1!A1:B2/)).toBeDefined();
    });

    it('renders error bubble', () => {
      render(
        <ChatBubble
          item={{ kind: 'error', text: 'Something failed' }}
          onApprove={() => {}}
          onReject={() => {}}
          onUndo={() => {}}
        />,
      );
      expect(screen.getByText('Something failed')).toBeDefined();
    });

    it('renders stopped bubble', () => {
      render(
        <ChatBubble
          item={{ kind: 'stopped', text: 'Generation stopped by user.' }}
          onApprove={() => {}}
          onReject={() => {}}
          onUndo={() => {}}
        />,
      );
      expect(screen.getByText('Generation stopped by user.')).toBeDefined();
    });
  });

  describe('Composer', () => {
    it('renders textarea and send button when idle', () => {
      const onSendMessage = vi.fn();
      render(
        <Composer
          input="Test prompt"
          setInput={() => {}}
          busy={false}
          stopping={false}
          status="connected"
          selection={null}
          attachSelection={false}
          setAttachSelection={() => {}}
          onSendMessage={onSendMessage}
          onStopGeneration={() => {}}
        />,
      );

      const textarea = screen.getByPlaceholderText(/Ask about this sheet/);
      expect(textarea).toBeDefined();

      const sendBtn = screen.getByRole('button', { name: 'Send' });
      expect(sendBtn).toBeDefined();

      fireEvent.click(sendBtn);
      expect(onSendMessage).toHaveBeenCalled();
    });

    it('renders stop button when busy and triggers onStopGeneration on click', () => {
      const onStopGeneration = vi.fn();
      render(
        <Composer
          input="Test prompt"
          setInput={() => {}}
          busy={true}
          stopping={false}
          status="connected"
          selection={null}
          attachSelection={false}
          setAttachSelection={() => {}}
          onSendMessage={() => {}}
          onStopGeneration={onStopGeneration}
        />,
      );

      const stopBtn = screen.getByRole('button', { name: 'Stop generation' });
      expect(stopBtn).toBeDefined();

      fireEvent.click(stopBtn);
      expect(onStopGeneration).toHaveBeenCalled();
    });

    it('triggers onStopGeneration when Escape key is pressed while busy', () => {
      const onStopGeneration = vi.fn();
      render(
        <Composer
          input="Test prompt"
          setInput={() => {}}
          busy={true}
          stopping={false}
          status="connected"
          selection={null}
          attachSelection={false}
          setAttachSelection={() => {}}
          onSendMessage={() => {}}
          onStopGeneration={onStopGeneration}
        />,
      );

      const textarea = screen.getByPlaceholderText(/Ask about this sheet/);
      fireEvent.keyDown(textarea, { key: 'Escape' });
      expect(onStopGeneration).toHaveBeenCalled();
    });

    it('renders selection chip and allows toggling attachment', () => {
      const setAttachSelection = vi.fn();
      render(
        <Composer
          input=""
          setInput={() => {}}
          busy={false}
          stopping={false}
          status="connected"
          selection={{ address: 'Sheet1!A1:C10', cellCount: 30 }}
          attachSelection={true}
          setAttachSelection={setAttachSelection}
          onSendMessage={() => {}}
          onStopGeneration={() => {}}
        />,
      );

      expect(screen.getByText(/Sheet1!A1:C10/)).toBeDefined();
      expect(screen.getByText('Attached')).toBeDefined();

      const chip = screen.getByRole('button', { name: /Sheet1!A1:C10/ });
      fireEvent.click(chip);
      expect(setAttachSelection).toHaveBeenCalled();
    });

    it('renders attached file chip and allows detaching file', () => {
      const setAttachedFile = vi.fn();
      render(
        <Composer
          input=""
          setInput={() => {}}
          busy={false}
          stopping={false}
          status="connected"
          selection={null}
          attachSelection={false}
          setAttachSelection={() => {}}
          attachedFile={{ name: 'revenue.csv', size: 1024, content: 'a,b\n1,2' }}
          setAttachedFile={setAttachedFile}
          onSendMessage={() => {}}
          onStopGeneration={() => {}}
        />,
      );

      expect(screen.getByText('revenue.csv')).toBeDefined();
      const detachBtn = screen.getByRole('button', { name: /Detach file/i });
      fireEvent.click(detachBtn);
      expect(setAttachedFile).toHaveBeenCalledWith(null);
    });
  });

  describe('ThinkingBlock & Activity Drawer', () => {
    it('starts expanded while active, auto-collapses on completion, and toggles on click', () => {
      const steps = [
        { kind: 'thinking' as const, text: 'Inspecting revenue columns A1:D10', durationMs: 1400 },
        { kind: 'tool' as const, text: 'read_range Sheet1!A1:D10' },
      ];

      const { rerender } = render(
        <ThinkingBlock steps={steps} totalDurationMs={1400} active={true} />,
      );

      // While active, it is expanded and shows Thinking… · 1 tool call
      const toggleBtn = screen.getByRole('button', { name: /Thinking… · 1 tool call/i });
      expect(toggleBtn.getAttribute('aria-expanded')).toBe('true');
      expect(screen.getByText('Inspecting revenue columns A1:D10')).toBeDefined();
      expect(screen.getByText('read_range Sheet1!A1:D10')).toBeDefined();

      // When active becomes false, it auto-collapses and updates header summary
      rerender(<ThinkingBlock steps={steps} totalDurationMs={1400} active={false} />);
      const completedBtn = screen.getByRole('button', {
        name: /Thought for 1\.4s · 1 tool call/i,
      });
      expect(completedBtn.getAttribute('aria-expanded')).toBe('false');
      expect(screen.queryByRole('region')).toBeNull();

      // Clicking the header expands the drawer to inspect thoughts & tools
      fireEvent.click(completedBtn);
      expect(completedBtn.getAttribute('aria-expanded')).toBe('true');
      expect(screen.getByText('Inspecting revenue columns A1:D10')).toBeDefined();
      expect(screen.getByText('read_range Sheet1!A1:D10')).toBeDefined();
    });

    it('groups contiguous thinking and tool_use events into a single collapsible activity item and closes it on final answer', () => {
      let items: ChatItem[] = [{ kind: 'user', text: 'Calculate totals' }];

      items = applyAgentEventToItems(items, {
        kind: 'thinking',
        text: 'Let me check the sheet layout.',
        durationMs: 1200,
      });
      items = applyAgentEventToItems(items, {
        kind: 'tool_use',
        tool: 'get_workbook_map',
        summary: '',
      });
      items = applyAgentEventToItems(items, {
        kind: 'thinking',
        text: 'Now I will read Sheet1!A1:B5.',
        durationMs: 800,
      });
      items = applyAgentEventToItems(items, {
        kind: 'tool_use',
        tool: 'read_range',
        summary: 'Sheet1!A1:B5',
      });

      expect(items).toHaveLength(2);
      expect(items[1]).toEqual({
        kind: 'activity',
        active: true,
        totalDurationMs: 2000,
        steps: [
          { kind: 'thinking', text: 'Let me check the sheet layout.', durationMs: 1200 },
          { kind: 'tool', text: 'get_workbook_map' },
          { kind: 'thinking', text: 'Now I will read Sheet1!A1:B5.', durationMs: 800 },
          { kind: 'tool', text: 'read_range Sheet1!A1:B5' },
        ],
      });

      // When assistant text arrives, the activity drawer is marked active: false
      items = applyAgentEventToItems(items, {
        kind: 'text',
        text: 'The sum is 250.',
      });
      expect(items).toHaveLength(3);
      expect(items[1]).toMatchObject({ kind: 'activity', active: false, totalDurationMs: 2000 });
      expect(items[2]).toEqual({ kind: 'assistant', text: 'The sum is 250.' });
    });

    it('coalesces live thinking_delta and text_delta chunks and finalizes cleanly when thinking/text events arrive', () => {
      let items: ChatItem[] = [{ kind: 'user', text: 'Analyze Q1' }];

      // Stream reasoning tokens
      items = applyAgentEventToItems(items, { kind: 'thinking_delta', text: 'Reading ' });
      items = applyAgentEventToItems(items, { kind: 'thinking_delta', text: 'columns A:C…' });

      expect(items).toHaveLength(2);
      expect(items[1]).toEqual({
        kind: 'activity',
        active: true,
        totalDurationMs: 0,
        steps: [{ kind: 'thinking', text: 'Reading columns A:C…' }],
      });

      // Finalized thinking event replaces in-progress thinking step with durationMs
      items = applyAgentEventToItems(items, {
        kind: 'thinking',
        text: 'Reading columns A:C…',
        durationMs: 950,
      });
      expect(items[1]).toEqual({
        kind: 'activity',
        active: true,
        totalDurationMs: 950,
        steps: [{ kind: 'thinking', text: 'Reading columns A:C…', durationMs: 950 }],
      });

      // Stream assistant text tokens — closes active activity drawer and appends to streaming bubble
      items = applyAgentEventToItems(items, { kind: 'text_delta', text: 'Q1 total ' });
      items = applyAgentEventToItems(items, { kind: 'text_delta', text: 'is $1,200.' });

      expect(items).toHaveLength(3);
      expect(items[1]).toMatchObject({ kind: 'activity', active: false, totalDurationMs: 950 });
      expect(items[2]).toEqual({
        kind: 'assistant',
        text: 'Q1 total is $1,200.',
        streaming: true,
      });

      // Finalized text event replaces streaming bubble cleanly
      items = applyAgentEventToItems(items, { kind: 'text', text: 'Q1 total is $1,200.' });
      expect(items).toHaveLength(3);
      expect(items[2]).toEqual({
        kind: 'assistant',
        text: 'Q1 total is $1,200.',
      });
    });

    it('does not duplicate the assistant bubble when finalized thinking and text events arrive after both thinking_delta and text_delta', () => {
      let items: ChatItem[] = [{ kind: 'user', text: 'Draft balance sheet' }];

      items = applyAgentEventToItems(items, { kind: 'thinking_delta', text: 'Mapping rows…' });
      items = applyAgentEventToItems(items, { kind: 'text_delta', text: 'Done — staged!' });
      // provider.chat() finishes and loop.ts emits finalized thinking followed by finalized text
      items = applyAgentEventToItems(items, {
        kind: 'thinking',
        text: 'Mapping rows…',
        durationMs: 1500,
      });
      items = applyAgentEventToItems(items, {
        kind: 'text',
        text: 'Done — staged!',
      });

      expect(items).toHaveLength(3);
      expect(items[1]).toEqual({
        kind: 'activity',
        active: false,
        totalDurationMs: 1500,
        steps: [{ kind: 'thinking', text: 'Mapping rows…', durationMs: 1500 }],
      });
      expect(items[2]).toEqual({
        kind: 'assistant',
        text: 'Done — staged!',
      });
    });
  });
});
