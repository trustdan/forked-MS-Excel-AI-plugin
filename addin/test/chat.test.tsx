import { describe, expect, it, vi } from 'vitest';
import { render, screen, fireEvent } from '@testing-library/react';
import { ChatBubble } from '../src/taskpane/components/chat/ChatBubble.js';
import { Composer } from '../src/taskpane/components/chat/Composer.js';

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
});
