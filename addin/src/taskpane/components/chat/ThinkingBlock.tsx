import { useEffect, useRef, useState } from 'react';
import type { ActivityStep } from '../../types.js';

export interface ThinkingBlockProps {
  steps: ActivityStep[];
  totalDurationMs?: number;
  active?: boolean;
}

export function formatThinkingDuration(ms: number): string {
  if (!Number.isFinite(ms) || ms <= 0) return '';
  if (ms < 100) return '<0.1s';
  if (ms < 10_000) {
    const secs = (ms / 1000).toFixed(1).replace(/\.0$/, '');
    return `${secs}s`;
  }
  return `${Math.round(ms / 1000)}s`;
}

function getLatestPreviewLine(step: ActivityStep | undefined): string {
  if (!step) return '';
  if (step.kind === 'tool') {
    return `🔧 ${step.text}`;
  }
  const lines = step.text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  return lines[lines.length - 1] ?? '';
}

export function ThinkingBlock({
  steps,
  totalDurationMs = 0,
  active = false,
}: ThinkingBlockProps): JSX.Element {
  const [expanded, setExpanded] = useState<boolean>(() => Boolean(active));
  const userToggled = useRef(false);
  const prevActive = useRef<boolean>(Boolean(active));
  const bodyRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (prevActive.current && !active && !userToggled.current) {
      setExpanded(false);
    } else if (!prevActive.current && active && !userToggled.current) {
      setExpanded(true);
    }
    prevActive.current = Boolean(active);
  }, [active]);

  const lastStepText = steps[steps.length - 1]?.text ?? '';
  useEffect(() => {
    if (expanded && active && bodyRef.current) {
      bodyRef.current.scrollTo?.({ top: bodyRef.current.scrollHeight });
    }
  }, [expanded, active, steps.length, lastStepText]);

  if (steps.length === 0) return <></>;

  const thoughtCount = steps.filter((s) => s.kind === 'thinking').length;
  const toolCount = steps.filter((s) => s.kind === 'tool').length;
  const durationLabel = totalDurationMs > 0 ? formatThinkingDuration(totalDurationMs) : '';

  let summaryText: string;
  if (active) {
    if (thoughtCount > 0 && toolCount > 0) {
      summaryText = `Thinking… · ${toolCount} ${toolCount === 1 ? 'tool call' : 'tool calls'}`;
    } else if (thoughtCount > 0) {
      summaryText = 'Thinking…';
    } else {
      summaryText = `Working… · ${toolCount} ${toolCount === 1 ? 'tool call' : 'tool calls'}`;
    }
  } else if (thoughtCount > 0) {
    const base = durationLabel ? `Thought for ${durationLabel}` : 'Thought';
    const suffix =
      toolCount > 0
        ? ` · ${toolCount} ${toolCount === 1 ? 'tool call' : 'tool calls'}`
        : thoughtCount > 1
          ? ` · ${thoughtCount} steps`
          : '';
    summaryText = `${base}${suffix}`;
  } else {
    summaryText = `Used ${toolCount} ${toolCount === 1 ? 'tool' : 'tools'}`;
  }

  const latestPreview = getLatestPreviewLine(steps[steps.length - 1]);

  const handleToggle = (): void => {
    userToggled.current = true;
    setExpanded((prev) => !prev);
  };

  return (
    <div className={`thinking-block${active ? ' active' : ''}${expanded ? ' expanded' : ''}`}>
      <button
        type="button"
        className="thinking-header"
        aria-expanded={expanded}
        onClick={handleToggle}
      >
        <span className="thinking-chevron" aria-hidden="true">
          {expanded ? '▾' : '▸'}
        </span>
        <span className="thinking-icon" aria-hidden="true">
          {active ? '✦' : '💭'}
        </span>
        <span className="thinking-summary">{summaryText}</span>
        {!expanded && latestPreview && (
          <span className="thinking-preview" title={latestPreview}>
            {latestPreview}
          </span>
        )}
      </button>
      {expanded && (
        <div
          ref={bodyRef}
          className="thinking-body"
          role="region"
          aria-label="Agent thinking and tool progress"
        >
          {steps.map((step, idx) =>
            step.kind === 'thinking' ? (
              <div key={idx} className="thinking-step">
                <div className="thinking-step-text">{step.text}</div>
                {step.durationMs !== undefined && step.durationMs > 0 && (
                  <span className="thinking-step-duration">
                    {formatThinkingDuration(step.durationMs)}
                  </span>
                )}
              </div>
            ) : (
              <div key={idx} className="thinking-tool-step">
                <span className="thinking-tool-icon" aria-hidden="true">
                  🔧
                </span>
                <code className="thinking-tool-text">{step.text}</code>
              </div>
            ),
          )}
        </div>
      )}
    </div>
  );
}
