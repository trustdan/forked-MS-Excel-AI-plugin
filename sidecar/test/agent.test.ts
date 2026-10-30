import { describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn, type AgentEvent } from '../src/agent/loop.js';

async function runScript(
  wb: WorkbookEmulator,
  provider: ScriptedProvider,
): Promise<{ events: AgentEvent[]; changeSets: ChangeSetManager; stagedId?: string }> {
  const changeSets = new ChangeSetManager(wb, 'wb-test');
  const events: AgentEvent[] = [];
  const result = await runAgentTurn({
    provider,
    executor: wb,
    changeSets,
    workbookName: 'Test.xlsx',
    history: [],
    userMessage: 'do the thing',
    onEvent: (ev) => events.push(ev),
  });
  const out: { events: AgentEvent[]; changeSets: ChangeSetManager; stagedId?: string } = {
    events,
    changeSets,
  };
  if (result.stagedChangeSetId !== undefined) out.stagedId = result.stagedChangeSetId;
  return out;
}

describe('agent loop (full integration against the emulator)', () => {
  it('reads the workbook and stages a write as a change-set', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 5);
    wb.setCell('Sheet1!A2', 7);

    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'get_workbook_map', args: {} }] },
      { toolCalls: [{ name: 'read_range', args: { range: 'Sheet1!A1:A2' } }] },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A3',
              cells: [[{ formula: '=SUM(A1:A2)' }]],
              reason: 'total of A1:A2',
            },
          },
        ],
      },
      { text: 'I staged a SUM formula in A3.' },
    ]);

    const { events, changeSets, stagedId } = await runScript(wb, provider);

    expect(stagedId).toBeDefined();
    expect(wb.getValue('Sheet1!A3')).toBeNull(); // nothing applied yet

    const staged = events.find((e) => e.kind === 'changeset_staged');
    expect(staged).toBeDefined();

    const result = await changeSets.apply(stagedId!);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!A3')).toBe(12); // formula actually recalculates

    // The model saw real read results.
    const readResult = provider.requests.at(2)?.messages.findLast((m) => m.role === 'tool');
    expect(readResult && 'content' in readResult ? readResult.content : '').toContain('5');
  });

  it('feeds schema violations back for repair instead of executing them', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      // missing required "reason", wrong cells shape
      { toolCalls: [{ name: 'write_range', args: { range: 'Sheet1!A1', cells: 'oops' } }] },
      // repaired call (plain scalars are values; objects are formulas)
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!A1', cells: [[1]], reason: 'fixed' },
          },
        ],
      },
      { text: 'done' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);

    // The invalid call produced an error tool-result that the model saw next turn.
    const repairInput = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(repairInput && 'content' in repairInput ? repairInput.content : '').toContain(
      'Invalid arguments',
    );

    // Only the valid op was staged.
    expect(stagedId).toBeDefined();
    expect(changeSets.get(stagedId!)?.ops).toHaveLength(1);
  });

  it('rejects unknown tools and malformed JSON gracefully', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'delete_everything', args: {} }] },
      { toolCalls: [{ name: 'read_range', args: '{not json' }] },
      { text: 'giving up politely' },
    ]);

    const { stagedId, events } = await runScript(wb, provider);
    expect(stagedId).toBeUndefined();
    expect(events.some((e) => e.kind === 'changeset_staged')).toBe(false);

    const unknownToolReply = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(unknownToolReply && 'content' in unknownToolReply ? unknownToolReply.content : '').toContain(
      'Unknown tool',
    );
  });

  it('stages formatting alongside values and applies both', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!B1', cells: [[1234.5]], reason: 'value' },
          },
          {
            name: 'format_range',
            args: {
              range: 'Sheet1!B1',
              format: { bold: true, numberFormat: '#,##0.00', fillColor: '#FFF2CC' },
              reason: 'highlight the total',
            },
          },
        ],
      },
      { text: 'Wrote and formatted B1.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    expect(stagedId).toBeDefined();
    const cs = changeSets.get(stagedId!)!;
    expect(cs.ops).toHaveLength(2);
    expect(cs.undoFidelity).toBe('partial'); // formatting is not snapshot-restorable yet

    const result = await changeSets.apply(stagedId!);
    expect(result.ok).toBe(true);
    expect(wb.getValue('Sheet1!B1')).toBe(1234.5);
    expect(wb.getFormat('Sheet1!B1')).toMatchObject({
      bold: true,
      numberFormat: '#,##0.00',
      fillColor: '#FFF2CC',
    });
  });

  it('rejects malformed format payloads', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'format_range',
            args: { range: 'Sheet1!B1', format: {}, reason: 'empty format' },
          },
        ],
      },
      { text: 'ok' },
    ]);
    const { stagedId } = await runScript(wb, provider);
    expect(stagedId).toBeUndefined(); // empty format object fails validation
  });

  it('supports emulated tool calling for models without native tools', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'hello');

    class NoToolsProvider extends ScriptedProvider {
      override readonly supportsTools = false;
    }
    const provider = new NoToolsProvider([
      { text: '{"tool": "read_range", "args": {"range": "Sheet1!A1"}}' },
      { text: 'The cell says hello.' },
    ]);

    const { events } = await runScript(wb, provider);
    expect(events.some((e) => e.kind === 'tool_use' && e.tool === 'read_range')).toBe(true);
    expect(events.some((e) => e.kind === 'text' && e.text.includes('hello'))).toBe(true);
  });

  it('stops at the iteration bound even if the model loops forever', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const looping = Array.from({ length: 50 }, () => ({
      toolCalls: [{ name: 'get_workbook_map', args: {} }],
    }));
    const provider = new ScriptedProvider(looping);

    const changeSets = new ChangeSetManager(wb, 'wb-test');
    await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Test.xlsx',
      history: [],
      userMessage: 'loop',
      onEvent: () => {},
      maxIterations: 5,
    });
    expect(provider.requests.length).toBe(5);
  });

  it('aborts cleanly when abortSignal is triggered during provider chat', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const controller = new AbortController();
    const events: AgentEvent[] = [];

    const provider: ScriptedProvider = new (class extends ScriptedProvider {
      constructor() {
        super([]);
      }
      override async chat(): Promise<never> {
        controller.abort();
        const err = new Error('The operation was aborted');
        err.name = 'AbortError';
        throw err;
      }
    })();

    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Test.xlsx',
      history: [],
      userMessage: 'stop me',
      abortSignal: controller.signal,
      onEvent: (ev) => events.push(ev),
    });

    expect(events.some((e) => e.kind === 'stopped')).toBe(true);
    expect(result.stagedChangeSetId).toBeUndefined();
    expect(result.history).toEqual([]);
  });

  it('aborts cleanly and skips staging when abortSignal triggers during tool calls', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const controller = new AbortController();
    const events: AgentEvent[] = [];

    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!A1', cells: [[42]], reason: 'should not stage' },
          },
        ],
      },
    ]);

    // Abort right before or after the first tool execution
    const origGetCell = wb.readRange.bind(wb);
    wb.readRange = async (range: string) => {
      controller.abort();
      return origGetCell(range);
    };

    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Test.xlsx',
      history: [],
      userMessage: 'write and cancel',
      abortSignal: controller.signal,
      onEvent: (ev) => {
        events.push(ev);
        if (ev.kind === 'tool_use') {
          controller.abort();
        }
      },
    });

    expect(events.some((e) => e.kind === 'stopped')).toBe(true);
    expect(result.stagedChangeSetId).toBeUndefined();
  });

  it('validates write_range cell dimensions against target range immediately and deduplicates same-range writes', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A2:E2',
              cells: [[null, '2026-10-01', 'Office Supplies', 'Supplies', 'Expense', 120.5]],
              reason: '6 cols in 5-col range',
            },
          },
        ],
      },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A2:E2',
              cells: [['2026-10-01', 'Office Supplies', 'Supplies', 'Expense', 120.5]],
              reason: 'initial valid write',
            },
          },
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A2:E2',
              cells: [['2026-10-01', 'Office Supplies', 'Supplies', 'Expense', 150]],
              reason: 'updated valid write on same range',
            },
          },
        ],
      },
      { text: 'Staged clean row.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    const repairMsg = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(repairMsg && 'content' in repairMsg ? repairMsg.content : '').toContain(
      'Row 1 has 6 column(s), expected 5',
    );

    expect(stagedId).toBeDefined();
    const cs = changeSets.get(stagedId!)!;
    expect(cs.ops).toHaveLength(1);
    expect(cs.ops[0]?.kind === 'write_range' ? cs.ops[0].cells[0]?.[4] : undefined).toEqual({
      value: 150,
    });
  });

  it('recovers with a nudge, preserves partial reasoning draft, and expands maxTokens when a reasoning model exhausts max_tokens', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        text: '',
        stopReason: 'max_tokens',
        reasoning: 'Drafting 2024 balance sheet: Cash 401900, Raw Materials 419500',
      },
      {
        text: '',
        stopReason: 'max_tokens',
        reasoning: 'Row 1: Pax Vans, Row 2: Cash 401900, Row 3: Raw Materials 419500',
      },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: { range: 'Sheet1!A1:B1', cells: [['Net Total', 100]], reason: 'summary' },
          },
        ],
      },
      { text: 'Completed summary.' },
    ]);

    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];
    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Test.xlsx',
      history: [],
      userMessage: 'do the thing',
      onEvent: (ev) => events.push(ev),
    });

    expect(result.stagedChangeSetId).toBeDefined();
    expect(events.some((e) => e.kind === 'text' && e.text.includes('Completed summary.'))).toBe(true);
    expect(events.some((e) => e.kind === 'error')).toBe(false);

    // Normal step uses 8192 + reasoningEffort: 'low'; recovery step expands to 16384 and includes reasoning tail
    expect(provider.requests.at(0)?.maxTokens).toBe(8192);
    expect(provider.requests.at(0)?.reasoningEffort).toBe('low');
    expect(provider.requests.at(1)?.maxTokens).toBe(16384);
    const firstRecoveryMsg = provider.requests.at(1)?.messages.at(-1);
    expect(firstRecoveryMsg?.role === 'user' ? firstRecoveryMsg.content : '').toContain('Cash 401900');

    // Recovery nudge messages must not remain in persisted conversation history
    expect(
      result.history.some((m) => m.role === 'user' && m.content.includes('[System note:')),
    ).toBe(false);
  });

  it('normalizes workbook filename used as sheet name and breaks duplicate tool-call loops by disabling tools', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 'Header');
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];

    const provider = new ScriptedProvider([
      { toolCalls: [{ name: 'read_range', args: { range: 'Book 2026-10-28.xlsx!A1:B5' } }] },
      { toolCalls: [{ name: 'read_range', args: { range: 'Book 2026-10-28.xlsx!A1:B5' } }] },
      { text: 'I read Sheet1!A1:B5. Which cells would you like me to update?' },
    ]);

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Book 2026-10-28.xlsx',
      history: [],
      userMessage: 'help with this sheet',
      onEvent: (ev) => events.push(ev),
    });

    // First call resolved Book 2026-10-28.xlsx!A1:B5 -> Sheet1!A1:B5 and read 'Header'
    const firstToolMsg = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(firstToolMsg && 'content' in firstToolMsg ? firstToolMsg.content : '').toContain('Header');

    // Second identical call triggered duplicate guard and disabled tools on step 3
    const secondToolMsg = provider.requests.at(2)?.messages.findLast((m) => m.role === 'tool');
    expect(secondToolMsg && 'content' in secondToolMsg ? secondToolMsg.content : '').toContain(
      'duplicateCall',
    );
    expect(provider.requests.at(2)?.tools).toBeUndefined();

    // Duplicate assistant/tool pair was stripped from persisted history
    expect(result.history.filter((m) => m.role === 'tool')).toHaveLength(1);
  });

  it('rescues raw XML <tool_call> tags emitted in content (e.g. GLM/Flint <tool_call>get_active_sheet</arg_value></tool_call>) and aliases get_active_sheet to get_workbook_map', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];

    const provider = new ScriptedProvider([
      {
        text: '<tool_call>get_active_sheet</arg_value></tool_call>',
        toolCalls: [],
        stopReason: 'end',
      },
      {
        text: '<tool_call>write_range<arg_key>range</arg_key><arg_value>Sheet1!A1:B1</arg_value><arg_key>cells</arg_key><arg_value>[["Van Model", "Capacity"]]</arg_value><arg_key>reason</arg_key><arg_value>Add Pax Vans header</arg_value></tool_call>',
        toolCalls: [],
        stopReason: 'end',
      },
      { text: 'Created the Pax Vans table header in Sheet1!A1:B1.' },
    ]);

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Book 2026-10-28.xlsx',
      history: [],
      userMessage: 'Build the Pax Vans table',
      onEvent: (ev) => events.push(ev),
    });

    // Raw <tool_call> tags must never leak into user-visible text events
    expect(events.some((e) => e.kind === 'text' && e.text.includes('<tool_call>'))).toBe(false);
    // get_active_sheet was aliased and executed as get_workbook_map
    expect(events.some((e) => e.kind === 'tool_use' && e.tool === 'get_workbook_map')).toBe(true);
    // XML write_range was rescued and staged
    expect(result.stagedChangeSetId).toBeDefined();
    await changeSets.apply(result.stagedChangeSetId!);
    expect(wb.getValue('Sheet1!A1')).toBe('Van Model');
    expect(wb.getValue('Sheet1!B1')).toBe('Capacity');
  });

  it('auto-expands write_range dimensions from a start cell or approximate row count and pads shorter header rows', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1:B3', // 4 rows passed for a 3-row range, plus 1-col header row
              cells: [
                ['Pax Vans Balance Sheet'],
                ['Cash', 401900],
                ['Raw Materials', 419500],
                ['Total Current Assets', { formula: '=SUM(B2:B3)' }],
              ],
              reason: 'Record 2024 starting balances',
            },
          },
        ],
      },
      { text: 'Staged 2024 starting balances.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    expect(stagedId).toBeDefined();
    const cs = changeSets.get(stagedId!)!;
    expect(cs.ops[0]).toMatchObject({
      kind: 'write_range',
      range: 'Sheet1!A1:B4',
    });

    const applied = await changeSets.apply(stagedId!);
    expect(applied.ok).toBe(true);
    expect(wb.getValue('Sheet1!A1')).toBe('Pax Vans Balance Sheet');
    expect(wb.getValue('Sheet1!B4')).toBe(821400);
  });

  it('emits thinking events with durationMs from reasoning fields and inline <think> tags without leaking <think> into text', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    wb.setCell('Sheet1!A1', 100);

    const provider = new ScriptedProvider([
      {
        reasoning: 'First, let me inspect the workbook structure.',
        toolCalls: [{ name: 'get_workbook_map', args: {} }],
      },
      {
        text: '<think>Now I know Sheet1!A1 has 100. I will report it.</think>Cell A1 contains 100.',
      },
    ]);

    const { events } = await runScript(wb, provider);

    const thinkingEvents = events.filter((e) => e.kind === 'thinking');
    expect(thinkingEvents).toHaveLength(2);
    expect(thinkingEvents[0]).toMatchObject({
      kind: 'thinking',
      text: 'First, let me inspect the workbook structure.',
    });
    if (thinkingEvents[0]?.kind === 'thinking') {
      expect(thinkingEvents[0].durationMs).toBeGreaterThan(0);
    }
    expect(thinkingEvents[1]).toMatchObject({
      kind: 'thinking',
      text: 'Now I know Sheet1!A1 has 100. I will report it.',
    });

    const textEvents = events.filter((e) => e.kind === 'text');
    expect(textEvents).toEqual([{ kind: 'text', text: 'Cell A1 contains 100.' }]);
  });

  it('streams thinking_delta and text_delta events live before finalized thinking and text events', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        reasoningChunks: ['Step 1: ', 'check Sheet1.'],
        reasoning: 'Step 1: check Sheet1.',
        textChunks: ['All ', 'done!'],
        text: 'All done!',
      },
    ]);

    const { events } = await runScript(wb, provider);

    expect(events.filter((e) => e.kind === 'thinking_delta')).toEqual([
      { kind: 'thinking_delta', text: 'Step 1: ' },
      { kind: 'thinking_delta', text: 'check Sheet1.' },
    ]);
    expect(events.filter((e) => e.kind === 'text_delta')).toEqual([
      { kind: 'text_delta', text: 'All ' },
      { kind: 'text_delta', text: 'done!' },
    ]);
    expect(events.some((e) => e.kind === 'thinking' && e.text === 'Step 1: check Sheet1.')).toBe(
      true,
    );
    expect(events.some((e) => e.kind === 'text' && e.text === 'All done!')).toBe(true);
  });

  it('pre-populates the current workbook sheet structure into the system prompt so the model can skip redundant get_workbook_map calls', async () => {
    const wb = new WorkbookEmulator(['Sheet1', 'Summary']);
    wb.setCell('Sheet1!A1', 'Revenue');
    wb.setCell('Sheet1!B2', 500);

    const provider = new ScriptedProvider([{ text: 'Ready.' }]);
    await runScript(wb, provider);

    const sysMsg = provider.requests[0]?.messages.find((m) => m.role === 'system');
    expect(sysMsg?.content).toContain('Current workbook structure (already loaded');
    expect(sysMsg?.content).toContain('Sheet "Sheet1": usedRange Sheet1!A1:B2 (2×2)');
    expect(sysMsg?.content).toContain('Sheet "Summary": empty');
  });

  it('rescues a complete write_range JSON block drafted inside reasoning_content when reasoning hits max_tokens and unwraps 3-D cells nesting', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      {
        text: '',
        stopReason: 'max_tokens',
        reasoning: `Let me write the balance sheet:\n\`\`\`json\n{\n  "range": "Sheet1!A1",\n  "cells": [\n    [["ASSETS:"]],\n    [["Cash", 401900]],\n    [["Raw Materials", 419500]]\n  ],\n  "reason": "Record starting balances"\n}\n\`\`\`\nWait, let me double check...`,
      },
      { text: 'Staged the 2024 balance sheet.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    expect(stagedId).toBeDefined();
    await changeSets.apply(stagedId!);
    expect(wb.getValue('Sheet1!A1')).toBe('ASSETS:');
    expect(wb.getValue('Sheet1!A2')).toBe('Cash');
    expect(wb.getValue('Sheet1!B2')).toBe(401900);
    expect(wb.getValue('Sheet1!A3')).toBe('Raw Materials');
    expect(wb.getValue('Sheet1!B3')).toBe(419500);
  });

  it('auto-balances missing parens, repairs off-by-one SUM/subtotal rows in write_range, replaces same-start-cell drafts, and stops multi-pass rewrite loops', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const provider = new ScriptedProvider([
      // Attempt 1: 31 rows with missing closing paren "=SUM(B7:B10"
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1',
              cells: [
                ['Pax Vans'],
                ['Balance Sheet'],
                ['As of December 31, 2024'],
                [''],
                ['ASSETS:'],
                ['Current Assets:'],
                ['Cash', 401900],
                ['Raw Materials', 419500],
                ['WIP Inventory', 0],
                ['FG Inventory', 0],
                ['Total Current Assets', { formula: '=SUM(B7:B10' }],
              ],
              reason: 'Initial draft',
            },
          },
        ],
      },
      // Attempt 2: 33 rows at Sheet1!A1 with off-by-one liabilities/equity formulas (=SUM(B22:B23), =B24, =SUM(B28:B29), =B25+B30)
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1',
              cells: [
                ['Pax Vans'],
                ['Balance Sheet'],
                ['As of December 31, 2024'],
                [''],
                ['ASSETS:'],
                ['Current Assets:'],
                ['Cash', 401900],
                ['Raw Materials', 419500],
                ['WIP Inventory', 0],
                ['FG Inventory', 0],
                ['Total Current Assets', { formula: '=SUM(B7:B10' }],
                [''],
                ['Non-Current Assets:'],
                ['Property, Plant and Equipment:'],
                ['Rental Vans', 469700],
                ['Manufacturing Equipment, net', 165000],
                ['Property, Plant, and Equip, Net', { formula: '=SUM(B15:B16' }],
                [''],
                ['TOTAL ASSETS', { formula: '=SUM(B11,B17)' }],
                [''],
                ['LIABILITIES:'],
                ['Current Liabilities:'],
                ['Accounts Payable', 220400],
                ['Deferred Revenue (Customer Deposits)', 122500],
                ['Total Current Liabilities', { formula: '=SUM(B22:B23)' }],
                [''],
                ['TOTAL LIABILITIES', { formula: '=B24' }],
                [''],
                ["SHAREHOLDERS' EQUITY:"],
                ['Common Stock', 1250000],
                ['Retained Earnings (Loss)', -136800],
                ["TOTAL SHAREHOLDERS' EQUITY", { formula: '=SUM(B28:B29)' }],
                ["TOTAL LIAB. & SHAREHOLDERS' EQUITY", { formula: '=B25+B30' }],
              ],
              reason: 'Create baseline balance sheet as of 12/31/2024',
            },
          },
        ],
      },
      { text: 'Staged the baseline balance sheet as of December 31, 2024.' },
    ]);

    const { stagedId, changeSets } = await runScript(wb, provider);
    expect(stagedId).toBeDefined();

    // Attempt 2 replaced Attempt 1 (despite 11 rows vs 33 rows) so stagedOpsCount stayed 1
    const cs = changeSets.get(stagedId!)!;
    expect(cs.ops).toHaveLength(1);
    expect(cs.ops[0]?.kind === 'write_range' ? cs.ops[0].range : '').toBe('Sheet1!A1:B33');

    // Step 3 had tools disabled to prevent a 3rd/4th write_range rewrite loop
    expect(provider.requests.at(2)?.tools).toBeUndefined();

    // Apply and verify that all formulas were balanced and off-by-one rows repaired
    const applied = await changeSets.apply(stagedId!);
    expect(applied.ok).toBe(true);
    expect(wb.getValue('Sheet1!B11')).toBe(821400); // =SUM(B7:B10)
    expect(wb.getValue('Sheet1!B17')).toBe(634700); // =SUM(B15:B16)
    expect(wb.getValue('Sheet1!B19')).toBe(1456100); // =SUM(B11,B17)
    expect(wb.getValue('Sheet1!B25')).toBe(342900); // repaired from =SUM(B22:B23) -> =SUM(B23:B24)
    expect(wb.getValue('Sheet1!B27')).toBe(342900); // repaired from =B24 -> =B25
    expect(wb.getValue('Sheet1!B32')).toBe(1113200); // repaired from =SUM(B28:B29) -> =SUM(B30:B31)
    expect(wb.getValue('Sheet1!B33')).toBe(1456100); // repaired from =B25+B30 -> =B27+B32
  });

  it('deduplicates repeated [Attached Document] blocks across multi-turn history so Turn 2+ does not overflow local model context', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const attachedPayload = `[Attached Document: pax-vans.md (full contents provided inline below — read directly from this message; do NOT search the filesystem for this file)]\n\`\`\`\n# Pax Vans Case\nCash: 401,900\nRaw Materials: 419,500\n\`\`\``;

    const priorHistory = [
      {
        role: 'user' as const,
        content: `${attachedPayload}\n\ncan you help me draft starting balances?`,
      },
      {
        role: 'assistant' as const,
        content: 'Got it — before I transform this, I need the 2025 transaction data.',
      },
    ];

    const provider = new ScriptedProvider([{ text: 'Here is the 2025 transaction schedule.' }]);
    const events: AgentEvent[] = [];

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Book.xlsx',
      history: priorHistory,
      userMessage: `${attachedPayload}\n\nHelp me create the spreadsheet described in pax-vans.md`,
      onEvent: (ev) => events.push(ev),
    });

    const sentMessages = provider.requests.at(0)?.messages ?? [];
    const turn1UserMsg = sentMessages.find(
      (m) => m.role === 'user' && m.content.includes('can you help me draft starting balances?'),
    );
    const turn2UserMsg = sentMessages.find(
      (m) =>
        m.role === 'user' &&
        m.content.includes('Help me create the spreadsheet described in pax-vans.md'),
    );

    expect(turn1UserMsg?.content).toContain(
      '[Attached Document: pax-vans.md — full contents included in latest message below]',
    );
    expect(turn1UserMsg?.content).not.toContain('Cash: 401,900');
    expect(turn2UserMsg?.content).toContain('Cash: 401,900');
    expect(result.history).toHaveLength(4);
  });

  it('recovers from left-to-right layout rumination by locking a 5-column side-by-side decision and repairs wrong label-column references in right-hand value columns', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');

    const ruminatingThought = [
      'We have Exhibit 1 balance sheet 12/31/2024.\nAssets:\nCash 401,900\nRaw Materials 419,500\nRental Vans 469,700\nManufacturing Equipment, net 165,000\n\nLiabilities:\nAccounts Payable 220,400\nDeferred Revenue 122,500\nCommon Stock 1,250,000\nRetained Earnings (136,800)',
      'A5: "Cash" B5: 401900\nA6: "Raw Materials" B6: 419500',
      "Maybe they want a horizontal balance sheet with assets in columns B-F and liabilities/equity in columns H-N. This is getting too complicated. Let's step back.",
      "Perhaps they want a single row. Given the ambiguity, maybe I should just create a simple balance sheet. I think I'm overcomplicating.",
    ].join('\n\n');

    const provider = new ScriptedProvider([
      {
        text: '',
        stopReason: 'max_tokens',
        reasoning: ruminatingThought,
      },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1',
              cells: [
                ['Pax Vans - Balance Sheet as of 12/31/2024'],
                [''],
                ['ASSETS', null, null, 'LIABILITIES & EQUITY', null],
                ['Current Assets:', null, null, 'Current Liabilities:', null],
                ['Cash', 401900, null, 'Accounts Payable', 220400],
                ['Raw Materials', 419500, null, 'Deferred Revenue', 122500],
                ['WIP Inventory', 0, null, 'Total Liabilities', { formula: '=SUM(D5:D6)' }],
                ['FG Inventory', 0, null, null, null],
                ['Total Current Assets', { formula: '=SUM(B5:B8)' }, null, "SHAREHOLDERS' EQUITY", null],
                ['', null, null, 'Common Stock', 1250000],
                ['Non-Current Assets:', null, null, 'Retained Earnings', -136800],
                ['Rental Vans', 469700, null, 'Total Equity', { formula: '=SUM(D10:D11)' }],
                ['Manufacturing Equipment, net', 165000, null, null, null],
                ['PPE, Net', { formula: '=SUM(B12:B13)' }, null, null, null],
                ['TOTAL ASSETS', { formula: '=B9+B14' }, null, 'TOTAL LIAB. & EQUITY', { formula: '=D7+D12' }],
              ],
              reason: 'Create 2024 side-by-side balance sheet with Assets on left and Liabilities & Equity on right',
            },
          },
        ],
      },
      { text: 'Created the 2024 opening balance sheet side-by-side.' },
    ]);

    const events: AgentEvent[] = [];
    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Book.xlsx',
      history: [],
      userMessage:
        'can you create an accounting spreadsheet of the 2024 balances that 2025 started with, going from left to right, such that assets is on the left, liabilities and equity is on the right?',
      onEvent: (ev) => events.push(ev),
    });

    // Recovery message locked the 5-column side-by-side layout and kept the initial balances while dropping trailing rumination
    const recoveryUserMsg = provider.requests.at(1)?.messages.at(-1);
    const recoveryText = recoveryUserMsg?.role === 'user' ? recoveryUserMsg.content : '';
    expect(recoveryText).toContain('LAYOUT DECISION LOCKED');
    expect(recoveryText).toContain('Cash 401,900');
    expect(recoveryText).not.toContain("This is getting too complicated. Let's step back.");

    expect(result.stagedChangeSetId).toBeDefined();
    const applied = await changeSets.apply(result.stagedChangeSetId!);
    expect(applied.ok).toBe(true);

    // Left side (Assets in Column B)
    expect(wb.getValue('Sheet1!B9')).toBe(821400); // Total Current Assets
    expect(wb.getValue('Sheet1!B14')).toBe(634700); // PPE, Net
    expect(wb.getValue('Sheet1!B15')).toBe(1456100); // TOTAL ASSETS

    // Right side (Liabilities & Equity in Column E — repaired from D5:D6, D10:D11, and D7+D12)
    expect(wb.getValue('Sheet1!E7')).toBe(342900); // Total Liabilities
    expect(wb.getValue('Sheet1!E12')).toBe(1113200); // Total Equity
    expect(wb.getValue('Sheet1!E15')).toBe(1456100); // TOTAL LIAB. & EQUITY
  });

  it('allows multi-turn layout reorganization after an initial write_range and retries automatically when the model defers overwriting a prior staged version', async () => {
    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');

    // Turn 1: initial vertical balance sheet staged at Sheet1!A1 (>1200 chars of JSON)
    const turn1Provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1',
              cells: [
                ['Pax Vans - Balance Sheet as of 12/31/2024'],
                ['ASSETS'],
                ['Cash', 401900],
                ['Raw Materials', 419500],
                ['WIP Inventory', 0],
                ['FG Inventory', 0],
                ['Rental Vans', 469700],
                ['Manufacturing Equipment, net', 165000],
                ['Total Assets', { formula: '=SUM(B3:B8)' }],
                ['LIABILITIES'],
                ['Accounts Payable', 220400],
                ['Deferred Revenue', 122500],
                ['Total Liabilities', { formula: '=SUM(B11:B12)' }],
                ["SHAREHOLDERS' EQUITY"],
                ['Common Stock', 1250000],
                ['Retained Earnings', -136800],
                ['Total Equity', { formula: '=SUM(B15:B16)' }],
              ],
              reason: 'Initial vertical balance sheet',
            },
          },
        ],
      },
      { text: 'Staged the vertical balance sheet.' },
    ]);

    const turn1Events: AgentEvent[] = [];
    const turn1Res = await runAgentTurn({
      provider: turn1Provider,
      executor: wb,
      changeSets,
      workbookName: 'Book.xlsx',
      history: [],
      userMessage: 'Create the 2024 balance sheet',
      onEvent: (ev) => turn1Events.push(ev),
    });

    expect(turn1Res.stagedChangeSetId).toBeDefined();

    // Turn 2: User asks to reorganize left-to-right with accounts across columns and transactions in rows below.
    // On Step 1 of Turn 2, the model tries to refuse overwriting the prior staged version ("I've kept the staged version as-is...").
    // The loop catches the deferred refusal, suppresses the refusal text, and forces the actual write_range call on Step 2.
    const turn2Provider = new ScriptedProvider([
      {
        text: 'Yes — I can reorganize it into a left-to-right layout. I’ve kept the staged version as-is rather than overwriting it again; I can stage the reorganized version next if you’d like.',
      },
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1',
              cells: [
                ['', 'ASSETS', null, null, null, null, null, 'LIABILITIES', null, "SHAREHOLDERS' EQUITY", null],
                [
                  'Transaction / Description',
                  'Cash',
                  'Raw Materials',
                  'WIP Inventory',
                  'FG Inventory',
                  'Rental Vans',
                  'Manufacturing Equipment, net',
                  'Accounts Payable',
                  'Deferred Revenue',
                  'Common Stock',
                  'Retained Earnings',
                ],
                [
                  '2025 Starting Balance (12/31/2024)',
                  401900,
                  419500,
                  0,
                  0,
                  469700,
                  165000,
                  220400,
                  122500,
                  1250000,
                  -136800,
                ],
              ],
              reason: 'Reorganize accounts horizontally left-to-right with starting balances in Row 3',
            },
          },
        ],
      },
      { text: 'Reorganized the 2025 starting balances horizontally from left to right.' },
    ]);

    const turn2Events: AgentEvent[] = [];
    const turn2Res = await runAgentTurn({
      provider: turn2Provider,
      executor: wb,
      changeSets,
      workbookName: 'Book.xlsx',
      history: turn1Res.history,
      userMessage:
        '[📎 Selection: Sheet1!A1]\ncan you reorganize it left to right rather than up down; it should be sorted by overarching assets on left, liabilities, then equity on the right, such that we can add the accounts / transactions in the rows below the starting balances',
      onEvent: (ev) => turn2Events.push(ev),
    });

    // 1. Prior turn's tool message in providerMessages was updated so it never says "Do NOT call write_range"
    const turn1ToolMsgInTurn2 = turn2Provider.requests
      .at(0)
      ?.messages.find((m) => m.role === 'tool');
    expect(turn1ToolMsgInTurn2?.content).not.toContain('Do NOT call write_range');
    expect(turn1ToolMsgInTurn2?.content).toContain('Completed in a previous turn');

    // 2. The refusal text ("kept the staged version as-is") was NOT emitted as a user-visible text event
    const textEvents = turn2Events.filter((e) => e.kind === 'text');
    expect(textEvents).toEqual([
      {
        kind: 'text',
        text: 'Reorganized the 2025 starting balances horizontally from left to right.',
      },
    ]);

    // 3. The reorganized horizontal table was staged and applies cleanly
    expect(turn2Res.stagedChangeSetId).toBeDefined();
    const applied = await changeSets.apply(turn2Res.stagedChangeSetId!);
    expect(applied.ok).toBe(true);
    expect(wb.getValue('Sheet1!B1')).toBe('ASSETS');
    expect(wb.getValue('Sheet1!H1')).toBe('LIABILITIES');
    expect(wb.getValue('Sheet1!J1')).toBe("SHAREHOLDERS' EQUITY");
    expect(wb.getValue('Sheet1!B2')).toBe('Cash');
    expect(wb.getValue('Sheet1!B3')).toBe(401900);
    expect(wb.getValue('Sheet1!K2')).toBe('Retained Earnings');
    expect(wb.getValue('Sheet1!K3')).toBe(-136800);
  });

  it('resolves default "Sheet1!A1" writes to the actual worksheet tab when Sheet1 does not exist so Apply succeeds', async () => {
    const wb = new WorkbookEmulator(['2025 Balances']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');

    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'write_range',
            args: {
              range: 'Sheet1!A1:C2',
              cells: [
                ['Cash', 'Raw Materials', 'Total Current Assets'],
                [401900, 419500, { formula: '=SUM(A2:B2)' }],
              ],
              reason: 'List 2025 starting balances horizontally',
            },
          },
        ],
      },
      { text: 'Staged starting balances.' },
    ]);

    const events: AgentEvent[] = [];
    const res = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Book.xlsx',
      history: [],
      userMessage: 'Stage starting balances horizontally',
      onEvent: (ev) => events.push(ev),
    });

    expect(res.stagedChangeSetId).toBeDefined();
    const applied = await changeSets.apply(res.stagedChangeSetId!);
    expect(applied.ok).toBe(true);
    expect(wb.getValue("'2025 Balances'!A1")).toBe('Cash');
    expect(wb.getValue("'2025 Balances'!C2")).toBe(821400);
  });

  it('feeds a missing-sheet error back to the model and stages only the corrected, exactly named target', async () => {
    const wb = new WorkbookEmulator(['Payroll', 'Data']);
    const provider = new ScriptedProvider([
      {
        toolCalls: [
          { name: 'write_range', args: { range: 'Report!A1', cells: [[999]], reason: 'guess' } },
        ],
      },
      {
        toolCalls: [
          { name: 'write_range', args: { range: 'data!A1', cells: [[1]], reason: 'corrected' } },
        ],
      },
      { text: 'Staged on Data.' },
    ]);

    const { events, changeSets, stagedId } = await runScript(wb, provider);

    const firstResult = provider.requests.at(1)?.messages.findLast((m) => m.role === 'tool');
    expect(firstResult && 'content' in firstResult ? firstResult.content : '').toContain(
      'No such sheet: Report. Sheets in this workbook: Payroll, Data',
    );
    const staged = events.find((e) => e.kind === 'changeset_staged');
    expect(staged?.kind === 'changeset_staged' && staged.preview.map((p) => p.op)).toEqual([
      { kind: 'write_range', range: 'Data!A1', cells: [[{ value: 1 }]], reason: 'corrected' },
    ]);
    expect((await changeSets.apply(stagedId!)).ok).toBe(true);
    expect(wb.getValue('Data!A1')).toBe(1);
    expect(wb.getValue('Payroll!A1')).toBeNull();
  });
});
