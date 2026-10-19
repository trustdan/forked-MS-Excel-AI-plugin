import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookEmulator } from '../src/workbook/emulator.js';
import { ChangeSetManager } from '../src/changeset/manager.js';
import { ScriptedProvider } from '../src/providers/scripted.js';
import { runAgentTurn, type AgentEvent } from '../src/agent/loop.js';
import { openDb } from '../src/db.js';
import { startServer, type SidecarServer } from '../src/server.js';

describe('import_local_data tool integration', () => {
  let tempDir: string;
  let sidecar: SidecarServer | null = null;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-import-tool-${crypto.randomUUID()}`);
    mkdirSync(tempDir, { recursive: true });
  });

  afterEach(async () => {
    await sidecar?.close();
    sidecar = null;
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  it('stages local CSV file import into workbook and applies cleanly upon approval', async () => {
    const csvFile = join(tempDir, 'sales.csv');
    writeFileSync(csvFile, 'Department,Q1,Q2\nHardware,12000,15000\nSoftware,45000,52000');

    const wb = new WorkbookEmulator(['Sheet1']);
    const changeSets = new ChangeSetManager(wb, 'wb-test');
    const events: AgentEvent[] = [];

    const provider = new ScriptedProvider([
      {
        toolCalls: [
          {
            name: 'import_local_data',
            args: {
              filePath: csvFile,
              targetSheet: 'Sheet1',
              startCell: 'A1',
              hasHeaders: true,
              reason: 'Import quarterly department sales',
            },
          },
        ],
      },
      { text: 'I staged the imported sales data starting at Sheet1!A1.' },
    ]);

    const result = await runAgentTurn({
      provider,
      executor: wb,
      changeSets,
      workbookName: 'Sales.xlsx',
      history: [],
      userMessage: 'Import sales.csv into Sheet1',
      onEvent: (ev) => events.push(ev),
    });

    expect(result.stagedChangeSetId).toBeDefined();

    // Verify cell untouched before approval
    expect(wb.getValue('Sheet1!A1')).toBeNull();

    // Verify events
    const toolEvent = events.find((e) => e.kind === 'tool_use' && e.tool === 'import_local_data');
    expect(toolEvent).toBeDefined();

    const stagedEvent = events.find((e) => e.kind === 'changeset_staged');
    expect(stagedEvent).toBeDefined();

    // Now approve the staged change-set
    const changeSet = changeSets.get(result.stagedChangeSetId!);
    expect(changeSet).toBeDefined();
    await wb.applyOps(changeSet!.ops);

    // Verify imported values in emulator
    expect(wb.getValue('Sheet1!A1')).toBe('Department');
    expect(wb.getValue('Sheet1!B1')).toBe('Q1');
    expect(wb.getValue('Sheet1!C1')).toBe('Q2');
    expect(wb.getValue('Sheet1!A2')).toBe('Hardware');
    expect(wb.getValue('Sheet1!B2')).toBe(12000);
    expect(wb.getValue('Sheet1!C2')).toBe(15000);
    expect(wb.getValue('Sheet1!A3')).toBe('Software');
    expect(wb.getValue('Sheet1!B3')).toBe(45000);
    expect(wb.getValue('Sheet1!C3')).toBe(52000);
  });

  it('serves /api/importer/parse endpoint with file path and raw content', async () => {
    sidecar = await startServer({ db: openDb(':memory:'), port: 0 });
    const base = `http://127.0.0.1:${sidecar.port}`;
    const code = sidecar.auth.currentPairingCode!;
    const pairRes = await fetch(`${base}/api/pair`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ code }),
    });
    const { token } = (await pairRes.json()) as { token: string };
    const authHeaders = {
      authorization: `Bearer ${token}`,
      'content-type': 'application/json',
    };

    // Test 1: parse local CSV file
    const sampleCsv = join(tempDir, 'metrics.csv');
    writeFileSync(sampleCsv, 'Metric,Value\nCPU,45.2\nMemory,78.1');

    const fileRes = await fetch(`${base}/api/importer/parse`, {
      method: 'POST',
      headers: authHeaders,
      body: JSON.stringify({ filePath: sampleCsv }),
    });
    expect(fileRes.status).toBe(200);
    const fileBody = (await fileRes.json()) as {
      ok: boolean;
      format: string;
      totalRows: number;
      totalCols: number;
      headers: string[];
      previewRows: unknown[][];
    };
    expect(fileBody.ok).toBe(true);
    expect(fileBody.totalRows).toBe(3);
    expect(fileBody.headers).toEqual(['Metric', 'Value']);

    // Test 2: parse raw markdown table string
    const mdTable = `| Name | Role |\n|---|---|\n| Ada | Pioneer |\n| Alan | Cryptanalyst |`;
    const mdRes = await fetch(`${base}/api/importer/parse`, {
      method: 'POST',
      headers: authHeaders,
      body: JSON.stringify({ content: mdTable, format: 'markdown' }),
    });
    expect(mdRes.status).toBe(200);
    const mdBody = (await mdRes.json()) as {
      ok: boolean;
      totalRows: number;
      headers: string[];
    };
    expect(mdBody.ok).toBe(true);
    expect(mdBody.totalRows).toBe(3);
    expect(mdBody.headers).toEqual(['Name', 'Role']);
  });
});
