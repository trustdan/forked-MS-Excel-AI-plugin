import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  generateImportChangeOps,
  inferCellValue,
  loadAndParseLocalFile,
  parseDelimited,
  parseJsonTable,
  parseMarkdownTable,
  parseTabularContent,
  sanitizeFormulaInjection,
} from '../src/importers/tabular.js';
import { McpSecurityError } from '../src/mcp/pathGuard.js';

describe('Local Tabular Importer (FEAT-02)', () => {
  let tempDir: string;

  beforeEach(() => {
    tempDir = join(tmpdir(), `excelai-importer-${crypto.randomUUID()}`);
    mkdirSync(tempDir, { recursive: true });
  });

  afterEach(() => {
    try {
      rmSync(tempDir, { recursive: true, force: true });
    } catch {
      // ignore
    }
  });

  describe('CSV / TSV Parsing (RFC 4180)', () => {
    it('parses standard comma-delimited data', () => {
      const csv = 'Name,Age,City\nAlice,30,New York\nBob,25,Chicago';
      const rows = parseDelimited(csv, ',');
      expect(rows).toEqual([
        ['Name', 'Age', 'City'],
        ['Alice', '30', 'New York'],
        ['Bob', '25', 'Chicago'],
      ]);
    });

    it('handles quoted fields with commas and escaped quotes', () => {
      const csv = '"Company, Inc.","Product ""Pro""",100\nAcme,"Standard",50';
      const rows = parseDelimited(csv, ',');
      expect(rows).toEqual([
        ['Company, Inc.', 'Product "Pro"', '100'],
        ['Acme', 'Standard', '50'],
      ]);
    });

    it('handles multiline quoted fields and CRLF line breaks', () => {
      const csv = 'Header1,Header2\r\n"Line 1\nLine 2",Value2\r\n';
      const rows = parseDelimited(csv, ',');
      expect(rows).toEqual([
        ['Header1', 'Header2'],
        ['Line 1\nLine 2', 'Value2'],
      ]);
    });

    it('parses tab-delimited (TSV) and semicolon-delimited data', () => {
      const tsv = 'A\tB\tC\n1\t2\t3';
      expect(parseDelimited(tsv, '\t')).toEqual([
        ['A', 'B', 'C'],
        ['1', '2', '3'],
      ]);

      const semi = 'X;Y;Z\n10;20;30';
      expect(parseDelimited(semi, ';')).toEqual([
        ['X', 'Y', 'Z'],
        ['10', '20', '30'],
      ]);
    });
  });

  describe('Markdown Table Parsing', () => {
    it('extracts table rows and skips separator lines', () => {
      const md = `
# Sales Report
Here is the data:
| Region | Q1 | Q2 |
| :----- | -: | -: |
| North  | 100 | 150 |
| South  | 200 | 250 |

Footer text
`;
      const rows = parseMarkdownTable(md);
      expect(rows).toEqual([
        ['Region', 'Q1', 'Q2'],
        ['North', '100', '150'],
        ['South', '200', '250'],
      ]);
    });
  });

  describe('JSON Table Parsing', () => {
    it('parses array of objects into headers and rows', () => {
      const json = JSON.stringify([
        { item: 'Widget', qty: 5, price: 19.99 },
        { item: 'Gadget', qty: 2, price: 49.99 },
      ]);
      const rows = parseJsonTable(json);
      expect(rows).toEqual([
        ['item', 'qty', 'price'],
        ['Widget', '5', '19.99'],
        ['Gadget', '2', '49.99'],
      ]);
    });

    it('parses array of arrays', () => {
      const json = JSON.stringify([
        ['Col1', 'Col2'],
        ['Val1', 'Val2'],
      ]);
      const rows = parseJsonTable(json);
      expect(rows).toEqual([
        ['Col1', 'Col2'],
        ['Val1', 'Val2'],
      ]);
    });
  });

  describe('Formula Injection Defense (CWE-1236)', () => {
    it('escapes dangerous leading formula characters with single quote', () => {
      expect(sanitizeFormulaInjection("=cmd|' /C calc'!A0")).toBe("'=cmd|' /C calc'!A0");
      expect(sanitizeFormulaInjection('@SUM(1+1)')).toBe("'@SUM(1+1)");
      expect(sanitizeFormulaInjection('+1234-cmd')).toBe("'+1234-cmd");
      expect(sanitizeFormulaInjection('-DDE("cmd";"calc")')).toBe("'-DDE(\"cmd\";\"calc\")");
    });

    it('does not escape valid signed numeric values', () => {
      expect(sanitizeFormulaInjection('-42')).toBe('-42');
      expect(sanitizeFormulaInjection('+3.14159')).toBe('+3.14159');
      expect(sanitizeFormulaInjection('-0.5e10')).toBe('-0.5e10');
    });

    it('infers typed cell scalars safely', () => {
      expect(inferCellValue('TRUE')).toBe(true);
      expect(inferCellValue('false')).toBe(false);
      expect(inferCellValue('1234')).toBe(1234);
      expect(inferCellValue('$1,250.50')).toBe(1250.5);
      expect(inferCellValue('25%')).toBe(0.25);
      expect(inferCellValue('   ')).toBeNull();
      expect(inferCellValue('=1+1', true)).toBe("'=1+1");
      expect(inferCellValue('Normal text')).toBe('Normal text');
    });
  });

  describe('File Loading & Security Validation', () => {
    it('loads and parses a valid local CSV file', () => {
      const file = join(tempDir, 'data.csv');
      writeFileSync(file, 'ID,Score\n101,95.5\n102,88');

      const result = loadAndParseLocalFile(file, [tempDir]);
      expect(result.format).toBe('csv');
      expect(result.totalRows).toBe(3);
      expect(result.totalCols).toBe(2);
      expect(result.rows).toEqual([
        ['ID', 'Score'],
        [101, 95.5],
        [102, 88],
      ]);
    });

    it('rejects unsupported extensions', () => {
      const file = join(tempDir, 'script.exe');
      writeFileSync(file, 'binary-data');

      expect(() => loadAndParseLocalFile(file, [tempDir])).toThrow(/Unsupported file format/);
    });

    it('rejects non-existent files', () => {
      const file = join(tempDir, 'missing.csv');
      expect(() => loadAndParseLocalFile(file, [tempDir])).toThrow(/File not found/);
    });

    it('refuses files outside the approved folders before reading them (F9)', () => {
      const file = join(tempDir, 'private.csv');
      writeFileSync(file, 'Secret\nPRIVATE-MARKER');

      expect(() => loadAndParseLocalFile(file, [])).toThrow(McpSecurityError);
      expect(() => loadAndParseLocalFile(file, [join(tempDir, 'other')])).toThrow(McpSecurityError);
    });

    it('rejects files exceeding byte limit', () => {
      const file = join(tempDir, 'huge.csv');
      writeFileSync(file, 'A'.repeat(2048));

      expect(() => loadAndParseLocalFile(file, [tempDir], { maxBytes: 1024 })).toThrow(/exceeds limit/);
    });
  });

  describe('generateImportChangeOps', () => {
    it('generates write_range op with correct target dimensions and formatting', () => {
      const rows = [
        ['Product', 'Revenue'],
        ['Apples', 500],
        ['Oranges', 300],
      ];

      const ops = generateImportChangeOps({
        rows,
        targetSheet: 'Sheet1',
        startCell: 'B2',
        hasHeaders: true,
        reason: 'Import Q1 revenue',
      });

      expect(ops.length).toBe(2); // write_range + header format_range
      expect(ops[0]).toEqual({
        kind: 'write_range',
        range: 'Sheet1!B2:C4',
        cells: [
          [{ value: 'Product' }, { value: 'Revenue' }],
          [{ value: 'Apples' }, { value: 500 }],
          [{ value: 'Oranges' }, { value: 300 }],
        ],
        reason: 'Import Q1 revenue',
      });

      expect(ops[1]).toEqual({
        kind: 'format_range',
        range: 'Sheet1!B2:C2',
        format: {
          bold: true,
          border: {
            edges: ['bottom'],
            style: 'medium',
          },
        },
        reason: 'Format header row for imported data',
      });
    });
  });
});
