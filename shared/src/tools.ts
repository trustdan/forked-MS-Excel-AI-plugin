import { z } from 'zod';
import { parseA1, rangeHeight, rangeWidth } from './cells.js';

/**
 * The workbook tool surface. These schemas are the single source of truth:
 * the model sees them (as JSON schema), the sidecar validates against them,
 * and both the Office.js bridge and the headless emulator implement them.
 */

const cellScalar = z.union([z.string(), z.number(), z.boolean(), z.null()]);

/** A cell write: either a literal value or a canonical en-US formula. */
const cellWrite = z.union([
  cellScalar,
  z.object({ formula: z.string().startsWith('=') }).strict(),
]);

export const toolSchemas = {
  get_workbook_map: {
    description:
      'Get the structural map of the workbook: sheets, their used ranges and dimensions. Call this only when the workbook sheets are not already provided in the system prompt.',
    parameters: z.object({}).strict(),
    readonly: true,
  },
  read_range: {
    description:
      'Read values and formulas of a range. Address is A1 style with sheet, e.g. "Sheet1!A1:C20". Keep ranges small and targeted; never read entire sheets blindly.',
    parameters: z
      .object({
        range: z.string().min(1).describe('A1-style range, e.g. "Sheet1!A1:C20"'),
      })
      .strict(),
    readonly: true,
  },
  write_range: {
    description:
      'Write a 2-D array of rows starting at a top-left cell (e.g. "Sheet1!A1"). Example: `{"range": "Sheet1!A1", "cells": [["ASSETS"], ["Cash", 401900], ["Raw Materials", 419500], ["Total", {"formula": "=SUM(B2:B3)"}]], "reason": "Add balance sheet"}`. Pass just the top-left cell like "Sheet1!A1" — the end row/column and padding for 1-cell header rows (`["ASSETS"]`) are computed automatically. Do NOT triple-nest rows (`[[["Cash", 100]]]` is wrong; use `[["Cash", 100]]`). Staged into a change-set for review.',
    parameters: z
      .object({
        range: z
          .string()
          .min(1)
          .describe('Top-left A1 cell or target range, e.g. "Sheet1!A1" or "Sheet1!A1:B20"'),
        cells: z.array(z.array(cellWrite)).min(1),
        reason: z.string().min(1).max(300).describe('One line: why this change is being made'),
      })
      .strict()
      .superRefine((val, ctx) => {
        let addr;
        try {
          addr = parseA1(val.range);
        } catch (e) {
          ctx.addIssue({
            code: z.ZodIssueCode.custom,
            path: ['range'],
            message: e instanceof Error ? e.message : `Invalid A1 range: ${val.range}`,
          });
          return;
        }
        const expectedRows = rangeHeight(addr);
        const expectedCols = rangeWidth(addr);
        if (expectedCols > 1) {
          for (let r = 0; r < val.cells.length; r++) {
            const row = val.cells[r]!;
            if (row.length > expectedCols || (val.cells.length === 1 && row.length !== expectedCols)) {
              ctx.addIssue({
                code: z.ZodIssueCode.custom,
                path: ['cells', r],
                message: `Row ${r + 1} has ${row.length} column(s), expected ${expectedCols} for range "${val.range}" (${expectedRows}×${expectedCols})`,
              });
            }
          }
        }
      }),
    readonly: false,
  },
  format_range: {
    description:
      'Apply formatting to a range: number format (e.g. "#,##0.00", "0%", "mm/dd/yyyy"), bold, italic, font size, font/fill colors (#RRGGBB), horizontal alignment, text wrapping, borders, column autofit. Staged into the change-set.',
    parameters: z
      .object({
        range: z.string().min(1).describe('A1-style range, e.g. "Sheet1!A1:C1"'),
        format: z
          .object({
            numberFormat: z.string().min(1).optional(),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            fontSize: z.number().min(6).max(72).optional(),
            fontColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
            fillColor: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
            horizontalAlignment: z.enum(['left', 'center', 'right']).optional(),
            wrapText: z.boolean().optional(),
            border: z
              .object({
                edges: z
                  .array(
                    z.enum([
                      'all',
                      'outline',
                      'top',
                      'bottom',
                      'left',
                      'right',
                      'insideHorizontal',
                      'insideVertical',
                    ]),
                  )
                  .min(1)
                  .describe('"all" = outline + inside grid lines; "outline" = box around the range'),
                style: z.enum(['thin', 'medium', 'thick', 'double']).optional(),
                color: z.string().regex(/^#[0-9A-Fa-f]{6}$/).optional(),
              })
              .strict()
              .optional(),
            autofitColumns: z.boolean().optional(),
          })
          .strict()
          .refine((f) => Object.keys(f).length > 0, 'format must set at least one property'),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  insert_rows: {
    description:
      'Insert blank rows, shifting existing rows down. `at` is the 1-based row number where insertion starts. Staged into the change-set.',
    parameters: z
      .object({
        sheet: z.string().min(1),
        at: z.number().int().min(1),
        count: z.number().int().min(1).max(1000).default(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  delete_rows: {
    description:
      'Delete rows, shifting the rows below up. `at` is the 1-based first row to delete. Staged into the change-set.',
    parameters: z
      .object({
        sheet: z.string().min(1),
        at: z.number().int().min(1),
        count: z.number().int().min(1).max(1000).default(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  insert_cols: {
    description:
      'Insert blank columns, shifting existing columns right. `at` is the column letter where insertion starts (e.g. "B"). Staged into the change-set.',
    parameters: z
      .object({
        sheet: z.string().min(1),
        at: z.string().regex(/^[A-Za-z]{1,3}$/),
        count: z.number().int().min(1).max(100).default(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  delete_cols: {
    description:
      'Delete columns, shifting the columns to the right left. `at` is the first column letter to delete (e.g. "B"). Staged into the change-set.',
    parameters: z
      .object({
        sheet: z.string().min(1),
        at: z.string().regex(/^[A-Za-z]{1,3}$/),
        count: z.number().int().min(1).max(100).default(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  sort_range: {
    description:
      'Sort the rows of a range by one of its columns. `keyColumn` is a column letter that must lie inside the range. Set hasHeader if the first row is a header and should stay in place. Staged into the change-set.',
    parameters: z
      .object({
        range: z.string().min(1),
        keyColumn: z.string().regex(/^[A-Za-z]{1,3}$/),
        ascending: z.boolean().default(true),
        hasHeader: z.boolean().default(false),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  rename_sheet: {
    description: 'Rename a worksheet. Staged into the change-set.',
    parameters: z
      .object({
        name: z.string().min(1),
        newName: z.string().min(1).max(31),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  delete_sheet: {
    description:
      'Delete a worksheet and everything on it. Destructive — use only when the user clearly asked for it. Staged into the change-set.',
    parameters: z
      .object({
        name: z.string().min(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  find: {
    description:
      'Search the workbook (or one sheet) for cells whose value contains the query text (case-insensitive). Returns up to 50 matches with addresses — use it to locate data instead of reading large ranges.',
    parameters: z
      .object({
        query: z.string().min(1),
        sheet: z.string().min(1).optional(),
      })
      .strict(),
    readonly: true,
  },
  add_sheet: {
    description: 'Create a new worksheet with the given name. Staged into the change-set.',
    parameters: z
      .object({
        name: z.string().min(1).max(31),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  clear_range: {
    description: 'Clear values and formulas in a range. Staged into the change-set.',
    parameters: z
      .object({
        range: z.string().min(1),
        reason: z.string().min(1).max(300),
      })
      .strict(),
    readonly: false,
  },
  import_local_data: {
    description:
      'Parse a local CSV, TSV, JSON, or Markdown file and stage its tabular data into an Excel sheet. Staged into the change-set for user review.',
    parameters: z
      .object({
        filePath: z.string().min(1).describe('Absolute path or filename of the CSV/TSV/Markdown file to import'),
        targetSheet: z.string().min(1).optional().describe('Target sheet name (defaults to active sheet)'),
        startCell: z.string().optional().describe('Top-left cell address (e.g. "A1", defaults to "A1")'),
        hasHeaders: z.boolean().optional().describe('Whether the first row is a header (defaults to true)'),
        reason: z.string().min(1).max(300).describe('One line: why this data is being imported'),
      })
      .strict(),
    readonly: false,
  },
} as const;

export type ToolName = keyof typeof toolSchemas;

export const toolNames = Object.keys(toolSchemas) as ToolName[];

export type ToolParams<T extends ToolName> = z.infer<(typeof toolSchemas)[T]['parameters']>;

export function isWriteTool(name: ToolName): boolean {
  return !toolSchemas[name].readonly;
}

export interface ToolCallRequest {
  name: string;
  /** Raw, unvalidated arguments as produced by the model. */
  args: unknown;
}

export interface SheetInfo {
  name: string;
  rowCount: number;
  colCount: number;
  usedRange: string | null;
}

export interface WorkbookMap {
  sheets: SheetInfo[];
}
