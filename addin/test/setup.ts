import { afterEach } from 'vitest';
import { cleanup } from '@testing-library/react';

afterEach(() => {
  cleanup();
});

// Mock Excel and Office globals for unit test environment
(globalThis as any).Excel = {
  BorderIndex: {
    edgeTop: 'EdgeTop',
    edgeBottom: 'EdgeBottom',
    edgeLeft: 'EdgeLeft',
    edgeRight: 'EdgeRight',
    insideHorizontal: 'InsideHorizontal',
    insideVertical: 'InsideVertical',
  },
  BorderLineStyle: {
    continuous: 'Continuous',
    double: 'Double',
  },
  BorderWeight: {
    thin: 'Thin',
    medium: 'Medium',
    thick: 'Thick',
  },
  HorizontalAlignment: {
    left: 'Left',
    center: 'Center',
    right: 'Right',
  },
  ClearApplyTo: {
    contents: 'Contents',
  },
  InsertShiftDirection: {
    down: 'Down',
    right: 'Right',
  },
  DeleteShiftDirection: {
    up: 'Up',
    left: 'Left',
  },
  SortOrientation: {
    rows: 'Rows',
  },
  run: async (callback: (ctx: any) => Promise<any>) => {
    return callback({
      workbook: {
        worksheets: {
          items: [],
          getActiveWorksheet: () => ({ getRange: () => ({}) }),
          getItem: () => ({ getRange: () => ({}) }),
        },
      },
      sync: async () => {},
    });
  },
};

(globalThis as any).Office = {
  context: {
    document: {
      settings: {
        get: () => null,
        set: () => {},
        saveAsync: (cb: () => void) => cb(),
      },
    },
  },
};
