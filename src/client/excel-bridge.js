import { splitSheetAndRange } from './range-utils.js';

/**
 * Check whether code is running inside an active Excel Office.js host.
 */
export function isExcelHostReady() {
  return typeof globalThis.Excel !== 'undefined' && typeof globalThis.Excel.run === 'function';
}

function getTargetWorksheet(context, sheetName) {
  if (sheetName) {
    return context.workbook.worksheets.getItem(sheetName);
  }
  return context.workbook.worksheets.getActiveWorksheet();
}

/**
 * Strip active sheet prefix if address belongs to the active sheet for cleaner display.
 */
function formatDisplayAddress(fullAddress, activeSheetName) {
  const { sheetName, rangeAddress } = splitSheetAndRange(fullAddress);
  if (!sheetName || sheetName === activeSheetName) {
    return rangeAddress;
  }
  return fullAddress;
}

/**
 * Read the currently selected range in Excel along with its 2D values.
 * @returns {Promise<{ address: string, values: any[][] }>}
 */
export async function readCurrentSelection() {
  if (!isExcelHostReady()) {
    throw new Error('当前未连接到 Excel 宿主环境');
  }

  return globalThis.Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = context.workbook.getSelectedRange();
    sheet.load('name');
    range.load(['address', 'values']);
    await context.sync();

    return {
      address: formatDisplayAddress(range.address, sheet.name),
      values: range.values
    };
  });
}

/**
 * Read values from a specific A1-style address (e.g. "A1", "A1:B3", or "Sheet2!A1").
 * @param {string} address
 * @returns {Promise<{ address: string, values: any[][] }>}
 */
export async function readRangeValues(address) {
  if (!isExcelHostReady()) {
    throw new Error('当前未连接到 Excel 宿主环境');
  }

  const { sheetName, rangeAddress } = splitSheetAndRange(address);
  return globalThis.Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load('name');
    const sheet = getTargetWorksheet(context, sheetName);
    const range = sheet.getRange(rangeAddress);
    range.load(['address', 'values']);
    await context.sync();

    return {
      address: formatDisplayAddress(range.address, activeSheet.name),
      values: range.values
    };
  });
}

/**
 * Write a 2D array of translated strings into the target Excel range.
 * @param {string} targetAddress
 * @param {string[][]} values2D
 */
export async function writeRangeValues(targetAddress, values2D) {
  if (!isExcelHostReady()) {
    throw new Error('当前未连接到 Excel 宿主环境');
  }

  const { sheetName, rangeAddress } = splitSheetAndRange(targetAddress);
  return globalThis.Excel.run(async (context) => {
    const sheet = getTargetWorksheet(context, sheetName);
    const range = sheet.getRange(rangeAddress);
    range.values = values2D;
    range.format.autofitRows();
    await context.sync();
  });
}

/**
 * Register a selection change listener across the active workbook.
 * Returns an unsubscribe function.
 */
export async function subscribeSelectionChanged(callback) {
  if (!isExcelHostReady()) {
    return () => {};
  }

  let eventResult = null;
  await globalThis.Excel.run(async (context) => {
    eventResult = context.workbook.onSelectionChanged.add(async () => {
      try {
        const selection = await readCurrentSelection();
        callback(selection);
      } catch {
        // Ignore transient selection read failures while Excel is editing a cell
      }
    });
    await context.sync();
  });

  return async () => {
    if (!eventResult) return;
    await globalThis.Excel.run(eventResult.context, async (context) => {
      eventResult.remove();
      await context.sync();
    }).catch(() => {});
  };
}
