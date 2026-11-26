import ExcelJS from 'exceljs';
import {
  computeTargetAddress,
  indexToColumnLetters,
  parseRangeAddress,
  splitSheetAndRange
} from '../client/range-utils.js';
import { translateBatch } from './translators/service.js';

/**
 * Extract plain displayable string from an ExcelJS cell value.
 */
export function extractCellPlainText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('');
    }
    if (value.text !== undefined) {
      return String(value.text);
    }
    if (value.result !== undefined) {
      return extractCellPlainText(value.result);
    }
  }
  return String(value);
}

/**
 * Translate a range in an .xlsx buffer and write results to the computed target range.
 */
export async function translateXlsxBuffer({
  buffer,
  sheetName = '',
  sourceAddress = 'A1',
  targetMode = 'below',
  customTargetAddress = '',
  engine = 'bing',
  from = 'auto',
  to = 'zh-CN',
  config = {},
  fetchImpl = fetch
}) {
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('Uploaded .xlsx file buffer is empty');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const srcSplit = splitSheetAndRange(sourceAddress);
  const resolvedSourceSheetName = srcSplit.sheetName || sheetName.trim();
  const sourceWorksheet = resolvedSourceSheetName
    ? workbook.getWorksheet(resolvedSourceSheetName)
    : workbook.worksheets[0];

  if (!sourceWorksheet) {
    throw new Error(
      resolvedSourceSheetName
        ? `Worksheet "${resolvedSourceSheetName}" not found`
        : 'No worksheets found in workbook'
    );
  }

  const srcCoords = parseRangeAddress(srcSplit.rangeAddress);
  const targetFullAddress = computeTargetAddress(
    sourceAddress,
    targetMode,
    customTargetAddress
  );
  const dstSplit = splitSheetAndRange(targetFullAddress);
  const targetWorksheet = dstSplit.sheetName
    ? workbook.getWorksheet(dstSplit.sheetName)
    : sourceWorksheet;

  if (!targetWorksheet) {
    throw new Error(`Target worksheet "${dstSplit.sheetName}" not found`);
  }

  const dstCoords = parseRangeAddress(dstSplit.rangeAddress);

  const tasks = [];
  for (let rOffset = 0; rOffset < srcCoords.rowCount; rOffset += 1) {
    for (let cOffset = 0; cOffset < srcCoords.colCount; cOffset += 1) {
      const srcRow = srcCoords.startRow + rOffset;
      const srcCol = srcCoords.startCol + cOffset;
      const dstRow = dstCoords.startRow + rOffset;
      const dstCol = dstCoords.startCol + cOffset;
      const cell = sourceWorksheet.getCell(srcRow, srcCol);
      const rawText = extractCellPlainText(cell.value);
      tasks.push({ srcRow, srcCol, dstRow, dstCol, rawText });
    }
  }

  const nonEmptyTasks = tasks.filter((t) => t.rawText.trim() !== '');
  if (nonEmptyTasks.length === 0) {
    // Look for nearby non-empty cells in the same start column to give an actionable hint
    const colLetter = indexToColumnLetters(srcCoords.startCol);
    let firstNonEmptyRow = null;
    let firstSample = '';
    for (let r = 1; r <= Math.min(sourceWorksheet.rowCount, 1000); r += 1) {
      const sample = extractCellPlainText(sourceWorksheet.getCell(r, srcCoords.startCol).value).trim();
      if (sample) {
        firstNonEmptyRow = r;
        firstSample = sample.slice(0, 40);
        break;
      }
    }
    const hint = firstNonEmptyRow
      ? `（该列首个非空单元格为 ${colLetter}${firstNonEmptyRow}: "${firstSample}"）`
      : '（该工作表此列暂无文本内容）';
    throw new Error(`源区域 ${sourceAddress} 为空，没有可翻译的文本${hint}`);
  }

  const batchResults = await translateBatch({
    texts: nonEmptyTasks.map((t) => t.rawText),
    engine,
    from,
    to,
    config,
    fetchImpl
  });

  let successCount = 0;
  const errors = [];

  nonEmptyTasks.forEach((task, index) => {
    const res = batchResults[index];
    if (res && res.ok) {
      targetWorksheet.getCell(task.dstRow, task.dstCol).value = res.translatedText;
      successCount += 1;
    } else {
      errors.push(res?.error || 'Unknown error');
    }
  });

  if (successCount === 0 && errors.length > 0) {
    throw new Error(errors[0]);
  }

  const outArrayBuffer = await workbook.xlsx.writeBuffer();
  const formattedTarget =
    dstCoords.rowCount === 1 && dstCoords.colCount === 1
      ? `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}`
      : `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}:${indexToColumnLetters(dstCoords.endCol)}${dstCoords.endRow}`;

  return {
    outputBuffer: Buffer.from(outArrayBuffer),
    sheetNames: workbook.worksheets.map((ws) => ws.name),
    sourceSheet: sourceWorksheet.name,
    targetSheet: targetWorksheet.name,
    sourceRange: srcSplit.rangeAddress,
    targetRange: formattedTarget,
    translatedCount: successCount,
    failedCount: errors.length,
    firstError: errors[0] || null
  };
}

/**
 * Inspect worksheet names, preview grid, and per-column non-empty ranges from an uploaded .xlsx buffer.
 */
export async function inspectXlsxBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return {
    sheets: workbook.worksheets.map((ws) => {
      const maxPreviewRows = Math.min(Math.max(ws.rowCount, 1), 8);
      const maxPreviewCols = Math.min(Math.max(ws.columnCount, 1), 8);
      const previewRows = [];
      for (let r = 1; r <= maxPreviewRows; r += 1) {
        const rowVals = [];
        for (let c = 1; c <= maxPreviewCols; c += 1) {
          const text = extractCellPlainText(ws.getCell(r, c).value).trim();
          rowVals.push(text.length > 60 ? `${text.slice(0, 60)}...` : text);
        }
        previewRows.push(rowVals);
      }

      const columnRanges = [];
      for (let c = 1; c <= maxPreviewCols; c += 1) {
        let firstRow = null;
        let lastRow = null;
        let nonEmptyCount = 0;
        for (let r = 1; r <= ws.rowCount; r += 1) {
          const text = extractCellPlainText(ws.getCell(r, c).value).trim();
          if (text) {
            if (firstRow === null) firstRow = r;
            lastRow = r;
            nonEmptyCount += 1;
          }
        }
        const colLetter = indexToColumnLetters(c);
        columnRanges.push({
          col: colLetter,
          firstRow,
          lastRow,
          nonEmptyCount,
          range: firstRow !== null ? `${colLetter}${firstRow}:${colLetter}${lastRow}` : null
        });
      }

      return {
        name: ws.name,
        rowCount: ws.rowCount,
        columnCount: ws.columnCount,
        previewRows,
        columnRanges
      };
    })
  };
}
