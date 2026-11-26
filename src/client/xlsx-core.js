import {
  computeTargetAddress,
  indexToColumnLetters,
  parseRangeAddress,
  splitSheetAndRange
} from './range-utils.js';

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
 * Inspect an ExcelJS Workbook for accurate non-empty dimensions,
 * per-column ranges, and preview grid without inflating counts from styled empty cells.
 */
export function inspectWorkbook(workbook, { maxPreviewRows = 20, maxPreviewCols = 50 } = {}) {
  return {
    sheets: workbook.worksheets.map((ws) => {
      let usedRowCount = 0;
      let usedColCount = 0;
      const colStats = new Map();

      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          const text = extractCellPlainText(cell.value).trim();
          if (!text) return;

          if (rowNumber > usedRowCount) usedRowCount = rowNumber;
          if (colNumber > usedColCount) usedColCount = colNumber;

          let stat = colStats.get(colNumber);
          if (!stat) {
            stat = { firstRow: rowNumber, lastRow: rowNumber, nonEmptyCount: 1 };
            colStats.set(colNumber, stat);
          } else {
            if (rowNumber < stat.firstRow) stat.firstRow = rowNumber;
            if (rowNumber > stat.lastRow) stat.lastRow = rowNumber;
            stat.nonEmptyCount += 1;
          }
        });
      });

      const effectiveRows = usedRowCount || ws.actualRowCount || 0;
      const effectiveCols = usedColCount || ws.actualColumnCount || 0;
      const previewRowLimit = Math.min(Math.max(effectiveRows, 1), maxPreviewRows);
      const previewColLimit = Math.min(Math.max(effectiveCols, 1), maxPreviewCols);

      const previewRows = [];
      for (let r = 1; r <= previewRowLimit; r += 1) {
        const rowVals = [];
        for (let c = 1; c <= previewColLimit; c += 1) {
          const text = extractCellPlainText(ws.getCell(r, c).value).trim();
          rowVals.push(text.length > 60 ? `${text.slice(0, 60)}...` : text);
        }
        previewRows.push(rowVals);
      }

      const columnRanges = [];
      for (let c = 1; c <= previewColLimit; c += 1) {
        const colLetter = indexToColumnLetters(c);
        const stat = colStats.get(c);
        columnRanges.push({
          col: colLetter,
          firstRow: stat ? stat.firstRow : null,
          lastRow: stat ? stat.lastRow : null,
          nonEmptyCount: stat ? stat.nonEmptyCount : 0,
          range: stat ? `${colLetter}${stat.firstRow}:${colLetter}${stat.lastRow}` : null
        });
      }

      return {
        name: ws.name,
        rowCount: effectiveRows,
        columnCount: effectiveCols,
        previewRows,
        columnRanges
      };
    })
  };
}

/**
 * Prepare cell extraction plan and empty-range hint shared between browser & server translators.
 */
export function prepareWorksheetTranslationPlan(workbook, { sheetName = '', sourceAddress = 'A1', targetMode = 'below', customTargetAddress = '' }) {
  const srcSplit = splitSheetAndRange(sourceAddress);
  const resolvedSourceSheet = srcSplit.sheetName || sheetName.trim();
  const sourceWorksheet = resolvedSourceSheet
    ? workbook.getWorksheet(resolvedSourceSheet)
    : workbook.worksheets[0];

  if (!sourceWorksheet) {
    throw new Error(
      resolvedSourceSheet ? `未找到工作表 "${resolvedSourceSheet}"` : '工作簿中没有工作表'
    );
  }

  const srcCoords = parseRangeAddress(srcSplit.rangeAddress);
  const targetFull = computeTargetAddress(sourceAddress, targetMode, customTargetAddress);
  const dstSplit = splitSheetAndRange(targetFull);
  const targetWorksheet = dstSplit.sheetName
    ? workbook.getWorksheet(dstSplit.sheetName)
    : sourceWorksheet;

  if (!targetWorksheet) {
    throw new Error(`未找到目标工作表 "${dstSplit.sheetName}"`);
  }

  const dstCoords = parseRangeAddress(dstSplit.rangeAddress);
  const tasks = [];

  for (let rOffset = 0; rOffset < srcCoords.rowCount; rOffset += 1) {
    for (let cOffset = 0; cOffset < srcCoords.colCount; cOffset += 1) {
      const srcRow = srcCoords.startRow + rOffset;
      const srcCol = srcCoords.startCol + cOffset;
      const dstRow = dstCoords.startRow + rOffset;
      const dstCol = dstCoords.startCol + cOffset;
      const rawText = extractCellPlainText(sourceWorksheet.getCell(srcRow, srcCol).value);
      tasks.push({ srcRow, srcCol, dstRow, dstCol, rawText });
    }
  }

  const nonEmptyTasks = tasks.filter((t) => t.rawText.trim() !== '');
  if (nonEmptyTasks.length === 0) {
    const colLetter = indexToColumnLetters(srcCoords.startCol);
    let firstNonEmptyRow = null;
    let firstSample = '';
    const scanLimit = Math.min(Math.max(sourceWorksheet.rowCount, 100), 1000);
    for (let r = 1; r <= scanLimit; r += 1) {
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

  const formattedTarget =
    dstCoords.rowCount === 1 && dstCoords.colCount === 1
      ? `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}`
      : `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}:${indexToColumnLetters(dstCoords.endCol)}${dstCoords.endRow}`;

  return {
    sourceWorksheet,
    targetWorksheet,
    sourceRange: srcSplit.rangeAddress,
    targetRange: formattedTarget,
    nonEmptyTasks
  };
}
