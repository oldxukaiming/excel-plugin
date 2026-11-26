import { apiFetch } from './api-client.js';
import {
  computeTargetAddress,
  indexToColumnLetters,
  parseRangeAddress,
  splitSheetAndRange
} from './range-utils.js';

let excelJsLoadPromise = null;

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  });
}

export async function ensureBrowserExcelJS() {
  if (globalThis.ExcelJS) {
    return globalThis.ExcelJS;
  }
  if (!excelJsLoadPromise) {
    excelJsLoadPromise = (async () => {
      try {
        await loadScript('./vendor/exceljs.min.js');
      } catch {
        await loadScript('https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js');
      }
      if (!globalThis.ExcelJS) {
        throw new Error('ExcelJS browser library unavailable');
      }
      return globalThis.ExcelJS;
    })().catch((err) => {
      excelJsLoadPromise = null;
      throw err;
    });
  }
  return excelJsLoadPromise;
}

function extractPlainText(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || '').join('');
    }
    if (value.text !== undefined) return String(value.text);
    if (value.result !== undefined) return extractPlainText(value.result);
  }
  return String(value);
}

export async function inspectXlsxClient(arrayBuffer) {
  const ExcelJS = await ensureBrowserExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  return {
    sheets: workbook.worksheets.map((ws) => {
      const maxPreviewRows = Math.min(Math.max(ws.rowCount, 1), 8);
      const maxPreviewCols = Math.min(Math.max(ws.columnCount, 1), 8);
      const previewRows = [];

      for (let r = 1; r <= maxPreviewRows; r += 1) {
        const rowVals = [];
        for (let c = 1; c <= maxPreviewCols; c += 1) {
          const text = extractPlainText(ws.getCell(r, c).value).trim();
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
          const text = extractPlainText(ws.getCell(r, c).value).trim();
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

/**
 * Translate .xlsx entirely in browser memory and send only text batches to /api/translate/batch.
 */
export async function translateXlsxClient({
  arrayBuffer,
  sheetName = '',
  sourceAddress = 'A1',
  targetMode = 'below',
  customTargetAddress = '',
  engine = 'bing',
  from = 'auto',
  to = 'zh-CN',
  settings = {},
  onProgress
}) {
  const ExcelJS = await ensureBrowserExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const srcSplit = splitSheetAndRange(sourceAddress);
  const resolvedSourceSheet = srcSplit.sheetName || sheetName.trim();
  const sourceWorksheet = resolvedSourceSheet
    ? workbook.getWorksheet(resolvedSourceSheet)
    : workbook.worksheets[0];

  if (!sourceWorksheet) {
    throw new Error(resolvedSourceSheet ? `未找到工作表 "${resolvedSourceSheet}"` : '工作簿中没有工作表');
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
      const rawText = extractPlainText(sourceWorksheet.getCell(srcRow, srcCol).value);
      tasks.push({ dstRow, dstCol, rawText });
    }
  }

  const nonEmptyTasks = tasks.filter((t) => t.rawText.trim() !== '');
  if (nonEmptyTasks.length === 0) {
    const colLetter = indexToColumnLetters(srcCoords.startCol);
    let firstNonEmptyRow = null;
    let firstSample = '';
    for (let r = 1; r <= Math.min(sourceWorksheet.rowCount, 1000); r += 1) {
      const sample = extractPlainText(sourceWorksheet.getCell(r, srcCoords.startCol).value).trim();
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

  // Chunk requests in groups of 25 cells to stay well within Worker subrequest/time limits
  const CHUNK_SIZE = 25;
  let translatedCount = 0;
  const errors = [];

  for (let start = 0; start < nonEmptyTasks.length; start += CHUNK_SIZE) {
    const chunk = nonEmptyTasks.slice(start, start + CHUNK_SIZE);
    if (typeof onProgress === 'function') {
      onProgress(Math.min(start + chunk.length, nonEmptyTasks.length), nonEmptyTasks.length);
    }

    const res = await apiFetch(
      '/api/translate/batch',
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          texts: chunk.map((t) => t.rawText),
          engine,
          from,
          to,
          config: settings
        })
      },
      settings
    );

    const data = await res.json();
    if (!res.ok || !data.ok) {
      throw new Error(data.error || `HTTP ${res.status}`);
    }

    chunk.forEach((task, idx) => {
      const item = data.results[idx];
      if (item && item.ok) {
        targetWorksheet.getCell(task.dstRow, task.dstCol).value = item.translatedText;
        translatedCount += 1;
      } else {
        errors.push(item?.error || 'Translation failed');
      }
    });
  }

  if (translatedCount === 0 && errors.length > 0) {
    throw new Error(errors[0]);
  }

  const outBuffer = await workbook.xlsx.writeBuffer();
  const formattedTarget =
    dstCoords.rowCount === 1 && dstCoords.colCount === 1
      ? `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}`
      : `${indexToColumnLetters(dstCoords.startCol)}${dstCoords.startRow}:${indexToColumnLetters(dstCoords.endCol)}${dstCoords.endRow}`;

  return {
    blob: new Blob([outBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    sourceRange: srcSplit.rangeAddress,
    targetRange: formattedTarget,
    translatedCount,
    failedCount: errors.length
  };
}
