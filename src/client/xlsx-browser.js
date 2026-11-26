import { apiFetch } from './api-client.js';
import { inspectWorkbook, prepareWorksheetTranslationPlan } from './xlsx-core.js';

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

export async function inspectXlsxClient(arrayBuffer, options = {}) {
  const ExcelJS = await ensureBrowserExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  return inspectWorkbook(workbook, options);
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

  const plan = prepareWorksheetTranslationPlan(workbook, {
    sheetName,
    sourceAddress,
    targetMode,
    customTargetAddress
  });

  const CHUNK_SIZE = 25;
  let translatedCount = 0;
  const errors = [];

  for (let start = 0; start < plan.nonEmptyTasks.length; start += CHUNK_SIZE) {
    const chunk = plan.nonEmptyTasks.slice(start, start + CHUNK_SIZE);
    if (typeof onProgress === 'function') {
      onProgress(
        Math.min(start + chunk.length, plan.nonEmptyTasks.length),
        plan.nonEmptyTasks.length
      );
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
        plan.targetWorksheet.getCell(task.dstRow, task.dstCol).value = item.translatedText;
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
  return {
    blob: new Blob([outBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    sourceRange: plan.sourceRange,
    targetRange: plan.targetRange,
    translatedCount,
    failedCount: errors.length
  };
}
