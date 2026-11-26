import ExcelJS from 'exceljs';
import {
  extractCellPlainText,
  inspectWorkbook,
  prepareWorksheetTranslationPlan
} from '../client/xlsx-core.js';
import { translateBatch } from './translators/service.js';

export { extractCellPlainText };

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

  const plan = prepareWorksheetTranslationPlan(workbook, {
    sheetName,
    sourceAddress,
    targetMode,
    customTargetAddress
  });

  const batchResults = await translateBatch({
    texts: plan.nonEmptyTasks.map((t) => t.rawText),
    engine,
    from,
    to,
    config,
    fetchImpl
  });

  let successCount = 0;
  const errors = [];

  plan.nonEmptyTasks.forEach((task, index) => {
    const res = batchResults[index];
    if (res && res.ok) {
      plan.targetWorksheet.getCell(task.dstRow, task.dstCol).value = res.translatedText;
      successCount += 1;
    } else {
      errors.push(res?.error || 'Unknown error');
    }
  });

  if (successCount === 0 && errors.length > 0) {
    throw new Error(errors[0]);
  }

  const outArrayBuffer = await workbook.xlsx.writeBuffer();
  return {
    outputBuffer: Buffer.from(outArrayBuffer),
    sheetNames: workbook.worksheets.map((ws) => ws.name),
    sourceSheet: plan.sourceWorksheet.name,
    targetSheet: plan.targetWorksheet.name,
    sourceRange: plan.sourceRange,
    targetRange: plan.targetRange,
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
  return inspectWorkbook(workbook);
}
