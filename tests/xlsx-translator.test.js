import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {
  extractCellPlainText,
  inspectXlsxBuffer,
  translateXlsxBuffer
} from '../src/server/xlsx-translator.js';
import { clearTranslationCache } from '../src/server/translators/service.js';

describe('xlsx-translator', () => {
  it('extracts plain text from primitive, richText, and formula cells', () => {
    assert.equal(extractCellPlainText('Hello'), 'Hello');
    assert.equal(extractCellPlainText(123), '123');
    assert.equal(
      extractCellPlainText({ richText: [{ text: 'Hello ' }, { text: 'World' }] }),
      'Hello World'
    );
    assert.equal(extractCellPlainText({ formula: 'A1', result: 'Cached' }), 'Cached');
  });

  it('translates A1 to A2 inside an .xlsx workbook without Excel COM or VBA', async () => {
    clearTranslationCache();

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.getCell('A1').value = 'Hello Excel';
    ws.getCell('B1').value = 'Good morning';
    const inputBuffer = Buffer.from(await wb.xlsx.writeBuffer());

    const inspectInfo = await inspectXlsxBuffer(inputBuffer);
    assert.equal(inspectInfo.sheets[0].name, 'Data');
    assert.equal(inspectInfo.sheets[0].columnRanges[0].range, 'A1:A1');
    assert.equal(inspectInfo.sheets[0].previewRows[0][0], 'Hello Excel');

    const mockFetch = async (url) => {
      if (url.includes('translate_a/single')) {
        const parsed = new URL(url);
        const q = parsed.searchParams.get('q');
        const map = {
          'Hello Excel': '你好 Excel',
          'Good morning': '早上好'
        };
        return new Response(JSON.stringify([[[map[q] || q, q]], null, 'en']), {
          status: 200,
          headers: { 'Content-Type': 'application/json' }
        });
      }
      throw new Error(`Unexpected URL: ${url}`);
    };

    const result = await translateXlsxBuffer({
      buffer: inputBuffer,
      sourceAddress: 'A1:B1',
      targetMode: 'below',
      engine: 'google',
      from: 'auto',
      to: 'zh-CN',
      fetchImpl: mockFetch
    });

    assert.equal(result.targetRange, 'A2:B2');
    assert.equal(result.translatedCount, 2);
    assert.equal(result.failedCount, 0);

    const outWb = new ExcelJS.Workbook();
    await outWb.xlsx.load(result.outputBuffer);
    const outWs = outWb.getWorksheet('Data');
    assert.equal(outWs.getCell('A1').value, 'Hello Excel');
    assert.equal(outWs.getCell('B1').value, 'Good morning');
    assert.equal(outWs.getCell('A2').value, '你好 Excel');
    assert.equal(outWs.getCell('B2').value, '早上好');
  });

  it('provides actionable column hint when selected cell is empty', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('B2').value = '1 Introduction';
    const inputBuffer = Buffer.from(await wb.xlsx.writeBuffer());

    await assert.rejects(
      () =>
        translateXlsxBuffer({
          buffer: inputBuffer,
          sourceAddress: 'B1',
          targetMode: 'below'
        }),
      /B2: "1 Introduction"/
    );
  });

  it('counts actual non-empty rows and columns accurately across wide sheets (> 8 columns)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Wide');
    for (let c = 1; c <= 12; c += 1) {
      ws.getCell(1, c).value = `Col${c}`;
    }
    // Add an empty styled row far below to simulate Excel formatted blank rows
    ws.getCell(500, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    const inputBuffer = Buffer.from(await wb.xlsx.writeBuffer());
    const info = await inspectXlsxBuffer(inputBuffer);
    assert.equal(info.sheets[0].rowCount, 1);
    assert.equal(info.sheets[0].columnCount, 12);
    assert.equal(info.sheets[0].columnRanges.length, 12);
    assert.equal(info.sheets[0].columnRanges[11].col, 'L');
    assert.equal(info.sheets[0].previewRows[0].length, 12);
  });
});
