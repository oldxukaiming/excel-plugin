import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import {
  columnLettersToIndex,
  computeTargetAddress,
  indexToColumnLetters,
  parseRangeAddress,
  splitSheetAndRange
} from '../src/client/range-utils.js';

describe('range-utils', () => {
  it('converts column letters and indices symmetrically', () => {
    assert.equal(columnLettersToIndex('A'), 1);
    assert.equal(columnLettersToIndex('Z'), 26);
    assert.equal(columnLettersToIndex('AA'), 27);
    assert.equal(columnLettersToIndex('AZ'), 52);

    assert.equal(indexToColumnLetters(1), 'A');
    assert.equal(indexToColumnLetters(26), 'Z');
    assert.equal(indexToColumnLetters(27), 'AA');
    assert.equal(indexToColumnLetters(52), 'AZ');
  });

  it('computes below target address (A1 -> A2)', () => {
    assert.equal(computeTargetAddress('A1', 'below'), 'A2');
    assert.equal(computeTargetAddress('A1:C1', 'below'), 'A2:C2');
    assert.equal(computeTargetAddress('A1:A3', 'below'), 'A4:A6');
    assert.equal(computeTargetAddress('Sheet1!$B$2', 'below'), 'Sheet1!B3');
  });

  it('computes right, replace, and custom target addresses', () => {
    assert.equal(computeTargetAddress('A1', 'right'), 'B1');
    assert.equal(computeTargetAddress('A1:B3', 'right'), 'C1:D3');
    assert.equal(computeTargetAddress('A1:B2', 'replace'), 'A1:B2');
    assert.equal(computeTargetAddress('A1:B2', 'custom', 'D5'), 'D5:E6');
    assert.equal(computeTargetAddress('A1:B2', 'custom', 'Sheet2!D5'), 'Sheet2!D5:E6');
  });

  it('splits sheet name and range address properly', () => {
    assert.deepEqual(splitSheetAndRange('A1'), { sheetName: null, rangeAddress: 'A1' });
    assert.deepEqual(splitSheetAndRange('Sheet1!A1:B2'), {
      sheetName: 'Sheet1',
      rangeAddress: 'A1:B2'
    });
    assert.deepEqual(splitSheetAndRange("'My ''Data'''!C3"), {
      sheetName: "My 'Data'",
      rangeAddress: 'C3'
    });
  });

  it('parses reversed bounds cleanly', () => {
    const parsed = parseRangeAddress('C5:A2');
    assert.equal(parsed.startRow, 2);
    assert.equal(parsed.endRow, 5);
    assert.equal(parsed.startCol, 1);
    assert.equal(parsed.endCol, 3);
    assert.equal(parsed.rowCount, 4);
    assert.equal(parsed.colCount, 3);
  });
});
