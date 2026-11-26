/**
 * Pure utility functions for parsing Excel A1-style range addresses
 * and computing target output addresses (e.g. A1 -> A2).
 */

/**
 * Convert Excel column letters (e.g. "A", "AA") to 1-based column index.
 * @param {string} colStr
 * @returns {number}
 */
export function columnLettersToIndex(colStr) {
  const upper = colStr.toUpperCase();
  let result = 0;
  for (let i = 0; i < upper.length; i += 1) {
    result = result * 26 + (upper.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Convert 1-based column index to Excel column letters (e.g. 1 -> "A", 27 -> "AA").
 * @param {number} index
 * @returns {string}
 */
export function indexToColumnLetters(index) {
  if (index < 1) {
    throw new Error(`Invalid column index: ${index}`);
  }
  let n = index;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * Parse an A1-style range address such as "A1", "Sheet1!A1:B3", or "'My Sheet'!$A$1:$B$3".
 * @param {string} address
 */
export function parseRangeAddress(address) {
  if (!address || typeof address !== 'string') {
    throw new Error('Range address must be a non-empty string');
  }

  let sheetPrefix = '';
  let cellPart = address.trim();

  const bangIndex = cellPart.lastIndexOf('!');
  if (bangIndex !== -1) {
    sheetPrefix = cellPart.slice(0, bangIndex + 1);
    cellPart = cellPart.slice(bangIndex + 1);
  }

  const parts = cellPart.split(':');
  if (parts.length > 2) {
    throw new Error(`Invalid range address: ${address}`);
  }

  const parseCell = (token) => {
    const match = token.trim().match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/);
    if (!match) {
      throw new Error(`Invalid cell reference: ${token}`);
    }
    return {
      col: columnLettersToIndex(match[1]),
      row: Number.parseInt(match[2], 10)
    };
  };

  const start = parseCell(parts[0]);
  const end = parts[1] ? parseCell(parts[1]) : { ...start };

  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);

  return {
    sheetPrefix,
    startRow,
    startCol,
    endRow,
    endCol,
    rowCount: endRow - startRow + 1,
    colCount: endCol - startCol + 1
  };
}

/**
 * Format parsed coordinates back into an Excel range address.
 */
export function formatRangeAddress({ sheetPrefix = '', startRow, startCol, endRow, endCol }) {
  const startCell = `${indexToColumnLetters(startCol)}${startRow}`;
  if (startRow === endRow && startCol === endCol) {
    return `${sheetPrefix}${startCell}`;
  }
  const endCell = `${indexToColumnLetters(endCol)}${endRow}`;
  return `${sheetPrefix}${startCell}:${endCell}`;
}

/**
 * Compute target cell/range address given a source address and fill mode.
 * - 'below': shifts down by source rowCount (e.g., A1 -> A2, A1:C1 -> A2:C2, A1:A3 -> A4:A6)
 * - 'right': shifts right by source colCount (e.g., A1 -> B1, A1:A5 -> B1:B5)
 * - 'replace': returns the same range
 * - 'custom': uses customStartCell (e.g., "C5" or "Sheet2!C5") sized to match source dimensions
 * @param {string} sourceAddress
 * @param {'below'|'right'|'replace'|'custom'} mode
 * @param {string} [customStartCell]
 * @returns {string}
 */
export function computeTargetAddress(sourceAddress, mode = 'below', customStartCell = '') {
  const src = parseRangeAddress(sourceAddress);

  if (mode === 'replace') {
    return formatRangeAddress(src);
  }

  if (mode === 'below') {
    return formatRangeAddress({
      sheetPrefix: src.sheetPrefix,
      startRow: src.startRow + src.rowCount,
      endRow: src.endRow + src.rowCount,
      startCol: src.startCol,
      endCol: src.endCol
    });
  }

  if (mode === 'right') {
    return formatRangeAddress({
      sheetPrefix: src.sheetPrefix,
      startRow: src.startRow,
      endRow: src.endRow,
      startCol: src.startCol + src.colCount,
      endCol: src.endCol + src.colCount
    });
  }

  if (mode === 'custom') {
    if (!customStartCell || !customStartCell.trim()) {
      throw new Error('Custom target address cannot be empty');
    }
    const custom = parseRangeAddress(customStartCell.trim());
    const prefix = custom.sheetPrefix || src.sheetPrefix;
    return formatRangeAddress({
      sheetPrefix: prefix,
      startRow: custom.startRow,
      startCol: custom.startCol,
      endRow: custom.startRow + src.rowCount - 1,
      endCol: custom.startCol + src.colCount - 1
    });
  }

  throw new Error(`Unsupported target mode: ${mode}`);
}

/**
 * Strip sheet prefix when passing an address into worksheet.getRange(...) if on the same sheet,
 * or split into { sheetName, rangePart } for cross-sheet references.
 */
export function splitSheetAndRange(address) {
  const trimmed = address.trim();
  const bangIndex = trimmed.lastIndexOf('!');
  if (bangIndex === -1) {
    return { sheetName: null, rangeAddress: trimmed };
  }
  let sheetName = trimmed.slice(0, bangIndex).trim();
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    sheetName = sheetName.slice(1, -1).replace(/''/g, "'");
  }
  return {
    sheetName,
    rangeAddress: trimmed.slice(bangIndex + 1).trim()
  };
}
