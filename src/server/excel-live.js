import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { computeTargetAddress, splitSheetAndRange } from '../client/range-utils.js';

const execFileAsync = promisify(execFile);

function escapePsSingleQuote(str) {
  return String(str).replace(/'/g, "''");
}

async function runPowerShellJson(script) {
  if (process.platform !== 'win32') {
    throw new Error('Live desktop Excel bridge is only supported on Windows.');
  }

  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  const { stdout, stderr } = await execFileAsync(
    'powershell',
    ['-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
    { timeout: 15000, maxBuffer: 4 * 1024 * 1024 }
  );

  const trimmed = (stdout || '').trim();
  if (!trimmed) {
    throw new Error(stderr ? stderr.trim() : 'No response from Excel automation script.');
  }

  return JSON.parse(trimmed);
}

/**
 * Read active selection (or explicit address) from an already-open Excel window via GetActiveObject.
 */
export async function readLiveExcelRange(explicitAddress = '') {
  const safeAddr = escapePsSingleQuote(explicitAddress.trim());
  const psScript = `
$ErrorActionPreference = 'Stop'
try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
  @{ ok = $false; error = 'No running Microsoft Excel window detected (or COM access is blocked by policy).' } | ConvertTo-Json -Compress
  exit 0
}
try {
  $wb = $excel.ActiveWorkbook
  if ($null -eq $wb) {
    @{ ok = $false; error = 'No active workbook is open in Excel.' } | ConvertTo-Json -Compress
    exit 0
  }
  $ws = $excel.ActiveSheet
  $addrInput = '${safeAddr}'
  if ($addrInput -ne '') {
    $targetRange = $ws.Range($addrInput)
  } else {
    $targetRange = $excel.Selection
  }
  $rows = $targetRange.Rows.Count
  $cols = $targetRange.Columns.Count
  $grid = @()
  for ($r = 1; $r -le $rows; $r++) {
    $rowArr = @()
    for ($c = 1; $c -le $cols; $c++) {
      $val = $targetRange.Cells.Item($r, $c).Text
      $rowArr += [string]$val
    }
    $grid += ,$rowArr
  }
  @{
    ok = $true
    workbook = [string]$wb.Name
    sheet = [string]$ws.Name
    address = [string]$targetRange.Address($false, $false)
    values = $grid
  } | ConvertTo-Json -Compress -Depth 6
} catch {
  @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress
}
`;

  const result = await runPowerShellJson(psScript);
  if (!result.ok) {
    throw new Error(result.error || 'Failed to read from active Excel window');
  }
  return result;
}

/**
 * Write a 2D array of translated values directly into an already-open Excel window.
 */
export async function writeLiveExcelRange(targetAddress, values2D) {
  const { sheetName, rangeAddress } = splitSheetAndRange(targetAddress);
  const safeRange = escapePsSingleQuote(rangeAddress);
  const safeSheet = sheetName ? escapePsSingleQuote(sheetName) : '';
  const payloadBase64 = Buffer.from(JSON.stringify(values2D), 'utf8').toString('base64');

  const psScript = `
$ErrorActionPreference = 'Stop'
try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
} catch {
  @{ ok = $false; error = 'No running Microsoft Excel window detected (or COM access is blocked by policy).' } | ConvertTo-Json -Compress
  exit 0
}
try {
  $wb = $excel.ActiveWorkbook
  if ($null -eq $wb) {
    @{ ok = $false; error = 'No active workbook is open in Excel.' } | ConvertTo-Json -Compress
    exit 0
  }
  $sheetArg = '${safeSheet}'
  if ($sheetArg -ne '') {
    $ws = $wb.Worksheets.Item($sheetArg)
  } else {
    $ws = $excel.ActiveSheet
  }
  $rng = $ws.Range('${safeRange}')
  $jsonStr = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('${payloadBase64}'))
  $grid = ConvertFrom-Json -InputObject $jsonStr
  $rowCount = $grid.Count
  for ($r = 0; $r -lt $rowCount; $r++) {
    $row = $grid[$r]
    if ($row -is [System.Array]) {
      for ($c = 0; $c -lt $row.Count; $c++) {
        $rng.Cells.Item($r + 1, $c + 1).Value2 = [string]$row[$c]
      }
    } else {
      $rng.Cells.Item($r + 1, 1).Value2 = [string]$row
    }
  }
  @{ ok = $true; address = [string]$rng.Address($false, $false) } | ConvertTo-Json -Compress
} catch {
  @{ ok = $false; error = $_.Exception.Message } | ConvertTo-Json -Compress
}
`;

  const result = await runPowerShellJson(psScript);
  if (!result.ok) {
    throw new Error(result.error || 'Failed to write to active Excel window');
  }
  return result;
}

export { computeTargetAddress };
