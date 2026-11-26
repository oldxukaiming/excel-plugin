param(
    [switch]$InstallToAddIns
)

$ErrorActionPreference = "Stop"

$RootDir = Split-Path -Parent $PSScriptRoot
$BasPath = Join-Path $RootDir "src\vba\ExcelTranslator.bas"
$DistDir = Join-Path $RootDir "dist"
$OutXlam = Join-Path $DistDir "ExcelTranslator.xlam"

if (-not (Test-Path $BasPath)) {
    throw "VBA module not found: $BasPath"
}

New-Item -ItemType Directory -Force -Path $DistDir | Out-Null
if (Test-Path $OutXlam) {
    Remove-Item -Force $OutXlam
}

function Show-ManualImportInstructions {
    param([string]$Reason)
    Write-Host ""
    Write-Host "[WARN] Automatic .xlam build via Excel COM failed: $Reason" -ForegroundColor Yellow
    Write-Host "This commonly happens on restricted machines where COM automation or VBA object model access is blocked by policy/security software." -ForegroundColor Gray
    Write-Host ""
    Write-Host "Recommended Option A (Modern Office.js Taskpane, no shared folder & no VBA needed):" -ForegroundColor Cyan
    Write-Host "  Run: npm run addin:register"
    Write-Host "  Then start server (npm start) and open Excel -> Add-ins -> Developer Add-ins."
    Write-Host ""
    Write-Host "Option B (Manual VBA import in Excel, ~10 seconds):" -ForegroundColor Cyan
    Write-Host "  1. Open Excel and press Alt + F11 to open the VBA Editor."
    Write-Host "  2. Right-click VBAProject -> Import File... -> select:"
    Write-Host "     $BasPath"
    Write-Host "  3. Use File -> Save As -> Excel Add-in (*.xlam), or call macros directly:"
    Write-Host "     - TranslateSelectionBelow (A1 -> A2)"
    Write-Host "     - SwitchTranslateEngine   (Bing <-> Google)"
    Write-Host "     - =XLTRANS(A1, ""zh-CN"", ""auto"", ""bing"")"
}

$originalCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
$originalUiCulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture
$excel = $null
$workbook = $null

try {
    $enUs = [System.Globalization.CultureInfo]::GetCultureInfo("en-US")
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $enUs
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $enUs

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Add()
    } catch {
        Show-ManualImportInstructions -Reason $_.Exception.Message
        exit 1
    }

    try {
        $vbProj = $workbook.VBProject
        $null = $vbProj.VBComponents.Import($BasPath)
    } catch {
        Show-ManualImportInstructions -Reason "VBA project access blocked ($($_.Exception.Message))"
        exit 1
    }

    # 55 = xlOpenXMLAddIn (.xlam)
    $workbook.SaveAs($OutXlam, 55)
    Write-Host "[OK] Built Excel Add-in: $OutXlam" -ForegroundColor Green

    if ($InstallToAddIns) {
        $userAddIns = Join-Path $env:APPDATA "Microsoft\AddIns"
        New-Item -ItemType Directory -Force -Path $userAddIns | Out-Null
        $installedPath = Join-Path $userAddIns "ExcelTranslator.xlam"
        Copy-Item -Force $OutXlam $installedPath
        Write-Host "[OK] Copied to user AddIns folder: $installedPath" -ForegroundColor Green
    }
} catch {
    Show-ManualImportInstructions -Reason $_.Exception.Message
    exit 1
} finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $originalUiCulture

    if ($workbook -ne $null) {
        try { $workbook.Close($false) | Out-Null } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        try { $excel.Quit() | Out-Null } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
