# ==============================================================================
# PowerShell Script: CreateInitialWorkbook.ps1
# Purpose: Creates the initial Excel workbook and imports builder modules
# Usage: Right-click → Run with PowerShell
# ==============================================================================

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Excel Billing System - Initial Setup  " -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Get script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Working directory: $scriptPath" -ForegroundColor Yellow
Write-Host ""

# Check if Excel is installed
try {
    $excel = New-Object -ComObject Excel.Application
    Write-Host "[OK] Excel detected: Version $($excel.Version)" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Excel not found. Please install Microsoft Excel." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

# Define file paths
$workbookPath = Join-Path $scriptPath "BillingSystem_Builder.xlsm"
$builderModule = Join-Path $scriptPath "modWorkbookBuilder.bas"
$importerModule = Join-Path $scriptPath "modModuleImporter.bas"

# Check if modules exist
$missingFiles = @()
if (-not (Test-Path $builderModule)) { $missingFiles += "modWorkbookBuilder.bas" }
if (-not (Test-Path $importerModule)) { $missingFiles += "modModuleImporter.bas" }

if ($missingFiles.Count -gt 0) {
    Write-Host "[ERROR] Missing files:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "  - $file" -ForegroundColor Red
    }
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "[OK] Required modules found" -ForegroundColor Green
Write-Host ""

# Check if workbook already exists
if (Test-Path $workbookPath) {
    Write-Host "[WARNING] BillingSystem_Builder.xlsm already exists!" -ForegroundColor Yellow
    $overwrite = Read-Host "Do you want to overwrite it? (yes/no)"
    if ($overwrite -ne "yes") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit
    }
    Remove-Item $workbookPath -Force
}

Write-Host "Creating new workbook..." -ForegroundColor Cyan

try {
    # Create Excel application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Create new workbook
    $workbook = $excel.Workbooks.Add()

    Write-Host "[OK] Workbook created" -ForegroundColor Green

    # Get VB Project
    $vbProj = $workbook.VBProject

    # Check VBA access
    try {
        $componentCount = $vbProj.VBComponents.Count
        Write-Host "[OK] VBA project access enabled" -ForegroundColor Green
    } catch {
        Write-Host "" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Yellow
        Write-Host " VBA ACCESS REQUIRED" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please enable VBA project access:" -ForegroundColor Yellow
        Write-Host "1. Open Excel Options" -ForegroundColor White
        Write-Host "2. Go to Trust Center → Trust Center Settings" -ForegroundColor White
        Write-Host "3. Go to Macro Settings" -ForegroundColor White
        Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor White
        Write-Host "5. Run this script again" -ForegroundColor White
        Write-Host ""

        # Save basic workbook anyway
        $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
        Write-Host "[OK] Basic workbook saved. Import modules manually after enabling VBA access." -ForegroundColor Yellow

        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Read-Host "Press Enter to exit"
        exit
    }

    Write-Host "Importing builder modules..." -ForegroundColor Cyan

    # Import modWorkbookBuilder
    Write-Host "  - Importing modWorkbookBuilder.bas..." -ForegroundColor White
    $vbProj.VBComponents.Import($builderModule) | Out-Null
    Write-Host "    [OK]" -ForegroundColor Green

    # Import modModuleImporter
    Write-Host "  - Importing modModuleImporter.bas..." -ForegroundColor White
    $vbProj.VBComponents.Import($importerModule) | Out-Null
    Write-Host "    [OK]" -ForegroundColor Green

    Write-Host ""
    Write-Host "[OK] All modules imported successfully" -ForegroundColor Green
    Write-Host ""

    # Save as macro-enabled workbook
    Write-Host "Saving workbook as .xlsm..." -ForegroundColor Cyan
    $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
    Write-Host "[OK] Saved: BillingSystem_Builder.xlsm" -ForegroundColor Green

    # Close workbook
    $workbook.Close($false)
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($vbProj) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host " SUCCESS!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Workbook created: $workbookPath" -ForegroundColor White
    Write-Host ""
    Write-Host "NEXT STEPS:" -ForegroundColor Yellow
    Write-Host "1. Open BillingSystem_Builder.xlsm in Excel" -ForegroundColor White
    Write-Host "2. Enable macros when prompted" -ForegroundColor White
    Write-Host "3. Press Alt+F8 and run: BuildCompleteWorkbook" -ForegroundColor White
    Write-Host "4. Press Alt+F8 and run: ImportAllModules" -ForegroundColor White
    Write-Host "5. Follow remaining steps in SETUP_INSTRUCTIONS.md" -ForegroundColor White
    Write-Host ""

    # Ask if user wants to open the workbook
    $openNow = Read-Host "Do you want to open the workbook now? (yes/no)"
    if ($openNow -eq "yes") {
        Start-Process $workbookPath
        Write-Host "Opening workbook..." -ForegroundColor Cyan
    }

} catch {
    Write-Host ""
    Write-Host "[ERROR] Failed to create workbook:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack trace:" -ForegroundColor Yellow
    Write-Host $_.Exception.StackTrace -ForegroundColor Yellow

    # Cleanup
    if ($workbook) { $workbook.Close($false) }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host ""
Read-Host "Press Enter to exit"
