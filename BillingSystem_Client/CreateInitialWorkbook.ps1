# ==============================================================================
# PowerShell Script: CreateInitialWorkbook.ps1
# Purpose: One-click setup for the Excel Billing System
# Usage: Right-click -> Run with PowerShell
# ==============================================================================

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Excel Billing System - Setup" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Get script directory
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Working directory: $scriptPath" -ForegroundColor Yellow
Write-Host ""

# ---- Auto-enable VBA Trust Access via Registry ----
Write-Host "Checking VBA trust access settings..." -ForegroundColor Cyan
$excelVersions = @("16.0", "15.0", "14.0")
$trustEnabled = $false
foreach ($ver in $excelVersions) {
    $regPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (Test-Path $regPath) {
        try {
            $currentVal = Get-ItemProperty -Path $regPath -Name "AccessVBOM" -ErrorAction SilentlyContinue
            if ($currentVal -and $currentVal.AccessVBOM -eq 1) {
                Write-Host "[OK] VBA trust access already enabled (Office $ver)" -ForegroundColor Green
                $trustEnabled = $true
            } else {
                Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -ErrorAction Stop
                Write-Host "[OK] VBA trust access enabled via registry (Office $ver)" -ForegroundColor Green
                $trustEnabled = $true
            }
        } catch {
            Write-Host "[WARN] Could not set registry for Office $ver" -ForegroundColor Yellow
        }
        break
    }
}
if (-not $trustEnabled) {
    Write-Host "[WARN] Could not auto-enable VBA trust. Will try anyway..." -ForegroundColor Yellow
    Write-Host "  If setup fails, manually enable:" -ForegroundColor Yellow
    Write-Host "  Excel -> File -> Options -> Trust Center -> Trust Center Settings" -ForegroundColor White
    Write-Host "  -> Macro Settings -> Check 'Trust access to the VBA project object model'" -ForegroundColor White
}
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

# Define all 16 modules
$allModules = @(
    "modWorkbookBuilder.bas",
    "modModuleImporter.bas",
    "modUtilities.bas",
    "modNumbering.bas",
    "modCustomer.bas",
    "modProduct.bas",
    "modTax.bas",
    "modInvoice.bas",
    "modPayment.bas",
    "modReceipt.bas",
    "modETR.bas",
    "modExport.bas",
    "modDashboard.bas",
    "modSecurity.bas",
    "modForms.bas",
    "modDiagnostics.bas"
)

# Check all modules exist
$missingFiles = @()
foreach ($mod in $allModules) {
    $modPath = Join-Path $scriptPath $mod
    if (-not (Test-Path $modPath)) { $missingFiles += $mod }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "[ERROR] Missing files:" -ForegroundColor Red
    foreach ($file in $missingFiles) {
        Write-Host "  - $file" -ForegroundColor Red
    }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "[OK] All 16 modules found" -ForegroundColor Green
Write-Host ""

# Check if workbook already exists
if (Test-Path $workbookPath) {
    Write-Host "[WARNING] BillingSystem_Builder.xlsm already exists!" -ForegroundColor Yellow
    $overwrite = Read-Host "Do you want to overwrite it? (yes/no)"
    if ($overwrite -ne "yes") {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Read-Host "Press Enter to exit"
        exit
    }
    Remove-Item $workbookPath -Force
}

Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "  Starting Setup..." -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Create Excel application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    Write-Host "[OK] Workbook created" -ForegroundColor Green

    # Check VBA access (access VBProject directly, don't cache)
    try {
        $testCount = $workbook.VBProject.VBComponents.Count
        Write-Host "[OK] VBA project access enabled (found $testCount components)" -ForegroundColor Green
    } catch {
        Write-Host "" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Yellow
        Write-Host " VBA ACCESS REQUIRED" -ForegroundColor Yellow
        Write-Host "=========================================" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Automatic registry setting did not work." -ForegroundColor Yellow
        Write-Host "Please enable VBA project access manually:" -ForegroundColor Yellow
        Write-Host "1. Open Excel" -ForegroundColor White
        Write-Host "2. File -> Options -> Trust Center -> Trust Center Settings" -ForegroundColor White
        Write-Host "3. Macro Settings" -ForegroundColor White
        Write-Host "4. Check 'Trust access to the VBA project object model'" -ForegroundColor White
        Write-Host "5. Click OK, close Excel, and run this script again" -ForegroundColor White
        Write-Host ""

        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Read-Host "Press Enter to exit"
        exit
    }

    Write-Host ""

    # STEP 1: Import all 16 modules (access VBProject directly each time - no caching)
    Write-Host "[Step 1/3] Importing all 16 VBA modules..." -ForegroundColor Cyan
    $importCount = 0
    foreach ($mod in $allModules) {
        $modPath = Join-Path $scriptPath $mod
        Write-Host "  - Importing $mod..." -ForegroundColor White -NoNewline
        try {
            $workbook.VBProject.VBComponents.Import($modPath) | Out-Null
            Write-Host " [OK]" -ForegroundColor Green
            $importCount++
        } catch {
            Write-Host " [FAILED]" -ForegroundColor Red
            Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Red
            throw "Failed to import module: $mod - $($_.Exception.Message)"
        }
    }
    Write-Host ""
    Write-Host "[OK] All $importCount modules imported" -ForegroundColor Green
    Write-Host ""

    # STEP 2: Save workbook (needed before running macros)
    Write-Host "[Step 2/3] Saving workbook..." -ForegroundColor Cyan
    $workbook.SaveAs($workbookPath, 52) # xlOpenXMLWorkbookMacroEnabled
    Write-Host "    [OK] Saved as .xlsm" -ForegroundColor Green
    Write-Host ""

    # STEP 3: Run BuildCompleteWorkbook macro
    Write-Host "[Step 3/3] Building complete workbook structure..." -ForegroundColor Cyan
    Write-Host "  Creating 10 sheets, tables, named ranges, and formatting." -ForegroundColor Gray
    Write-Host "  Please wait, this takes 1-2 minutes..." -ForegroundColor Gray
    Write-Host ""
    $excel.Run("BuildCompleteWorkbook")
    Write-Host "    [OK] Workbook structure built successfully" -ForegroundColor Green
    Write-Host ""

    # Final save
    Write-Host "Saving final workbook..." -ForegroundColor Cyan
    $workbook.Save()
    Write-Host "[OK] Workbook saved" -ForegroundColor Green

    # Close workbook
    $workbook.Close($false)
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host ""
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host " SETUP COMPLETE!" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Workbook created: $workbookPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  - 16 VBA modules imported" -ForegroundColor White
    Write-Host "  - 10 worksheets built and formatted" -ForegroundColor White
    Write-Host "  - 30+ named ranges created" -ForegroundColor White
    Write-Host "  - Dashboard, templates, and sample data ready" -ForegroundColor White
    Write-Host ""
    Write-Host "TO START:" -ForegroundColor Yellow
    Write-Host "1. Open BillingSystem_Builder.xlsm" -ForegroundColor White
    Write-Host "2. Enable macros when prompted" -ForegroundColor White
    Write-Host "3. Use the Dashboard to create invoices!" -ForegroundColor White
    Write-Host ""

    # Ask if user wants to open the workbook
    $openNow = Read-Host "Open the workbook now? (yes/no)"
    if ($openNow -eq "yes") {
        Start-Process $workbookPath
        Write-Host "Opening workbook..." -ForegroundColor Cyan
    }

} catch {
    Write-Host ""
    Write-Host "[ERROR] Setup failed:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Make sure Excel is closed before running this script" -ForegroundColor White
    Write-Host "2. Enable VBA trust: Excel -> File -> Options -> Trust Center" -ForegroundColor White
    Write-Host "   -> Trust Center Settings -> Macro Settings" -ForegroundColor White
    Write-Host "   -> Check 'Trust access to the VBA project object model'" -ForegroundColor White
    Write-Host "3. Run this script as Administrator (right-click -> Run as Administrator)" -ForegroundColor White
    Write-Host ""

    # Cleanup
    try {
        if ($workbook) { $workbook.Close($false) }
    } catch {}
    try {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    } catch {}
}

Write-Host ""
Read-Host "Press Enter to exit"
