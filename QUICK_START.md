# 🚀 QUICK START - Excel VBA Billing System

## ⚡ 5-Minute Setup

### Method 1: PowerShell Automation (Recommended - Windows Only)

1. **Enable PowerShell scripts** (if first time):
   - Right-click PowerShell as Administrator
   - Run: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

2. **Run the setup script**:
   - Right-click `CreateInitialWorkbook.ps1`
   - Select **"Run with PowerShell"**
   - Follow prompts
   - ✅ Done! Workbook created with builder modules imported

3. **Open the workbook**:
   - Open `BillingSystem_Builder.xlsm`
   - Enable macros when prompted

4. **Build the system**:
   - Press **Alt + F8**
   - Run: **`BuildCompleteWorkbook`** (takes 1-2 minutes)
   - Run: **`ImportAllModules`** (takes 30 seconds)

5. **Add event handlers**:
   - Press **Alt + F11** (VBA Editor)
   - Double-click **Dashboard** sheet → Paste:
     ```vba
     Private Sub Worksheet_SelectionChange(ByVal Target As Range)
         On Error Resume Next
         modDashboard.HandleDashboardClick Target
     End Sub
     ```
   - Double-click **ThisWorkbook** → Paste:
     ```vba
     Private Sub Workbook_Open()
         On Error Resume Next
         modDashboard.RefreshDashboard
         modDashboard.NavigateTo "Dashboard"
     End Sub
     ```

6. **Add references**:
   - Tools → References
   - Check: ✅ Microsoft Scripting Runtime

7. **Save and test**:
   - Press Ctrl + S
   - Close and reopen workbook
   - Click "NEW INVOICE" on Dashboard

✅ **DONE!** System ready to use.

---

### Method 2: Manual Setup (All Platforms)

1. **Create workbook**:
   - Open Excel
   - Save as `BillingSystem_Builder.xlsm` (Macro-Enabled)

2. **Import builder modules**:
   - Press **Alt + F11** (VBA Editor)
   - File → Import File → Select `modWorkbookBuilder.bas`
   - File → Import File → Select `modModuleImporter.bas`

3. **Enable VBA access**:
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Check "Trust access to the VBA project object model"
   - Click OK, restart Excel

4. **Follow steps 4-7 from Method 1** above

---

## 📊 What You'll Get

- **10 worksheets** fully configured
- **30+ named ranges** set up
- **14 VBA modules** imported (1500+ lines of code)
- **Sample data**: 10 customers, 20 products
- **Professional templates**: Invoice, Receipt, ETR
- **Dashboard** with KPIs and navigation
- **Multi-jurisdiction tax** support (Kenya, USA, UK)

---

## 🎯 First Invoice Test

After setup, create your first invoice:

1. Click **NEW INVOICE** on Dashboard
2. Select customer: **C001 - Safaricom PLC**
3. Click **Add Product**
4. Select: **SKU001 - IT Consulting**
5. Enter quantity: **10**
6. Click **Add**, then **Done**
7. Click **Finalize**
8. ✅ Invoice **INV-2026-0001** created!

---

## 🔧 Common Issues

| Issue | Solution |
|-------|----------|
| "Cannot run macro" | Enable macros: File → Options → Trust Center |
| "Compile error" | Add reference: Tools → References → Microsoft Scripting Runtime |
| "Permission denied" | Enable VBA access: Trust Center → Trust access to VBA project |
| Buttons don't work | Add event handler to Dashboard sheet (see Method 1, step 5) |

---

## 📁 Files Created

```
joseph-project1/
├── BillingSystem_Builder.xlsm    ← Your workbook (create this)
├── modWorkbookBuilder.bas         ← Builds structure
├── modModuleImporter.bas          ← Imports modules
├── [14 other .bas files]          ← System modules
├── SETUP_INSTRUCTIONS.md          ← Detailed guide
├── QUICK_START.md                 ← This file
└── CreateInitialWorkbook.ps1      ← PowerShell automation
```

---

## 📖 Full Documentation

- **SETUP_INSTRUCTIONS.md** - Detailed setup with troubleshooting
- **PlanA_Claude_in_Excel.md** - Workbook structure specs
- **PlanB_VBA_Module_Generation_Updated.md** - Module documentation

---

## 🎓 What's Next?

After setup:

1. **Customize Settings**
   - Go to Settings sheet
   - Enter your company details
   - Set tax IDs

2. **Add Real Data**
   - Update Customers sheet
   - Update Products sheet

3. **Test Features**
   - Create invoices
   - Record payments
   - Generate receipts
   - Export PDFs

4. **Go Live**
   - Train users
   - Create backups
   - Start billing!

---

## 💡 Pro Tips

- **Backup often**: File → Save As → New name
- **Export modules**: Run `ExportAllModules` for version control
- **Check AuditLog**: Tracks all system actions
- **Use Diagnostics**: Run `modDiagnostics` functions for debugging

---

**Need help?** Check SETUP_INSTRUCTIONS.md for detailed troubleshooting.

**Version:** 1.0 | **Date:** 2026-02-13
