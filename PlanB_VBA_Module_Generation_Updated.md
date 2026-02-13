# PLAN B: CLAUDE CODE (CLI) — VBA Module Generation

> **PURPOSE**: Use this document with Claude Code CLI to generate all 12 VBA modules + 3 UserForms for the Excel VBA Billing System. Each module is output as a `.bas` file for import into Excel.

> **PREREQUISITE**: The workbook from Plan A (Claude in Excel) must be built first. These modules depend on the sheet names, table names, and named ranges created there.

> **USAGE**: Either paste this as your CLAUDE.md in the project root, or feed sections as prompts to Claude Code.

---

## PROJECT CONTEXT (CLAUDE.md)

```
# Excel VBA Billing System — Project Context

## Overview
Excel VBA billing system generating invoices, receipts, and ETR-style receipts 
compliant with Kenya, USA, and UK jurisdictions.

## Workbook Structure (Built by Plan A)
- Dashboard: Navigation hub with KPI cards and action buttons
- Invoice_Template: Print-ready invoice layout (A4)
- Receipt_Template: Payment receipt layout (A4)
- ETR_Template: 80mm thermal receipt (Kenya KRA)
- Customers: Customer database table (tblCustomers)
- Products: Product/service catalog (tblProducts)
- Transactions: Invoice ledger (tblTransactions)
- Settings: Configuration, tax rates, counters
- PaymentLog: Payment history (tblPaymentLog)
- TaxSummary: Automated tax reports (tblTaxSummary)

## Key Named Ranges (MUST use these exact names)
rngCompanyName, rngJurisdiction, rngCurrency, rngTaxTable,
rngLastInvoice, rngLastReceipt, rngLastETR, rngYearPrefix,
rngPaymentMethods, rngPaymentTerms,
rngCustomers, rngCustomerNames, rngCustomerIDs,
rngProducts, rngProductNames, rngProductSKUs,
rngTransactions, rngPaymentLog,
rngInvNumber, rngInvDate, rngInvDueDate, rngInvCustomer,
rngInvLineItems, rngInvSubtotal, rngInvTax, rngInvTotal,
rngRcptNumber, rngRcptDate,
rngETRNumber, rngETRDate, rngTaxSummary

## Table Names
tblCustomers, tblProducts, tblTransactions, tblPaymentLog, tblTaxSummary

## Dashboard Layout
- KPI Labels: Row 5 (A5, C5, E5, G5, I5) — DO NOT OVERWRITE
- KPI Formula Cells: Row 7 (A7, C7, E7, G7, I7)
- Recent Activity Header: Row 20
- Recent Activity Column Headers: Row 21
- Recent Activity Data: Rows 22-29 (8 data rows)

## Naming Conventions
- Modules: modXxx (e.g., modInvoice, modTax)
- UserForms: frmXxx (e.g., frmCustomerSelect)
- Constants: ALL_CAPS with underscores
- Variables: camelCase with type prefix (str, lng, dbl, ws, rng)
- Public variables: g_prefix (e.g., g_selectedCustomerID)

## Code Standards (ENFORCE THESE IN EVERY MODULE)
- Option Explicit in EVERY module
- Error handling (On Error GoTo) in EVERY Public Sub/Function
- Application.ScreenUpdating = False at start of operations
- Application.Calculation = xlCalculationManual for batch operations
- ALL sheet references explicit: ThisWorkbook.Sheets("SheetName")
- NEVER use ActiveSheet or Selection in production code
- Comment every function: purpose, parameters, return value
- Unprotect before writing to protected sheets, reprotect in Finally block
- Password for protected sheets: "admin2026"
- Round all currency to 2 decimals: WorksheetFunction.Round(val, 2)

## Tax Rules
- Kenya: VAT 16% (standard), 0% (exempt), 8% (petroleum). KRA PIN format: P/A + 9 digits + letter
- USA: State rates (CA 7.25%, TX 6.25%, NY 8%). EIN format: ##-#######
- UK: VAT 20% (standard), 5% (reduced), 0% (zero-rated). VAT No: GB + 9 digits

## Output
Generate each module as a separate .bas file. Ready for VBA Editor import.
```

---

## GENERATION ORDER

Generate modules in this exact sequence. Each builds on previous ones.

| Order | File | Dependencies | Test Criteria |
|-------|------|-------------|---------------|
| 1 | modUtilities.bas | None | FormatCurrency(1234.56) returns "KES 1,234.56" |
| 2 | modNumbering.bas | modUtilities | GetNextInvoiceNumber() 3x → INV-2026-0001, 0002, 0003 |
| 3 | modCustomer.bas | modUtilities | LookupCustomer("C001") returns full Dictionary |
| 4 | modProduct.bas | modUtilities | LookupProduct("SKU001") returns price 150.00 |
| 5 | modTax.bas | modUtilities | GetTaxRate("Standard Rate","Kenya") returns 0.16 |
| 6 | modInvoice.bas | All above | Full GenerateInvoice() end-to-end |
| 7 | modPayment.bas | modUtilities, modNumbering | RecordPayment updates balance correctly |
| 8 | modReceipt.bas | modPayment, modNumbering | GenerateReceipt populates template |
| 9 | modETR.bas | modInvoice, modTax | ETR generates with KRA fields (Kenya only) |
| 10 | modExport.bas | All templates | ExportToPDF creates file in correct folder |
| 11 | modDashboard.bas | All modules | RefreshDashboard() updates KPIs |
| 12 | modSecurity.bas | All modules | SetupWorkbook() protects all sheets |
| 13 | 3 UserForms | modCustomer, modProduct, modPayment | Search, select, submit actions work |

---

## MODULE 1: modUtilities.bas

**Prompt for Claude Code:**

```
Generate modUtilities.bas with these 10 functions. Include Option Explicit, 
full error handling, and detailed inline comments.

1. FormatCurrency(amount As Double) As String
   - Read rngCurrency from Settings sheet
   - Format: "KES 1,234.56" or "$1,234.56" or "£1,234.56"
   - Handle KES, USD ($), GBP (£)
   - Always 2 decimal places

2. FormatDate(dt As Date, Optional style As String = "standard") As String
   - "standard" → dd-mmm-yyyy
   - "etr" → dd/mm/yyyy HH:MM
   - "file" → yyyy-mm-dd

3. GetSetting(settingName As String) As String
   - Search Settings sheet column A for label text
   - Return value from column B of matching row
   - Return "" if not found

4. ValidateInput(value As Variant, dataType As String, Optional minVal, Optional maxVal) As Boolean
   - "number" → is numeric, positive, within min/max range
   - "text" → non-empty string
   - "date" → is valid date
   - "email" → contains "@" and "."

5. ErrorHandler(procName As String, errNum As Long, errDesc As String)
   - Log to AuditLog sheet (create if not exists): timestamp, procedure, error#, description
   - Show user-friendly MsgBox: "An error occurred in [procName]. Error: [errDesc]. Please try again."
   - DO NOT use Resume — let calling code handle flow

6. TogglePerformance(turnOn As Boolean)
   - turnOn=True: ScreenUpdating=False, Calculation=xlCalculationManual, EnableEvents=False
   - turnOn=False: restore all to True/xlCalculationAutomatic/True
   - Wrap in error handler to ensure restoration

7. GetNextRow(ws As Worksheet, col As Long) As Long
   - Find next empty row in specified column
   - Handle empty sheet (return 2 for first data row below header)

8. SafeSheetRef(sheetName As String) As Worksheet
   - Return ThisWorkbook.Sheets(sheetName) with error handling
   - Return Nothing if sheet doesn't exist

9. AuditLog(action As String, details As String)
   - Write to AuditLog sheet: Now(), action, details, Application.UserName
   - Create AuditLog sheet if it doesn't exist (columns: Timestamp, Action, Details, User)
   - Handle protection: unprotect → write → reprotect

10. CleanupTemp()
    - Clear any named temp ranges used during generation
    - Reset status bar: Application.StatusBar = False

Output as modUtilities.bas ready for VBA import.
```

---

## MODULE 2: modNumbering.bas

**Prompt for Claude Code:**

```
Generate modNumbering.bas for auto-numbering invoices, receipts, and ETR receipts.

Functions required:

1. GetNextInvoiceNumber() As String
   - Read rngLastInvoice (Settings!B26), increment by 1
   - Read rngYearPrefix (Settings!B29)
   - Handle year rollover: if Year(Date) <> rngYearPrefix, reset counter to 1, update year
   - Format: "INV-2026-0001"
   - Write updated counter back to Settings (unprotect/reprotect)

2. GetNextReceiptNumber() As String
   - Same pattern using rngLastReceipt (Settings!B27)
   - Format: "RCPT-2026-0001"

3. GetNextETRNumber() As String
   - Same using rngLastETR (Settings!B28)
   - Format: "ETR-2026-0001"

4. FormatDocNumber(prefix As String, year As Long, counter As Long) As String
   - Shared formatter: prefix & "-" & year & "-" & Format(counter, "0000")

5. ValidateNumberSequence(docType As String) As Boolean
   - Check Transactions for gaps in invoice numbering
   - Log warning if gaps found
   - Return True if sequence is clean

6. Private IncrementCounter(counterRange As Range) As Long
   - Unprotect Settings with "admin2026"
   - Increment value
   - Reprotect in Finally block (even on error)
   - Return new value

CRITICAL: Settings sheet protection must be handled properly. Always 
reprotect even if increment fails. Use On Error GoTo ErrHandler pattern.
```

---

## MODULE 3: modCustomer.bas

**Prompt for Claude Code:**

```
Generate modCustomer.bas for customer management.

Functions:

1. LookupCustomer(identifier As String) As Object
   - Search tblCustomers by Cust_ID (column A) OR Company Name (column B)
   - Return Scripting.Dictionary with keys: ID, Name, Contact, Email, Phone, 
     Address, City, Country, TaxID, Terms, Status, Notes
   - Return Nothing if not found

2. PopulateInvoiceCustomer(custID As String)
   - Look up customer via LookupCustomer
   - Write to Invoice_Template: rngInvCustomer = Company Name
   - Write Address, Tax ID to adjacent cells (E10, E11)
   - Handle not-found with MsgBox

3. GetCustomerBalance(custID As String) As Double
   - SUMIFS on tblTransactions[Balance] where Cust_ID matches 
     AND Status <> "Paid" AND Status <> "Cancelled"

4. ListActiveCustomers() As Collection
   - Loop tblCustomers, return collection of names where Status = "Active"

5. ValidateCustomerTaxID(taxID As String, jurisdiction As String) As Boolean
   - Kenya PIN: starts with P or A, then 9 digits, then 1 letter
   - USA EIN: 2 digits, hyphen, 7 digits
   - UK VAT: "GB" then 9 digits

6. ShowCustomerSelector()
   - Display frmCustomerSelect UserForm
   - Set public variable g_selectedCustomerID
   - If cancelled, g_selectedCustomerID = ""

Use early binding: Dim dict As Scripting.Dictionary (reference Microsoft Scripting Runtime).
Include at module top: ' Requires reference: Microsoft Scripting Runtime
```

---

## MODULE 4: modProduct.bas

**Prompt for Claude Code:**

```
Generate modProduct.bas for product/service catalog management.

Functions:

1. LookupProduct(sku As String) As Object
   - Search tblProducts by SKU (column A)
   - Return Dictionary: SKU, Name, Description, Category, UnitPrice, Unit, TaxCategory, Status

2. AddLineItem(sku As String, qty As Double, Optional discountPct As Double = 0)
   - Find next empty row in rngInvLineItems (Invoice_Template rows 15-29)
   - Populate: row#, SKU, name+description, qty, unit price, discount%, tax rate, line total
   - Tax rate: call modTax.GetTaxRate(product.TaxCategory) 
   - Line total: qty * unitPrice * (1 - discount/100)
   - Max 15 items — show error if exceeded
   - Call RecalculateLineItems() after adding

3. RemoveLineItem(rowIndex As Long)
   - Clear row, shift remaining items up
   - Recalculate

4. RecalculateLineItems()
   - Loop filled rows in rngInvLineItems
   - Recalculate each line total
   - Update rngInvSubtotal = SUM of line totals
   - Call modTax.CalculateInvoiceTax()
   - Update rngInvTotal = subtotal - discount + tax

5. ListActiveProducts() As Collection
   - Return active product names from tblProducts

6. ShowProductSelector()
   - Display frmProductSelect UserForm
   - Form stays open for multiple adds until user clicks "Done"
```

---

## MODULE 5: modTax.bas

**Prompt for Claude Code:**

```
Generate modTax.bas — the multi-jurisdiction tax calculation engine.

Functions:

1. GetTaxRate(taxCategory As String, Optional jurisdiction As String = "") As Double
   - If jurisdiction empty, read rngJurisdiction
   - Search rngTaxTable for matching jurisdiction + tax name containing category
   - Return rate as decimal (0.16 for 16%)
   - Return 0 if exempt or not found

2. CalculateVAT(subtotal As Double, rate As Double) As Double
   - Return WorksheetFunction.Round(subtotal * rate, 2)

3. CalculateInvoiceTax() As Double
   - Loop rngInvLineItems on Invoice_Template
   - For each filled row: get tax category (from product), get rate, calculate tax on line amount
   - Handle multi-rate: UK invoice may have Standard(20%) + Reduced(5%) + Zero items
   - Sum all line-level taxes
   - Write total to rngInvTax
   - Return total tax amount

4. FormatTaxBreakdown() As String
   - Generate text: "VAT 16%: KES 1,200.00" or "VAT 20%: £500.00 | VAT 5%: £25.00"
   - Group by rate, show each rate's contribution

5. GetTaxLabel(jurisdiction As String) As String
   - Kenya → "VAT", USA → "Sales Tax", UK → "VAT"

6. ValidateTaxID(taxID As String) As Boolean
   - Delegate to modCustomer.ValidateCustomerTaxID with current rngJurisdiction

7. GenerateTaxSummary(fromDate As Date, toDate As Date, Optional jurisdiction As String = "All")
   - Aggregate from tblTransactions for date range
   - Group by month, jurisdiction, rate
   - Write to TaxSummary sheet

8. GetETRTaxFields() As Object
   - Return Dictionary: KRAPIN, VATAmount, ETRSerial
   - Only valid when jurisdiction = "Kenya"

All currency rounding: WorksheetFunction.Round(value, 2)
```

---

## MODULE 6: modInvoice.bas

**Prompt for Claude Code:**

```
Generate modInvoice.bas — the MAIN orchestrator module.

Functions:

1. GenerateInvoice()   ← PRIMARY ENTRY POINT (called from Dashboard button)
   Step 1:  ClearInvoiceTemplate()
   Step 2:  invoiceNo = modNumbering.GetNextInvoiceNumber()
            Write to rngInvNumber
   Step 3:  rngInvDate = Date
            Look up customer terms, calculate due date
   Step 4:  modCustomer.ShowCustomerSelector()
            If g_selectedCustomerID = "" Then Exit (cancelled)
   Step 5:  modCustomer.PopulateInvoiceCustomer(g_selectedCustomerID)
   Step 6:  modProduct.ShowProductSelector()  (user adds items, form loops)
   Step 7:  modTax.CalculateInvoiceTax()
   Step 8:  grandTotal = subtotal - discount + tax → write to rngInvTotal
   Step 9:  Activate Invoice_Template (preview)
   Step 10: MsgBox "Finalize this invoice?" → Yes: FinalizeInvoice(), No: keep editing

2. FinalizeInvoice()
   - Get next empty row in tblTransactions
   - Write: InvoiceNo, CustID, CustomerName, DateIssued, DueDate, Subtotal, 
     Tax, Discount, GrandTotal, AmountPaid=0, Balance=GrandTotal, 
     Status="Pending", Jurisdiction=rngJurisdiction, TaxRate
   - Call AuditLog("INVOICE_CREATED", invoiceNo)
   - Ask "Export to PDF?" → if yes, modExport.ExportToPDF("invoice")
   - MsgBox "Invoice " & invoiceNo & " created successfully!"

3. ClearInvoiceTemplate()
   - Unprotect Invoice_Template
   - Clear: B8:B11, E9:E11, A15:H29, H31:H35
   - Reprotect

4. EditInvoice(invoiceNo As String)
   - Find invoice in tblTransactions
   - Load back into template for editing

5. CancelInvoice(invoiceNo As String)
   - Set Status="Cancelled" in Transactions
   - Log cancellation

Use TogglePerformance at start/end. Rollback counter if finalize fails.
```

---

## MODULE 7: modPayment.bas

**Prompt for Claude Code:**

```
Generate modPayment.bas for payment processing.

Functions:

1. RecordPayment(invoiceNo As String, amount As Double, method As String, reference As String)
   - Validate: amount > 0, amount <= GetOutstandingBalance(invoiceNo)
   - Generate paymentID: "PAY-" & Year(Date) & "-" & Format(counter, "0000")
   - Write to tblPaymentLog: PaymentID, InvoiceNo, CustID, Date, Amount, Method, Reference
   - Update tblTransactions: AmountPaid += amount, Balance -= amount
   - Update Status: Balance=0→"Paid", Balance>0 and AmountPaid>0→"Partial", else "Pending"
   - AuditLog("PAYMENT_RECORDED", paymentID & " for " & invoiceNo)

2. AllocatePartialPayment(invoiceNo As String, amount As Double) As Double
   - Call RecordPayment with the amount
   - Return new remaining balance
   - Prevent overpayment (cap at balance)

3. GetOutstandingBalance(invoiceNo As String) As Double
   - From tblTransactions: GrandTotal - AmountPaid for matching invoice

4. GetPaymentHistory(invoiceNo As String) As Collection
   - All rows from tblPaymentLog matching invoice

5. GetCustomerOutstanding(custID As String) As Double
   - Sum balances from tblTransactions where CustID matches, Status<>"Paid","Cancelled"

6. ShowPaymentEntry(invoiceNo As String)
   - Display frmPaymentEntry
   - Pre-populate: invoice details, balance
   - Validate amount <= balance before recording

Thread-safe: unprotect Transactions & PaymentLog before write, reprotect after.
```

---

## MODULE 8: modReceipt.bas

**Prompt for Claude Code:**

```
Generate modReceipt.bas for receipt generation.

Functions:

1. GenerateReceipt(invoiceNo As String)
   - ClearReceiptTemplate()
   - Set receipt number from modNumbering.GetNextReceiptNumber()
   - Look up invoice in tblTransactions
   - Populate: receipt no, date, invoice ref, customer name, tax ID
   - Set Amount Due = GrandTotal - previous AmountPaid
   - ShowPaymentEntry() for user to enter payment
   - After payment: populate receipt with payment details
   - Ask to finalize → export option

2. GenerateReceiptFromPayment(paymentID As String)
   - Create receipt from existing PaymentLog entry (for reprinting)

3. ClearReceiptTemplate()
   - Unprotect → clear all input cells → reprotect

4. PreviewReceipt()
   - Activate Receipt_Template, trigger print preview
```

---

## MODULE 9: modETR.bas

**Prompt for Claude Code:**

```
Generate modETR.bas for Kenya KRA ETR thermal receipts.

Functions:

1. GenerateETR(invoiceNo As String)
   - VALIDATE: rngJurisdiction MUST = "Kenya". Error if not.
   - Clear ETR_Template
   - Set ETR number from modNumbering.GetNextETRNumber()
   - Populate: company details, KRA PIN, date/time (Now()), cashier
   - Copy line items from invoice to ETR compact format (name | qty | amount)
   - Calculate: subtotal, VAT 16%, total
   - Set payment method, amount tendered, change
   - Generate ETR serial number

2. FormatThermal()
   - Set all cells: Consolas 9pt
   - Set column widths for 80mm
   - Use "=" and "-" separator lines
   - Ensure precise alignment

3. PrintETR()
   - Print to default printer, narrow margins, no headers/footers
   - Log print event

4. GenerateETRSerial() As String
   - "ETR" & Format(counter, "000000") & "-" & Format(Date, "yyyymmdd")

5. AddKRAFields()
   - Populate KRA PIN, compliance footer, serial number

This module ONLY activates when jurisdiction = "Kenya". Show error otherwise.
```

---

## MODULE 10: modExport.bas

**Prompt for Claude Code:**

```
Generate modExport.bas for PDF export and file management.

Functions:

1. ExportToPDF(docType As String, Optional docNumber As String = "")
   - docType: "invoice", "receipt", "etr"
   - Map to sheet: Invoice_Template / Receipt_Template / ETR_Template
   - Build path: [Settings PDF path]\[docType]s\[Year]\[Month]\filename.pdf
   - Filename: INV-2026-0001_CustomerName_2026-02-13.pdf
   - CreateFolderStructure() first
   - ws.ExportAsFixedFormat xlTypePDF, filePath, xlQualityStandard
   - AuditLog("PDF_EXPORTED", filePath)
   - MsgBox "Open PDF?" → Shell to open

2. CreateFolderStructure(basePath As String, docType As String)
   - Create nested folders: basePath\Invoices\2026\02\
   - Use Dir() + MkDir, handle existing folders
   - Cross-platform: #If Mac Then use MacScript

3. GenerateFileName(docType As String, docNumber As String, customerName As String) As String
   - Sanitize name: remove /\:*?"<>| characters
   - Format: docNumber & "_" & name & "_" & Format(Date,"yyyy-mm-dd") & ".pdf"

4. BatchExport(docType As String, fromDate As Date, toDate As Date)
   - Find matching documents in Transactions
   - Loop: load each into template, export
   - Progress: Application.StatusBar = "Exporting " & i & " of " & total
   - Summary MsgBox at end

5. ExportToEmail(filePath As String, recipientEmail As String)
   - Create Outlook MailItem with PDF attachment (Windows only)
   - Display for review (don't auto-send)
   - #If Mac Then show instructions for manual attachment

Handle cross-platform paths: "\" for Windows, "/" for Mac.
```

---

## MODULE 11: modDashboard.bas

**Prompt for Claude Code:**

```
Generate modDashboard.bas for dashboard management and navigation.

CRITICAL LAYOUT REFERENCES:
- KPI Labels: Row 5 (A5, C5, E5, G5, I5) — DO NOT OVERWRITE these cells
- KPI Formula Cells: Row 7 (A7, C7, E7, G7, I7) — these contain the calculated values
- Recent Activity Header: Row 20
- Recent Activity Column Headers: Row 21
- Recent Activity Data: Rows 22-29 (8 data rows)

Functions:

1. RefreshDashboard()
   - Recalculate Dashboard KPI formulas (cells A7, C7, E7, G7, I7)
   - Update "Recent Activity" rows 22-29 (last 8 from tblTransactions)
   - Update timestamp
   - Call CheckOverdueInvoices()

2. NavigateTo(sheetName As String)
   - Activate sheet, select A1

3. Button Click Handlers (one Sub each):
   - btnNewInvoice_Click() → modInvoice.GenerateInvoice()
   - btnRecordPayment_Click() → show invoice selector, then modReceipt.GenerateReceipt()
   - btnGenerateReceipt_Click() → receipt from existing payment
   - btnETRReceipt_Click() → validate Kenya, call modETR.GenerateETR()
   - btnExportPDF_Click() → show export options
   - btnViewCustomers_Click() → NavigateTo("Customers")
   - btnViewProducts_Click() → NavigateTo("Products")
   - btnTransactions_Click() → NavigateTo("Transactions")
   - btnTaxSummary_Click() → NavigateTo("TaxSummary")
   - btnSettings_Click() → NavigateTo("Settings")

4. AssignMacrosToButtons()
   - Loop Dashboard shapes, assign macro names by shape.Name matching
   - Call during SetupWorkbook()

5. CheckOverdueInvoices() As Long
   - Scan tblTransactions: DueDate < Date AND Status = "Pending" or "Partial"
   - Update status to "Overdue"
   - Return count of newly overdue

6. UpdateRecentActivity()
   - Clear rows 22-29 on Dashboard
   - Get last 8 transactions from tblTransactions (sorted by date desc)
   - Populate: Invoice No, Customer, Date, Amount, Status
   - Handle fewer than 8 transactions gracefully

7. Workbook_Open() handler (NOTE: this goes in ThisWorkbook module, not modDashboard):
   - Call modDashboard.RefreshDashboard()
   - count = CheckOverdueInvoices()
   - NavigateTo "Dashboard"
   - If count > 0 Then MsgBox count & " invoices are now overdue!"
```

---

## MODULE 12: modSecurity.bas

**Prompt for Claude Code:**

```
Generate modSecurity.bas for workbook protection and security.

Functions:

1. ProtectAllSheets()
   Protection matrix:
   - Settings: password "admin2026", unlock only B2:B8,B11,B26:B29,B41:B43
   - Invoice/Receipt/ETR Templates: protect, unlock input areas, UseInterfaceOnly:=True
   - Customers/Products: protect structure, allow sort/filter/autofilter
   - Transactions/PaymentLog: protect with UseInterfaceOnly:=True
   - Dashboard: full protection, nothing unlocked
   - TaxSummary: full protection

2. UnprotectSheet(ws As Worksheet, Optional password As String = "admin2026")
   - On Error Resume Next for already-unprotected sheets

3. ReprotectSheet(ws As Worksheet, Optional password As String = "admin2026", Optional uiOnly As Boolean = False)
   - Apply correct protection settings per sheet type

4. LockFormulaCells()
   - Loop all sheets, find formula cells, set Locked = True
   - Ensure input cells remain Locked = False

5. ValidateUser(Optional requiredRole As String = "") As Boolean
   - InputBox password prompt for admin functions
   - Compare to "admin2026"

6. ProtectVBAProject()
   - Comment-only: instructions for manually locking VBA project
   - (Cannot be automated in standard VBA)

7. SetupWorkbook()   ← MASTER SETUP — RUN ONCE AFTER ALL MODULES IMPORTED
   - Call ProtectAllSheets()
   - Call LockFormulaCells()
   - Call modDashboard.AssignMacrosToButtons()
   - Call modDashboard.RefreshDashboard()
   - Create AuditLog sheet if not exists
   - MsgBox "Billing System setup complete! All sheets protected."
```

---

## 3 USERFORMS

**Prompt for Claude Code:**

```
Generate 3 UserForm VBA code files. Since .frm files require binary .frx,
generate them as code that creates forms programmatically, OR as standard 
.frm text format.

### frmCustomerSelect
- Title: "Select Customer"
- Controls: txtSearch (TextBox), lstCustomers (ListBox), btnSelect, btnCancel, lblBalance
- txtSearch_Change: filter lstCustomers to matching names
- lstCustomers_Click: show customer balance in lblBalance
- btnSelect_Click: set g_selectedCustomerID = selected customer ID, Unload Me
- btnCancel_Click: g_selectedCustomerID = "", Unload Me
- UserForm_Initialize: populate from modCustomer.ListActiveCustomers()

### frmProductSelect  
- Title: "Add Line Item"
- Controls: txtSearch, cmbCategory (category filter), lstProducts (Name|Price), 
  txtQty, txtDiscount (default 0), btnAdd, btnDone, lblLineTotal
- txtQty_Change / txtDiscount_Change: update lblLineTotal preview
- btnAdd_Click: call modProduct.AddLineItem(), clear qty/discount, keep form open
- btnDone_Click: Unload Me
- UserForm_Initialize: populate products, categories from modProduct

### frmPaymentEntry
- Title: "Record Payment"  
- Controls: lblInvoiceNo, lblCustomer, lblGrandTotal, lblAmountPaid, lblBalance,
  txtAmount, cmbMethod (from rngPaymentMethods), txtReference, btnRecord, btnCancel
- Validation: amount > 0 AND amount <= balance due
- cmbMethod_Change: if M-Pesa or Card, make txtReference mandatory
- btnRecord_Click: modPayment.RecordPayment(), MsgBox confirmation, Unload
- btnCancel_Click: Unload Me

All forms: centered on screen, modal, consistent font (Arial 10pt), 
consistent button styling.
```

---

## FINAL: IMPORT & SETUP INSTRUCTIONS

After Claude Code generates all files:

1. Open `BillingSystem_v1.xlsm` (from Plan A)
2. Press `Alt+F11` to open VBA Editor
3. Go to `Tools → References` and check:
   - Microsoft Scripting Runtime
   - Microsoft Outlook Object Library (for email export)
4. Import modules: `File → Import File` for each `.bas` file
5. Import UserForms: `File → Import File` for each `.frm` file
6. In the VBA Editor, find `ThisWorkbook` in the Project Explorer
7. Add the `Workbook_Open` event handler code from modDashboard
8. Run `modSecurity.SetupWorkbook` once to initialize everything
9. Save, close, and reopen to test `Workbook_Open`
10. Test: Create a test invoice, record a payment, generate receipt, export PDF

---

## TESTING CHECKLIST

After importing all modules, verify:

- [ ] New Invoice: Dashboard → New Invoice → Select customer → Add 3 items → Finalize → Check Transactions
- [ ] Auto-numbering: Create 3 invoices → verify INV-2026-0001, 0002, 0003
- [ ] Tax calculation: Kenya VAT 16% on standard items, 0% on exempt
- [ ] Switch jurisdiction to UK → verify 20% VAT label
- [ ] Record full payment → status changes to "Paid"
- [ ] Record partial payment → status "Partial", balance correct
- [ ] Generate receipt → receipt template populated correctly
- [ ] ETR receipt → only works when jurisdiction = "Kenya"
- [ ] Export PDF → file saved in correct folder with correct name
- [ ] Dashboard KPIs → reflect correct totals after transactions (check row 7: A7, C7, E7, G7, I7)
- [ ] Recent Activity → shows last 8 transactions in rows 22-29
- [ ] Overdue check → past-due invoices marked "Overdue"
- [ ] Sheet protection → cannot manually edit Transactions/PaymentLog
- [ ] Audit log → all actions recorded with timestamps
