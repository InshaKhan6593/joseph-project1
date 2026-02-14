Attribute VB_Name = "modWorkbookBuilder"
Option Explicit

' ==============================================================================
' Module: modWorkbookBuilder.bas
' Purpose: Programmatically builds the complete billing system workbook
'          Implements Plan A structure entirely in VBA for reproducibility
' Usage: Run BuildCompleteWorkbook() to create entire system
' ==============================================================================

' --------------------------------------------------------------------------
' MAIN ENTRY POINT: BuildCompleteWorkbook()
' Run this procedure to build the entire workbook from scratch
' --------------------------------------------------------------------------
Public Sub BuildCompleteWorkbook()
    Dim startTime As Double
    Dim stepName As String
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    MsgBox "Building Billing System Workbook..." & vbCrLf & _
           "This will take 1-2 minutes. Click OK to start.", vbInformation

    On Error GoTo StepError

    stepName = "CreateAllSheets"
    CreateAllSheets

    stepName = "BuildSettingsSheet"
    BuildSettingsSheet

    stepName = "BuildCustomersSheet"
    BuildCustomersSheet

    stepName = "BuildProductsSheet"
    BuildProductsSheet

    stepName = "BuildTransactionsSheet"
    BuildTransactionsSheet

    stepName = "BuildPaymentLogSheet"
    BuildPaymentLogSheet

    stepName = "BuildInvoiceTemplate"
    BuildInvoiceTemplate

    stepName = "BuildReceiptTemplate"
    BuildReceiptTemplate

    stepName = "BuildETRTemplate"
    BuildETRTemplate

    stepName = "BuildTaxSummarySheet"
    BuildTaxSummarySheet

    stepName = "BuildDashboardSheet"
    BuildDashboardSheet

    stepName = "CreateAllNamedRanges"
    CreateAllNamedRanges

    stepName = "ProtectAllSheets"
    ProtectAllSheets

    stepName = "InjectSheetCode"
    InjectSheetCode

    stepName = "FinalSetup"
    FinalSetup

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    Dim elapsed As Double
    elapsed = Timer - startTime

    MsgBox "Workbook built successfully!" & vbCrLf & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & _
           "Now save as .xlsm and run ImportAllModules.", vbInformation, "Build Complete"

    ThisWorkbook.Sheets("Dashboard").Activate
    Exit Sub

StepError:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Error in step [" & stepName & "]:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Build Failed"
End Sub

' --------------------------------------------------------------------------
' Step 1: Create All Sheets
' --------------------------------------------------------------------------
Private Sub CreateAllSheets()
    Dim sheetNames As Variant
    Dim tabColors As Variant
    Dim i As Long
    Dim ws As Worksheet

    ' Define sheets in order
    sheetNames = Array("Dashboard", "Invoice_Template", "Receipt_Template", "ETR_Template", _
                      "Customers", "Products", "Transactions", "Settings", "PaymentLog", "TaxSummary")

    ' RGB color codes
    tabColors = Array(RGB(27, 79, 114), RGB(230, 126, 34), RGB(230, 126, 34), RGB(230, 126, 34), _
                     RGB(39, 174, 96), RGB(39, 174, 96), RGB(17, 122, 101), RGB(192, 57, 43), _
                     RGB(17, 122, 101), RGB(125, 60, 152))

    ' Clear existing sheets except the first
    Application.DisplayAlerts = False
    Do While ThisWorkbook.Sheets.Count > 1
        ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True

    ' Rename first sheet
    ThisWorkbook.Sheets(1).Name = "Dashboard"
    ThisWorkbook.Sheets(1).Tab.Color = tabColors(0)

    ' Create remaining sheets
    For i = 1 To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetNames(i)
        ws.Tab.Color = tabColors(i)
    Next i
End Sub

' --------------------------------------------------------------------------
' Step 2: Build Settings Sheet
' --------------------------------------------------------------------------
Private Sub BuildSettingsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Settings")

    With ws
        ' Section A: Company Info
        .Range("A1:D1").Merge
        .Range("A1").Value = "COMPANY INFORMATION"
        FormatHeader .Range("A1")

        .Range("A2").Value = "Company Name"
        .Range("A3").Value = "Address Line 1"
        .Range("A4").Value = "Address Line 2"
        .Range("A5").Value = "Phone"
        .Range("A6").Value = "Email"
        .Range("A7").Value = "Website"
        .Range("A8").Value = "Logo Path"

        ' Section B: Jurisdiction
        .Range("A10:D10").Merge
        .Range("A10").Value = "ACTIVE JURISDICTION"
        FormatHeader .Range("A10")

        .Range("A11").Value = "Jurisdiction"
        .Range("B11").Value = "Kenya"
        On Error Resume Next
        .Range("B11").Validation.Delete
        .Range("B11").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Kenya,USA,UK"
        On Error GoTo 0

        .Range("A12").Value = "Currency Symbol"
        .Range("B12").NumberFormat = "@"
        .Range("B12").Value = "KES"

        ' Section C: Tax Rates
        .Range("A14:E14").Merge
        .Range("A14").Value = "TAX CONFIGURATION"
        FormatHeader .Range("A14")

        ' Tax table headers
        .Range("A15").Value = "Jurisdiction"
        .Range("B15").Value = "Tax Name"
        .Range("C15").Value = "Rate"
        .Range("D15").Value = "Tax ID Label"
        .Range("E15").Value = "Tax ID Value"
        .Range("A15:E15").Font.Bold = True

        ' Tax data
        Dim taxData As Variant
        taxData = Array( _
            Array("Kenya", "VAT Standard", "16%", "KRA PIN", ""), _
            Array("Kenya", "VAT Exempt", "0%", "KRA PIN", ""), _
            Array("Kenya", "Petroleum VAT", "8%", "KRA PIN", ""), _
            Array("USA", "Sales Tax (CA)", "7.25%", "EIN", ""), _
            Array("USA", "Sales Tax (TX)", "6.25%", "EIN", ""), _
            Array("USA", "Sales Tax (NY)", "8%", "EIN", ""), _
            Array("UK", "VAT Standard", "20%", "VAT Number", ""), _
            Array("UK", "VAT Reduced", "5%", "VAT Number", ""), _
            Array("UK", "VAT Zero", "0%", "VAT Number", "") _
        )

        Dim row As Long
        For row = 0 To UBound(taxData)
            .Cells(16 + row, 1).Value = taxData(row)(0)
            .Cells(16 + row, 2).Value = taxData(row)(1)
            .Cells(16 + row, 3).Value = taxData(row)(2)
            .Cells(16 + row, 4).Value = taxData(row)(3)
            .Cells(16 + row, 5).Value = taxData(row)(4)
        Next row
        
        ' Format Tax Rates as Percentage
        .Range("C16:C24").NumberFormat = "0.00%"

        ' Section D: Auto-Numbering Counters
        .Range("A25:D25").Merge
        .Range("A25").Value = "DOCUMENT COUNTERS"
        FormatHeader .Range("A25")

        .Range("A26").Value = "Last Invoice Number"
        .Range("B26").Value = 0
        .Range("A27").Value = "Last Receipt Number"
        .Range("B27").Value = 0
        .Range("A28").Value = "Last ETR Number"
        .Range("B28").Value = 0
        .Range("A29").Value = "Year Prefix"
        .Range("B29").Value = Year(Date)

        ' Section E: Payment Methods
        .Range("A31:D31").Merge
        .Range("A31").Value = "PAYMENT METHODS"
        FormatHeader .Range("A31")

        Dim paymentMethods As Variant
        paymentMethods = Array("Cash", "M-Pesa", "Bank Transfer", "Credit Card", "Debit Card", "Cheque", "Other")
        Dim i As Long
        For i = 0 To UBound(paymentMethods)
            .Cells(32 + i, 1).Value = paymentMethods(i)
        Next i

        ' Section F: Default Terms
        .Range("A45:D45").Merge
        .Range("A45").Value = "DEFAULT TERMS"
        FormatHeader .Range("A45")

        .Range("A46").Value = "Payment Terms"
        .Range("B46").Value = "Net 30"
        On Error Resume Next
        .Range("B46").Validation.Delete
        .Range("B46").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Due on Receipt,Net 15,Net 30,Net 60,Net 90"
        On Error GoTo 0

        .Range("A47").Value = "Default Discount %"
        .Range("B47").Value = 0
        .Range("B47").NumberFormat = "0.00%"

        .Range("A48").Value = "PDF Save Path"
        .Range("B48").Value = "C:\BillingSystem\"

        .Range("A49").Value = "Last Updated"
        .Range("B49").Formula = "=TODAY()"
        .Range("B49").NumberFormat = "dd-mmm-yyyy"

        ' Column widths
        .Columns("A:A").ColumnWidth = 20
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:E").ColumnWidth = 15
    End With
End Sub

' --------------------------------------------------------------------------
' Step 3: Build Customers Sheet
' --------------------------------------------------------------------------
Private Sub BuildCustomersSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Customers")

    With ws
        ' Headers
        Dim headers As Variant
        headers = Array("Cust_ID", "Company Name", "Contact Person", "Email", "Phone", _
                       "Billing Address", "City", "Country", "Tax ID (PIN/EIN/VAT)", "Credit Terms", "Status", "Notes")

        Dim col As Long
        For col = 1 To UBound(headers) + 1
            .Cells(1, col).Value = headers(col - 1)
        Next col

        ' Header formatting handled by TableStyle

        ' Sample data
        Dim custData As Variant
        custData = Array( _
            Array("C001", "Safaricom PLC", "James Mwangi", "james@safaricom.co.ke", "+254 722 000 001", "Waiyaki Way", "Nairobi", "Kenya", "P051234567X", "Net 30", "Active", ""), _
            Array("C002", "Kenya Airways", "Sarah Odhiambo", "sarah@kq.co.ke", "+254 733 000 002", "Airport North Rd", "Nairobi", "Kenya", "P051234568Y", "Net 60", "Active", ""), _
            Array("C003", "TechCorp Inc", "John Smith", "john@techcorp.com", "+1 415 555 0101", "100 Market St", "San Francisco", "USA", "12-3456789", "Net 30", "Active", ""), _
            Array("C004", "London Analytics Ltd", "Emma Thompson", "emma@londonanalytics.co.uk", "+44 20 7946 0958", "50 Baker Street", "London", "UK", "GB123456789", "Net 30", "Active", ""), _
            Array("C005", "Equity Bank", "Peter Ndung'u", "peter@equitybank.co.ke", "+254 711 000 005", "NHIF Building", "Nairobi", "Kenya", "P051234570A", "Due on Receipt", "Active", ""), _
            Array("C006", "Acme Corp", "Lisa Johnson", "lisa@acme.com", "+1 212 555 0202", "350 5th Avenue", "New York", "USA", "98-7654321", "Net 60", "Active", ""), _
            Array("C007", "British Gas", "Oliver Brown", "oliver@britishgas.co.uk", "+44 20 7946 1234", "1 Regent St", "London", "UK", "GB987654321", "Net 30", "Active", ""), _
            Array("C008", "Jumia Kenya", "Grace Wanjiku", "grace@jumia.co.ke", "+254 700 000 008", "Kilimani", "Nairobi", "Kenya", "P051234571B", "Net 30", "Active", ""), _
            Array("C009", "StarTech Solutions", "David Lee", "david@startech.com", "+1 650 555 0303", "200 University Ave", "Palo Alto", "USA", "55-1234567", "Net 30", "Active", ""), _
            Array("C010", "Manchester Digital", "Sophie Williams", "sophie@manchesterdigital.co.uk", "+44 161 555 0404", "1 St Peter's Sq", "Manchester", "UK", "GB555666777", "Net 60", "Active", "") _
        )

        Dim row As Long
        For row = 0 To UBound(custData)
            For col = 0 To UBound(custData(row))
                .Cells(2 + row, 1 + col).Value = custData(row)(col)
            Next col
        Next row

        ' Format IDs, Phones, Tax IDs, Credit Terms as Text
        .Columns("A:A").NumberFormat = "@" ' Cust_ID
        .Columns("E:E").NumberFormat = "@" ' Phone
        .Columns("I:J").NumberFormat = "@" ' Tax ID, Credit Terms

        ' Format as table
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1:L11"), , xlYes)
        tbl.Name = "tblCustomers"
        tbl.TableStyle = "TableStyleMedium2"

        ' Column widths
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 25
        .Columns("C:C").ColumnWidth = 20
        .Columns("D:D").ColumnWidth = 25
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 30
        .Columns("F:F").WrapText = True ' Address column wrap
        .Columns("G:G").ColumnWidth = 15
        .Columns("H:H").ColumnWidth = 15
        .Columns("I:I").ColumnWidth = 20
        .Columns("J:J").ColumnWidth = 15
        .Columns("K:K").ColumnWidth = 10
        .Columns("L:L").ColumnWidth = 25

        ' Freeze panes - activate sheet first
        On Error Resume Next
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 4: Build Products Sheet
' --------------------------------------------------------------------------
Private Sub BuildProductsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Products")

    With ws
        ' Headers
        Dim headers As Variant
        headers = Array("SKU", "Product/Service Name", "Description", "Category", "Unit Price", "Unit", "Tax Category", "Status")

        Dim col As Long
        For col = 1 To UBound(headers) + 1
            .Cells(1, col).Value = headers(col - 1)
        Next col

        ' Header formatting handled by TableStyle

        ' Sample data - 20 products
        Dim prodData As Variant
        prodData = Array( _
            Array("SKU001", "IT Consulting", "Hourly IT consulting services", "Consulting", 150, "Hour", "Standard Rate", "Active"), _
            Array("SKU002", "Web Development", "Full-stack web development", "Software", 200, "Hour", "Standard Rate", "Active"), _
            Array("SKU003", "Cloud Hosting (Monthly)", "AWS/Azure managed hosting", "Subscription", 500, "Month", "Standard Rate", "Active"), _
            Array("SKU004", "Laptop - Dell Latitude", "Business laptop 14"" i7", "Hardware", 1200, "Unit", "Standard Rate", "Active"), _
            Array("SKU005", "Software License - Annual", "Enterprise software license", "License", 999, "License", "Standard Rate", "Active"), _
            Array("SKU006", "Data Analysis Report", "Custom data analysis", "Consulting", 350, "Project", "Standard Rate", "Active"), _
            Array("SKU007", "Network Setup", "Office network installation", "Service", 800, "Project", "Standard Rate", "Active"), _
            Array("SKU008", "Cybersecurity Audit", "Security assessment", "Consulting", 2500, "Project", "Standard Rate", "Active"), _
            Array("SKU009", "Training Session", "On-site IT training", "Service", 175, "Hour", "Standard Rate", "Active"), _
            Array("SKU010", "Domain Registration", "Annual domain name", "Subscription", 15, "Month", "Zero Rated", "Active"), _
            Array("SKU011", "SSL Certificate", "Annual SSL cert", "Subscription", 75, "License", "Zero Rated", "Active"), _
            Array("SKU012", "Mobile App Development", "Cross-platform app dev", "Software", 250, "Hour", "Standard Rate", "Active"), _
            Array("SKU013", "Database Migration", "DB migration service", "Service", 3000, "Project", "Standard Rate", "Active"), _
            Array("SKU014", "Printer - HP LaserJet", "Office laser printer", "Hardware", 450, "Unit", "Standard Rate", "Active"), _
            Array("SKU015", "UPS Battery Backup", "APC 1500VA UPS", "Hardware", 280, "Unit", "Standard Rate", "Active"), _
            Array("SKU016", "SEO Services (Monthly)", "Search engine optimization", "Service", 600, "Month", "Standard Rate", "Active"), _
            Array("SKU017", "Logo Design", "Corporate logo design", "Service", 500, "Project", "Standard Rate", "Active"), _
            Array("SKU018", "Email Hosting", "Business email per user", "Subscription", 5, "Month", "Zero Rated", "Active"), _
            Array("SKU019", "Fiber Installation", "Office fiber optic setup", "Service", 1500, "Project", "Standard Rate", "Active"), _
            Array("SKU020", "Project Management", "PM services", "Consulting", 180, "Hour", "Standard Rate", "Active") _
        )

        Dim row As Long
        For row = 0 To UBound(prodData)
            For col = 0 To UBound(prodData(row))
                .Cells(2 + row, 1 + col).Value = prodData(row)(col)
            Next col
        Next row

        ' Format SKU, Category, Status as Text
        .Columns("A:A").NumberFormat = "@" ' SKU
        .Columns("D:D").NumberFormat = "@" ' Category
        .Columns("H:H").NumberFormat = "@" ' Status

        ' Format as table
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1:H21"), , xlYes)
        tbl.Name = "tblProducts"
        tbl.TableStyle = "TableStyleMedium2"

        ' Column widths
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 30
        .Columns("C:C").ColumnWidth = 35
        .Columns("D:D").ColumnWidth = 15
        .Columns("E:E").ColumnWidth = 15
        .Columns("F:F").ColumnWidth = 10
        .Columns("G:G").ColumnWidth = 18
        .Columns("H:H").ColumnWidth = 10

        ' Number format for price
        ' Number format for price
        ' Number format for price
        .Columns("E:E").NumberFormat = "[$KES] #,##0.00"
        .Columns("C:C").WrapText = True ' Description column wrap

        ' Freeze panes - activate sheet first
        On Error Resume Next
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 5: Build Transactions Sheet
' --------------------------------------------------------------------------
Private Sub BuildTransactionsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Transactions")

    With ws
        ' Headers
        Dim headers As Variant
        headers = Array("Invoice No", "Cust_ID", "Customer Name", "Date Issued", "Due Date", _
                       "Subtotal", "Tax Amount", "Discount", "Grand Total", "Amount Paid", _
                       "Balance", "Status", "Jurisdiction", "Tax Rate Applied", "Notes")

        Dim col As Long
        For col = 1 To UBound(headers) + 1
            .Cells(1, col).Value = headers(col - 1)
        Next col

        ' Format header
        ' Header formatting handled by TableStyle

        ' Format as table
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1:O1"), , xlYes)
        tbl.Name = "tblTransactions"
        tbl.TableStyle = "TableStyleMedium2"

        ' Column widths
        .Columns("A:A").ColumnWidth = 16
        .Columns("B:B").ColumnWidth = 10
        .Columns("C:C").ColumnWidth = 22
        .Columns("D:E").ColumnWidth = 12
        .Columns("F:K").ColumnWidth = 12
        .Columns("L:L").ColumnWidth = 10
        .Columns("M:N").ColumnWidth = 14
        .Columns("M:N").ColumnWidth = 14
        .Columns("O:O").ColumnWidth = 25
        .Columns("O:O").WrapText = True ' Notes column wrap

        ' Number formats
        .Columns("F:K").NumberFormat = "[$KES] #,##0.00"
        .Columns("N:N").NumberFormat = "0.00%" ' Tax Rate
        .Columns("D:E").NumberFormat = "dd-mmm-yyyy"
        .Columns("A:B").NumberFormat = "@" ' Invoice No, Cust_ID

        ' Freeze panes - activate sheet first
        On Error Resume Next
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 6: Build PaymentLog Sheet
' --------------------------------------------------------------------------
Private Sub BuildPaymentLogSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PaymentLog")

    With ws
        ' Headers
        Dim headers As Variant
        headers = Array("Payment ID", "Invoice No", "Cust_ID", "Payment Date", "Amount Paid", _
                       "Payment Method", "Reference No", "Received By", "Notes")

        Dim col As Long
        For col = 1 To UBound(headers) + 1
            .Cells(1, col).Value = headers(col - 1)
        Next col

        ' Format header
        ' Header formatting handled by TableStyle

        ' Format as table
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1:I1"), , xlYes)
        tbl.Name = "tblPaymentLog"
        tbl.TableStyle = "TableStyleMedium2"

        ' Column widths
        .Columns("A:A").ColumnWidth = 14
        .Columns("B:B").ColumnWidth = 16
        .Columns("C:C").ColumnWidth = 10
        .Columns("D:D").ColumnWidth = 14
        .Columns("E:E").ColumnWidth = 14
        .Columns("F:F").ColumnWidth = 16
        .Columns("G:G").ColumnWidth = 20
        .Columns("H:H").ColumnWidth = 15
        .Columns("I:I").ColumnWidth = 25

        ' Number formats
        .Columns("E:E").NumberFormat = "[$KES] #,##0.00"
        .Columns("D:D").NumberFormat = "dd-mmm-yyyy"
        .Columns("A:C").NumberFormat = "@" ' Payment ID, Invoice No, Cust_ID

        ' Freeze panes - activate sheet first
        On Error Resume Next
        ws.Activate
        ws.Range("A2").Select
        ActiveWindow.FreezePanes = False
        ActiveWindow.FreezePanes = True
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 7: Build Invoice Template
' --------------------------------------------------------------------------
Private Sub BuildInvoiceTemplate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Invoice_Template")

    With ws
        .Cells.Clear

        ' Hide gridlines (must activate sheet first)
        On Error Resume Next
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        On Error GoTo 0

        ' Company header placeholder (rows 1-4)
        ' Company header placeholder (rows 1-5 for more space)
        .Range("A1:C5").Merge
        .Range("A1").Value = "[LOGO]"
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
        .Range("A1").Interior.ColorIndex = xlNone ' Transparent background

        .Range("D1").Formula = "=IF(Settings!B2="""","""",Settings!B2)"
        .Range("D1").Font.Size = 18
        .Range("D1").Font.Bold = True

        .Range("D2").Formula = "=IF(AND(Settings!B3="""",Settings!B4=""""),"""",Settings!B3 & "" "" & Settings!B4)"
        .Range("D3").Formula = "=IF(AND(Settings!B5="""",Settings!B6=""""),"""",Settings!B5 & "" | "" & Settings!B6)"
        .Range("D4").Formula = "=IF(Settings!B7="""","""",Settings!B7)"

        ' Document title (row 6)
        .Range("A6:H6").Merge
        .Range("A6:H6").Merge
        .Range("A6").RowHeight = 40 ' Ensure height for Size 24 Font
        .Range("A6").Value = "INVOICE"
        With .Range("A6")
            .Font.Size = 24
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(27, 79, 114)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' Invoice info (left side, rows 8-11)
        .Range("A8").Value = "Invoice No:"
        .Range("A9").Value = "Date:"
        .Range("A10").Value = "Due Date:"
        .Range("A11").Value = "Terms:"
        .Range("A8:A11").Font.Bold = True
        
        ' Date formats
        .Range("B9:B10").NumberFormat = "dd-mmm-yyyy"

        ' Customer info (right side, rows 8-11)
        .Range("E8").Value = "BILL TO:"
        .Range("E8").Font.Bold = True
        .Range("E8").Interior.Color = RGB(27, 79, 114)
        .Range("E8").Font.Color = RGB(255, 255, 255)

        ' Line items header (row 14)
        Dim lineHeaders As Variant
        lineHeaders = Array("#", "SKU", "Description", "Qty", "Unit Price", "Discount%", "Tax", "Line Total")
        Dim col As Long
        For col = 1 To UBound(lineHeaders) + 1
            .Cells(14, col).Value = lineHeaders(col - 1)
        Next col

        With .Range("A14:H14")
            .Interior.Color = RGB(27, 79, 114)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' Line items data area (rows 15-29 = 15 rows)
        Dim row As Long
        For row = 15 To 29
            .Cells(row, 1).Value = row - 14  ' Row number
             ' Formula for line total in column H
            .Cells(row, 8).Formula = "=IF(D" & row & "="""","""",D" & row & "*E" & row & "*(1-F" & row & "/100))"
            .Cells(row, 8).NumberFormat = "[$KES] #,##0.00"
            .Cells(row, 6).NumberFormat = "0.00" ' Discount number (not %)
            .Cells(row, 5).NumberFormat = "[$KES] #,##0.00" ' Unit Price
        Next row

        ' Alternating row colors
        For row = 15 To 29 Step 2
            .Range("A" & row & ":H" & row).Interior.Color = RGB(235, 245, 251)
        Next row

        ' Add borders to main table
        With .Range("A14:H29")
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        .Range("A14:H14").Borders(xlEdgeBottom).Weight = xlMedium

        ' Totals section (rows 31-35)
        .Range("G31").Value = "Subtotal"
        .Range("G31").Value = "Subtotal"
        .Range("H31").Formula = "=IF(SUM(H15:H29)=0,"""",SUM(H15:H29))"
        .Range("H31").NumberFormat = "[$KES] #,##0.00"

        .Range("G32").Value = "Discount"
        .Range("H32").Value = 0
        .Range("H32").NumberFormat = "[$KES] #,##0.00"

        .Range("G33").Value = "Tax (VAT/Sales Tax)"
        .Range("H33").Value = 0
        .Range("H33").NumberFormat = "[$KES] #,##0.00"

        .Range("G35").Value = "GRAND TOTAL"
        .Range("G35").Value = "GRAND TOTAL"
        .Range("H35").Formula = "=IF(N(H31)-H32+H33=0,"""",N(H31)-H32+H33)"
        With .Range("G35:H35")
            .Font.Bold = True
            .Font.Size = 14
            .Interior.Color = RGB(230, 126, 34)
        End With
        .Range("H35").NumberFormat = "[$KES] #,##0.00"

        ' Add borders to totals
        With .Range("G31:H35")
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders(xlInsideHorizontal).Weight = xlHairline
        End With

        ' Footer (rows 37-40)
        .Range("A37").Value = "Payment Instructions:"
        .Range("A37").Font.Bold = True
        .Range("A38").Value = "(Bank details)"
        .Range("A39").Value = "Tax ID: (from Settings)"
        .Range("A40").Value = "Thank you for your business!"
        .Range("A40").HorizontalAlignment = xlCenter
        .Range("A40").Font.Italic = True
        .Range("A40").Font.Color = RGB(127, 127, 127)

        ' Column widths
        .Columns("A:A").ColumnWidth = 5
        .Columns("B:B").ColumnWidth = 12
        .Columns("C:C").ColumnWidth = 30
        .Columns("C:C").WrapText = True ' Description column wrap
        .Columns("D:D").ColumnWidth = 8
        .Columns("E:E").ColumnWidth = 12
        .Columns("F:F").ColumnWidth = 10
        .Columns("G:G").ColumnWidth = 20
        .Columns("H:H").ColumnWidth = 15

        ' Print setup
        On Error Resume Next
        With .PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .PrintArea = "A1:H40"
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 8: Build Receipt Template
' --------------------------------------------------------------------------
Private Sub BuildReceiptTemplate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Receipt_Template")

    With ws
        .Cells.Clear
        On Error Resume Next
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        On Error GoTo 0

        ' Company header (rows 1-4)
        .Range("A1:F1").Merge
        .Range("A1").Formula = "=IF(Settings!B2="""","""",Settings!B2)"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter

        .Range("A2:F2").Merge
        .Range("A2").Formula = "=IF(Settings!B3="""","""",Settings!B3)"
        .Range("A2").HorizontalAlignment = xlCenter

        ' Title (row 6)
        .Range("A6:F6").Merge
        .Range("A6:F6").Merge
        .Range("A6").RowHeight = 40 ' Ensure height for Size 24 Font
        .Range("A6").Value = "RECEIPT"
        With .Range("A6")
            .Font.Size = 24
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(27, 79, 114)
            .Font.Color = RGB(255, 255, 255)
        End With

        ' Receipt info (rows 8-12)
        .Range("A8").Value = "Receipt No:"
        .Range("A9").Value = "Date:"
        .Range("A10").Value = "Invoice Reference:"
        .Range("A11").Value = "Customer:"
        .Range("A12").Value = "Tax ID:"
        .Range("A8:A12").Font.Bold = True

        ' Payment details header (row 14)
        .Range("A14:F14").Merge
        .Range("A14").Value = "PAYMENT DETAILS"
        With .Range("A14")
            .Interior.Color = RGB(27, 79, 114)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' Payment details (rows 16-20)
        .Range("A16").Value = "Amount Due"
        .Range("A17").Value = "Amount Paid"
        .Range("A18").Value = "Payment Method"
        .Range("A19").Value = "Reference No"
        .Range("A20").Value = "Balance"
        .Range("A16:A20").Font.Bold = True

        .Range("B20").Formula = "=B16-B17"
        .Range("B20").Formula = "=B16-B17"
        .Range("B16:B20").NumberFormat = "[$KES] #,##0.00"

        ' Add borders to payment details
        With .Range("A16:B20")
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With

        ' Footer (rows 22-24)
        .Range("A22").Value = "Received by: _______________"
        .Range("D22").Value = "Date: _______________"
        .Range("A24:F24").Merge
        .Range("A24").Value = "This is a computer-generated receipt."
        .Range("A24").HorizontalAlignment = xlCenter
        .Range("A24").Font.Italic = True
        .Range("A24").Font.Size = 8

        ' Print setup
        On Error Resume Next
        With .PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .PrintArea = "A1:F24"
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
        On Error GoTo 0
    End With
End Sub

' --------------------------------------------------------------------------
' Step 9: Build ETR Template
' --------------------------------------------------------------------------
Private Sub BuildETRTemplate()
    Dim ws As Worksheet
    Dim etrStep As String
    Set ws = ThisWorkbook.Sheets("ETR_Template")
    
    On Error GoTo ETRError
    etrStep = "Clear cells"
    ws.Cells.Clear
    
    etrStep = "Activate sheet"
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo ETRError
    
    etrStep = "Column widths"
    ws.Columns("A:A").ColumnWidth = 15
    ws.Columns("B:B").ColumnWidth = 15
    ws.Columns("C:C").ColumnWidth = 15
    
    etrStep = "Format C as Currency"
    ws.Columns("C:C").NumberFormat = "[$KES] #,##0.00"
    
    etrStep = "Font setup A1:C44"
    ws.Range("A1:C44").Font.Name = "Consolas"
    ws.Range("A1:C44").Font.Size = 9
    
    etrStep = "Header row 1"
    ws.Range("A1:C1").Merge
    ws.Range("A1").Formula = "=IF(Settings!B2="""","""",Settings!B2)"
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 11
    
    etrStep = "Header row 2"
    ws.Range("A2:C2").Merge
    ws.Range("A2").Formula = "=IF(Settings!B3="""","""",Settings!B3)"
    ws.Range("A2").HorizontalAlignment = xlCenter
    
    etrStep = "Header row 3"
    ws.Range("A3:C3").Merge
    ws.Range("A3").Formula = "=IF(Settings!B5="""","""",Settings!B5)"
    ws.Range("A3").HorizontalAlignment = xlCenter
    
    etrStep = "Header row 4 KRA"
    ws.Range("A4:C4").Merge
    ws.Range("A4").Value = "KRA PIN: (from Settings)"
    ws.Range("A4").HorizontalAlignment = xlCenter
    
    etrStep = "Separator row 5"
    ws.Range("A5:C5").Merge
    ws.Range("A5").Value = "--------------------------------"
    
    etrStep = "ETR RECEIPT row 6"
    ws.Range("A6:C6").Merge
    ws.Range("A6").Value = "ETR RECEIPT"
    ws.Range("A6").HorizontalAlignment = xlCenter
    ws.Range("A6").Font.Bold = True
    
    etrStep = "Receipt details rows 7-9"
    ws.Range("A7").Value = "Receipt No:"
    ws.Range("A8").Value = "Date:"
    ws.Range("A9").Value = "Cashier:"
    
    etrStep = "Separator row 10"
    ws.Range("A10:C10").Merge
    ws.Range("A10").Value = "--------------------------------"
    
    etrStep = "Line items header row 11"
    ws.Range("A11:C11").Merge
    ws.Range("A11").Value = "ITEM          QTY    AMOUNT"
    ws.Range("A11").Font.Bold = True
    
    etrStep = "Separator row 12"
    ws.Range("A12:C12").Merge
    ws.Range("A12").Value = "--------------------------------"
    
    ' Line items rows 13-27 left blank
    
    etrStep = "Separator row 28"
    ws.Range("A28:C28").Merge
    ws.Range("A28").Value = "--------------------------------"
    
    etrStep = "Totals rows 29-30"
    ws.Range("A29").Value = "Subtotal:"
    ws.Range("A30").Value = "VAT (16%):"
    
    etrStep = "Separator row 31"
    ws.Range("A31:C31").Merge
    ws.Range("A31").Value = "--------------------------------"
    
    etrStep = "Total row 32"
    ws.Range("A32").Value = "TOTAL:"
    ws.Range("A32").Font.Bold = True
    
    etrStep = "Separator row 33"
    ws.Range("A33:C33").Merge
    ws.Range("A33").Value = "--------------------------------"
    
    etrStep = "Payment rows 34-36"
    ws.Range("A34").Value = "Payment:"
    ws.Range("A35").Value = "Amount Tendered:"
    ws.Range("A36").Value = "Change:"
    
    etrStep = "Separator row 37"
    ws.Range("A37:C37").Merge
    ws.Range("A37").Value = "--------------------------------"
    
    etrStep = "Footer row 38"
    ws.Range("A38:C38").Merge
    ws.Range("A38").Value = "Prices inclusive of VAT"
    ws.Range("A38").HorizontalAlignment = xlCenter
    ws.Range("A38").Font.Size = 8
    
    etrStep = "ETR Serial row 39"
    ws.Range("A39:C39").Merge
    ws.Range("A39").Value = "ETR Serial: (generated)"
    ws.Range("A39").HorizontalAlignment = xlCenter
    
    etrStep = "Thank you row 40"
    ws.Range("A40:C40").Merge
    ws.Range("A40").Value = "Thank you!"
    ws.Range("A40").HorizontalAlignment = xlCenter
    
    etrStep = "QR Code rows 41-43"
    ws.Range("A41:C43").Merge
    ws.Range("A41").Value = "[QR Code]"
    ws.Range("A41").HorizontalAlignment = xlCenter
    ws.Range("A41").VerticalAlignment = xlCenter
    ws.Range("A41").Interior.Color = RGB(240, 240, 240)
    
    etrStep = "End receipt row 44"
    ws.Range("A44:C44").Merge
    ws.Range("A44").Value = "*** END OF RECEIPT ***"
    ws.Range("A44").HorizontalAlignment = xlCenter
    
    etrStep = "PageSetup"
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
        .PrintArea = "A1:C44"
    End With
    On Error GoTo 0
    Exit Sub

ETRError:
    MsgBox "ETR Template error at [" & etrStep & "]:" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "BuildETRTemplate Failed"
    Err.Raise Err.Number, "BuildETRTemplate", "Failed at [" & etrStep & "]: " & Err.Description
End Sub

' --------------------------------------------------------------------------
' Step 10: Build TaxSummary Sheet
' --------------------------------------------------------------------------
Private Sub BuildTaxSummarySheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TaxSummary")

    With ws
        ' Title
        .Range("A1:G1").Merge
        .Range("A1").Value = "TAX SUMMARY REPORT"
        With .Range("A1")
            .Interior.Color = RGB(125, 60, 152)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 16
        End With

        ' Filters (rows 3-4)
        .Range("A3").Value = "Period:"
        .Range("B3").Value = "Monthly"
        On Error Resume Next
        .Range("B3").Validation.Delete
        .Range("B3").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Monthly,Quarterly,Annual"
        On Error GoTo 0

        .Range("C3").Value = "From:"
        .Range("E3").Value = "To:"

        .Range("A4").Value = "Jurisdiction:"
        .Range("B4").Value = "All"
        On Error Resume Next
        .Range("B4").Validation.Delete
        .Range("B4").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="All,Kenya,USA,UK"
        On Error GoTo 0

        ' Table headers (row 6)
        Dim headers As Variant
        headers = Array("Month", "Jurisdiction", "Total Revenue", "Total Tax Collected", "Tax Rate", "Invoices Count", "Outstanding")

        Dim col As Long
        For col = 1 To UBound(headers) + 1
            .Cells(6, col).Value = headers(col - 1)
        Next col

        ' Manual formatting removed - let ListObject TableStyle handle it

        ' Format as table
        Dim tbl As ListObject
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A6:G6"), , xlYes)
        tbl.Name = "tblTaxSummary"
        tbl.TableStyle = "TableStyleMedium2"

        ' Column widths
        .Columns("A:G").ColumnWidth = 18
        
        ' Formats
        .Columns("C:D").NumberFormat = "[$KES] #,##0.00" ' Revenue, Tax
        .Columns("G:G").NumberFormat = "[$KES] #,##0.00" ' Outstanding
        .Columns("E:E").NumberFormat = "0.00%" ' Tax Rate
        .Columns("F:F").NumberFormat = "0" ' Count
    End With
End Sub

' --------------------------------------------------------------------------
' Step 11: Build Dashboard Sheet
' --------------------------------------------------------------------------
Private Sub BuildDashboardSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")

    With ws
        .Cells.Clear
        On Error Resume Next
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        On Error GoTo 0

        ' Set column widths
        .Columns("A:J").ColumnWidth = 12

        ' Header banner (rows 1-3)
        .Range("A1:J3").Merge
        With .Range("A1")
            .Value = "Professional Billing System" & vbLf & "Dashboard & Navigation"
            .Rows.RowHeight = 35 ' Ensure enough height for 2 lines of large text
            .Interior.Color = RGB(27, 79, 114)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 24
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        ' KPI Cards (rows 5-9) - 5 cards side by side
        CreateKPICard ws, "A5:B9", "Total Revenue", "=SUMIF(Transactions!L:L,""<>Cancelled"",Transactions!I:I)", RGB(39, 174, 96)
        CreateKPICard ws, "C5:D9", "Outstanding", "=SUM(Transactions!K:K)", RGB(230, 126, 34)
        CreateKPICard ws, "E5:F9", "Invoices Issued", "=COUNTA(Transactions!A:A)-1", RGB(52, 152, 219)
        CreateKPICard ws, "G5:H9", "Overdue Count", "=COUNTIF(Transactions!L:L,""Overdue"")", RGB(231, 76, 60)
        CreateKPICard ws, "I5:J9", "Tax Collected", "=SUMIF(Transactions!L:L,""<>Cancelled"",Transactions!G:G)", RGB(125, 60, 152)

        ' Quick Actions header (rows 11-12)
        .Range("A11:J12").Merge
        With .Range("A11")
            .Value = "QUICK ACTIONS"
            .Interior.Color = RGB(230, 126, 34)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 14
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Navigation buttons (rows 13-14 and 16-17)
        CreateDashboardButton ws, "A13:B14", "NEW INVOICE", RGB(230, 126, 34)
        CreateDashboardButton ws, "C13:D14", "RECORD PAYMENT", RGB(39, 174, 96)
        CreateDashboardButton ws, "E13:F14", "GENERATE RECEIPT", RGB(17, 122, 101)
        CreateDashboardButton ws, "G13:H14", "ETR RECEIPT", RGB(27, 79, 114)
        CreateDashboardButton ws, "I13:J14", "EXPORT PDF", RGB(192, 57, 43)

        CreateDashboardButton ws, "A16:B17", "VIEW CUSTOMERS", RGB(52, 152, 219)
        CreateDashboardButton ws, "C16:D17", "VIEW PRODUCTS", RGB(52, 152, 219)
        CreateDashboardButton ws, "E16:F17", "TRANSACTIONS", RGB(17, 122, 101)
        CreateDashboardButton ws, "G16:H17", "TAX SUMMARY", RGB(125, 60, 152)
        
        ' Ensure button rows are tall enough for wrapped text
        .Rows("13:14").RowHeight = 25
        .Rows("16:17").RowHeight = 25
        CreateDashboardButton ws, "I16:J17", "SETTINGS", RGB(93, 109, 126)

        ' Recent Activity section (row 20+)
        .Range("A20:J20").Merge
        With .Range("A20")
            .Value = "RECENT ACTIVITY"
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(236, 240, 241)
        End With

        ' Recent Activity headers (row 21)
        .Range("A21").Value = "Invoice No"
        .Range("C21").Value = "Customer"
        .Range("E21").Value = "Date"
        .Range("G21").Value = "Amount"
        .Range("I21").Value = "Status"
        .Range("A21:J21").Font.Bold = True
        .Range("A21:J21").Interior.Color = RGB(189, 195, 199)

        ' Data rows 22-29 (8 rows) - will be populated by VBA
        Dim row As Long
        For row = 22 To 29
            If row Mod 2 = 0 Then
                .Range("A" & row & ":J" & row).Interior.Color = RGB(235, 245, 251)
            End If
        Next row
    End With
End Sub

' --------------------------------------------------------------------------
' Helper: Create KPI Card
' --------------------------------------------------------------------------
Private Sub CreateKPICard(ws As Worksheet, cellRange As String, label As String, formula As String, color As Long)
    On Error Resume Next
    With ws.Range(cellRange)
        .Merge
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
    ' Set composite formula
    ' Strip = from start
    If Left(formula, 1) = "=" Then formula = Mid(formula, 2)
    
    Dim fmt As String
    If InStr(formula, "COUNT") > 0 Then
        fmt = "0"
    Else
        fmt = "#,##0.00"
    End If
    
    ' Formula: ="Label" & CHAR(10) & TEXT(Formula, "Format")
    Dim finalFmla As String
    finalFmla = "=""" & label & """ & CHAR(10) & TEXT(" & formula & ", """ & fmt & """)"
    
    ws.Range(cellRange).Formula = finalFmla
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' Helper: Create Dashboard Button
' --------------------------------------------------------------------------
Private Sub CreateDashboardButton(ws As Worksheet, cellRange As String, label As String, color As Long)
    With ws.Range(cellRange)
        .Merge
        .Value = label
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With
End Sub

' --------------------------------------------------------------------------
' Step 12: Create All Named Ranges
' --------------------------------------------------------------------------
Private Sub CreateAllNamedRanges()
    On Error Resume Next

    ' Settings ranges
    AddNamedRange "rngCompanyName", "=Settings!$B$2"
    AddNamedRange "rngJurisdiction", "=Settings!$B$11"
    AddNamedRange "rngCurrency", "=Settings!$B$12"
    AddNamedRange "rngTaxTable", "=Settings!$A$16:$E$24"
    AddNamedRange "rngLastInvoice", "=Settings!$B$26"
    AddNamedRange "rngLastReceipt", "=Settings!$B$27"
    AddNamedRange "rngLastETR", "=Settings!$B$28"
    AddNamedRange "rngYearPrefix", "=Settings!$B$29"
    AddNamedRange "rngPaymentMethods", "=Settings!$A$32:$A$44"
    AddNamedRange "rngPaymentTerms", "=Settings!$B$46"

    ' Invoice Template ranges
    AddNamedRange "rngInvNumber", "=Invoice_Template!$B$8"
    AddNamedRange "rngInvDate", "=Invoice_Template!$B$9"
    AddNamedRange "rngInvDueDate", "=Invoice_Template!$B$10"
    AddNamedRange "rngInvCustomer", "=Invoice_Template!$E$9"
    AddNamedRange "rngInvLineItems", "=Invoice_Template!$A$15:$H$29"
    AddNamedRange "rngInvSubtotal", "=Invoice_Template!$H$31"
    AddNamedRange "rngInvTax", "=Invoice_Template!$H$33"
    AddNamedRange "rngInvTotal", "=Invoice_Template!$H$35"

    ' Receipt Template ranges
    AddNamedRange "rngRcptNumber", "=Receipt_Template!$B$8"
    AddNamedRange "rngRcptDate", "=Receipt_Template!$B$9"
    AddNamedRange "rngRcptInvoiceRef", "=Receipt_Template!$B$10"
    AddNamedRange "rngRcptAmountPaid", "=Receipt_Template!$B$17"
    AddNamedRange "rngRcptBalance", "=Receipt_Template!$B$20"

    ' ETR Template ranges
    AddNamedRange "rngETRNumber", "=ETR_Template!$B$7"
    AddNamedRange "rngETRDate", "=ETR_Template!$B$8"
    AddNamedRange "rngETRLineItems", "=ETR_Template!$A$13:$C$27"
    AddNamedRange "rngETRTotal", "=ETR_Template!$B$32"

    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' Helper: Add Named Range
' --------------------------------------------------------------------------
Private Sub AddNamedRange(rangeName As String, refersTo As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=rangeName, RefersTo:=refersTo
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' Step 13: Protect All Sheets
' --------------------------------------------------------------------------
Private Sub ProtectAllSheets()
    Dim ws As Worksheet

    ' Unprotect all first
    For Each ws In ThisWorkbook.Sheets
        ws.Unprotect "admin2026"
    Next ws

    ' Settings: Allow only input cells
    ThisWorkbook.Sheets("Settings").Protect Password:="admin2026", _
        UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True

    ' Templates: Allow data entry areas
    ThisWorkbook.Sheets("Invoice_Template").Protect Password:="admin2026", _
        UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Receipt_Template").Protect Password:="admin2026", _
        UserInterfaceOnly:=True
    ThisWorkbook.Sheets("ETR_Template").Protect Password:="admin2026", _
        UserInterfaceOnly:=True

    ' Data sheets: Allow sort/filter
    ThisWorkbook.Sheets("Customers").Protect Password:="admin2026", _
        UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True
    ThisWorkbook.Sheets("Products").Protect Password:="admin2026", _
        UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True
    ThisWorkbook.Sheets("Transactions").Protect Password:="admin2026", _
        UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True
    ThisWorkbook.Sheets("PaymentLog").Protect Password:="admin2026", _
        UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True

    ' Reports: Full protection
    ThisWorkbook.Sheets("TaxSummary").Protect Password:="admin2026", _
        UserInterfaceOnly:=True
    ThisWorkbook.Sheets("Dashboard").Protect Password:="admin2026", _
        UserInterfaceOnly:=True
End Sub

' --------------------------------------------------------------------------
' Helper: Format Header
' --------------------------------------------------------------------------
Private Sub FormatHeader(rng As Range)
    With rng
        .Interior.Color = RGB(27, 79, 114)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' --------------------------------------------------------------------------
' Helper: Format Data Table Header
' --------------------------------------------------------------------------
Private Sub FormatDataTableHeader(rng As Range)
    With rng
        .Interior.Color = RGB(27, 79, 114)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
End Sub

' --------------------------------------------------------------------------
' Step 14: Final Setup
' --------------------------------------------------------------------------
Private Sub FinalSetup()
    ' Set workbook properties
    With ThisWorkbook
        .Title = "Professional Billing System"
        .Subject = "Multi-Jurisdiction Invoice & Receipt Generator"
        .Author = "Built with VBA Workbook Builder"
    End With

    ' Add calculation to refresh on open
    Application.Calculate
End Sub

' --------------------------------------------------------------------------
' Step 15: Inject VBA Event Handlers
' --------------------------------------------------------------------------
Private Sub InjectSheetCode()
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    If Err.Number <> 0 Then
        MsgBox "Setup Warning: Code Injection Failed." & vbCrLf & vbCrLf & _
               "Please enable 'Trust access to the VBA project object model' in File > Options > Trust Center > Macro Settings.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' 1. Inject Dashboard Code
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    Dim modName As String
    modName = ws.CodeName
    
    ' If CodeName is empty (rare), try to find component
    If modName = "" Then
        Dim comp As Object
        For Each comp In vbProj.VBComponents
            If comp.Properties("Name").Value = "Dashboard" Then
                modName = comp.Name
                Exit For
            End If
        Next comp
    End If
    
    If modName <> "" Then
        With vbProj.VBComponents(modName).CodeModule
            .DeleteLines 1, .CountOfLines
            .AddFromString "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
                           "    On Error Resume Next" & vbCrLf & _
                           "    modDashboard.HandleDashboardClick Target" & vbCrLf & _
                           "End Sub" & vbCrLf & vbCrLf & _
                           "Private Sub Worksheet_Activate()" & vbCrLf & _
                           "    On Error Resume Next" & vbCrLf & _
                           "    modDashboard.RefreshDashboard" & vbCrLf & _
                           "End Sub"
        End With
    End If

    ' 2. Inject Invoice_Template Code
    Set ws = ThisWorkbook.Sheets("Invoice_Template")
    modName = ws.CodeName
    
    If modName = "" Then
        For Each comp In vbProj.VBComponents
            If comp.Properties("Name").Value = "Invoice_Template" Then
                modName = comp.Name
                Exit For
            End If
        Next comp
    End If
    
    If modName <> "" Then
        With vbProj.VBComponents(modName).CodeModule
            ' Only add if not present
            If Not .Find("Worksheet_Activate", 1, 1, .CountOfLines, 1) Then
                .AddFromString "" & vbCrLf & _
                               "Private Sub Worksheet_Activate()" & vbCrLf & _
                               "    On Error Resume Next" & vbCrLf & _
                               "    modInvoice.UpdateLogo Me" & vbCrLf & _
                               "End Sub"
            End If
        End With
    End If

    ' 3. Inject ThisWorkbook Code
    With vbProj.VBComponents("ThisWorkbook").CodeModule
        ' Only add if not present
        If Not .Find("Workbook_Open", 1, 1, .CountOfLines, 1) Then
             .AddFromString "" & vbCrLf & _
                            "Private Sub Workbook_Open()" & vbCrLf & _
                            "    On Error Resume Next" & vbCrLf & _
                            "    modDashboard.RefreshDashboard" & vbCrLf & _
                            "    modDashboard.NavigateTo ""Dashboard""" & vbCrLf & _
                            "End Sub"
        End If
    End With
End Sub
