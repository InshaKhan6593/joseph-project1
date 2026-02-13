Attribute VB_Name = "modDashboard"
Option Explicit

' ==============================================================================
' Module: modDashboard.bas
' Purpose: Dashboard management and navigation
' Created: 2026-02-13
' Dependencies: All modules
'
' ACTUAL DASHBOARD LAYOUT (from workbook):
'   Rows 1-3: Header banner (merged A1:J3)
'   Rows 5-9: KPI cards (merged cell pairs)
'   Row 11-12: "QUICK ACTIONS" header
'   Row 13-14: Button row 1 (merged cell pairs)
'   Row 16-17: Button row 2 (merged cell pairs)
'   Row 20: RECENT ACTIVITY header
'   Row 21: Column headers
'   Row 22+: Recent activity data
'
' BUTTON MAP (merged cells, NOT shapes):
'   A13:B14 = NEW INVOICE
'   C13:D14 = RECORD PAYMENT
'   E13:F14 = GENERATE RECEIPT
'   G13:H14 = ETR RECEIPT
'   I13:J14 = EXPORT PDF
'   A16:B17 = VIEW CUSTOMERS
'   C16:D17 = VIEW PRODUCTS
'   E16:F17 = TRANSACTIONS
'   G16:H17 = TAX SUMMARY
'   I16:J17 = SETTINGS
' ==============================================================================

' --------------------------------------------------------------------------
' 1. HandleDashboardClick(Target As Range)
'    Called from Dashboard sheet's Worksheet_SelectionChange event.
'    Maps clicked cell region to the correct macro.
' --------------------------------------------------------------------------
Public Sub HandleDashboardClick(Target As Range)
    On Error GoTo ErrHandler
    
    Dim r As Long, c As Long
    r = Target.Row
    c = Target.Column
    
    ' Row 13-14: First button row
    If r >= 13 And r <= 14 Then
        Select Case True
            Case c >= 1 And c <= 2:   btnNewInvoice_Click
            Case c >= 3 And c <= 4:   btnRecordPayment_Click
            Case c >= 5 And c <= 6:   btnGenerateReceipt_Click
            Case c >= 7 And c <= 8:   btnETRReceipt_Click
            Case c >= 9 And c <= 10:  btnExportPDF_Click
        End Select
    End If
    
    ' Row 16-17: Second button row
    If r >= 16 And r <= 17 Then
        Select Case True
            Case c >= 1 And c <= 2:   btnViewCustomers_Click
            Case c >= 3 And c <= 4:   btnViewProducts_Click
            Case c >= 5 And c <= 6:   btnTransactions_Click
            Case c >= 7 And c <= 8:   btnTaxSummary_Click
            Case c >= 9 And c <= 10:  btnSettings_Click
        End Select
    End If
    
    Exit Sub
ErrHandler:
    ErrorHandler "HandleDashboardClick", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 2. RefreshDashboard()
'    Force-recalculates KPI formulas. Dashboard is NOT protected.
' --------------------------------------------------------------------------
Public Sub RefreshDashboard()
    On Error GoTo ErrHandler
    
    TogglePerformance True
    
    Dim wsDash As Worksheet
    Set wsDash = SafeSheetRef("Dashboard")
    If wsDash Is Nothing Then
        TogglePerformance False
        Exit Sub
    End If
    
    ' Force recalc of KPI area
    On Error Resume Next
    wsDash.Range("A5:J9").Calculate
    On Error GoTo ErrHandler
    
    ' Update Recent Activity
    UpdateRecentActivity
    
    ' Check overdue invoices
    CheckOverdueInvoices
    
    TogglePerformance False
    Exit Sub
ErrHandler:
    TogglePerformance False
    ErrorHandler "RefreshDashboard", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 3. NavigateTo(sheetName As String)
' --------------------------------------------------------------------------
Public Sub NavigateTo(sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SafeSheetRef(sheetName)
    If Not ws Is Nothing Then
        ws.Activate
        ws.Range("A1").Select
    End If
    Exit Sub
ErrHandler:
    ErrorHandler "NavigateTo", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 4. Button Click Handlers
' --------------------------------------------------------------------------
Public Sub btnNewInvoice_Click()
    modInvoice.GenerateInvoice
End Sub

Public Sub btnRecordPayment_Click()
    modForms.ShowPaymentEntry
End Sub

Public Sub btnGenerateReceipt_Click()
    Dim paymentID As String
    paymentID = InputBox("Enter Payment ID:", "Generate Receipt")
    If paymentID <> "" Then modReceipt.GenerateReceiptFromPayment paymentID
End Sub

Public Sub btnETRReceipt_Click()
    If LCase(GetSetting("Jurisdiction")) <> "kenya" Then
        MsgBox "ETR is only available for Kenya jurisdiction.", vbExclamation
        Exit Sub
    End If
    Dim invoiceNo As String
    invoiceNo = InputBox("Enter Invoice Number:", "Generate ETR")
    If invoiceNo <> "" Then modETR.GenerateETR invoiceNo
End Sub

Public Sub btnExportPDF_Click()
    Dim docType As String
    docType = InputBox("Enter document type (invoice/receipt/etr):", "Export PDF")
    If docType <> "" Then modExport.ExportToPDF docType
End Sub

Public Sub btnViewCustomers_Click()
    NavigateTo "Customers"
End Sub

Public Sub btnViewProducts_Click()
    NavigateTo "Products"
End Sub

Public Sub btnTransactions_Click()
    NavigateTo "Transactions"
End Sub

Public Sub btnTaxSummary_Click()
    NavigateTo "TaxSummary"
End Sub

Public Sub btnSettings_Click()
    NavigateTo "Settings"
End Sub

' --------------------------------------------------------------------------
' 5. CheckOverdueInvoices() As Long
' --------------------------------------------------------------------------
Public Function CheckOverdueInvoices() As Long
    On Error GoTo ErrHandler
    
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    If wsTrans Is Nothing Then Exit Function
    
    Dim count As Long
    count = 0
    
    Dim lastRow As Long
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    Dim i As Long
    For i = 2 To lastRow
        Dim status As String
        status = CStr(wsTrans.Cells(i, 12).Value)
        
        If status = "Pending" Or status = "Partial" Then
            If IsDate(wsTrans.Cells(i, 5).Value) Then
                If CDate(wsTrans.Cells(i, 5).Value) < Date Then
                    wsTrans.Cells(i, 12).Value = "Overdue"
                    count = count + 1
                End If
            End If
        End If
    Next i
    
    CheckOverdueInvoices = count
    Exit Function
ErrHandler:
    ErrorHandler "CheckOverdueInvoices", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' 6. UpdateRecentActivity()
'    Fills Dashboard Recent Activity section (rows 22-29)
' --------------------------------------------------------------------------
Public Sub UpdateRecentActivity()
    On Error GoTo ErrHandler
    
    Dim wsDash As Worksheet
    Set wsDash = SafeSheetRef("Dashboard")
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    
    ' Clear data rows 22-29
    On Error Resume Next
    wsDash.Range("A22:J29").ClearContents
    On Error GoTo ErrHandler
    
    If wsTrans Is Nothing Then Exit Sub
    
    Dim lastTransRow As Long
    lastTransRow = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
    
    If lastTransRow < 2 Then
        wsDash.Range("A22").Value = "(No transactions yet)"
        Exit Sub
    End If
    
    ' Get last 8 transactions (most recent first)
    Dim dashRow As Long
    dashRow = 22
    
    Dim i As Long
    For i = lastTransRow To 2 Step -1
        If dashRow > 29 Then Exit For
        
        wsDash.Cells(dashRow, 1).Value = wsTrans.Cells(i, 1).Value  ' Invoice No
        wsDash.Cells(dashRow, 3).Value = wsTrans.Cells(i, 3).Value  ' Customer
        wsDash.Cells(dashRow, 5).Value = wsTrans.Cells(i, 4).Value  ' Date
        wsDash.Cells(dashRow, 7).Value = wsTrans.Cells(i, 9).Value  ' Amount
        wsDash.Cells(dashRow, 9).Value = wsTrans.Cells(i, 12).Value ' Status
        
        dashRow = dashRow + 1
    Next i
    Exit Sub
ErrHandler:
    ErrorHandler "UpdateRecentActivity", Err.Number, Err.Description
End Sub
