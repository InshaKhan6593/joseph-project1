Attribute VB_Name = "modReceipt"
Option Explicit

' ==============================================================================
' Module: modReceipt.bas — Receipt generation (NO PROTECTION)
'
' Receipt_Template layout (from Plan A):
'   B8  = Receipt No
'   B9  = Date
'   B10 = Invoice Reference
'   B11 = Customer
'   B12 = Tax ID
'   B16 = Amount Due
'   B17 = Amount Paid
'   B18 = Payment Method
'   B19 = Reference No
'   B20 = Balance (=B16-B17)
' ==============================================================================

' --------------------------------------------------------------------------
' GenerateReceipt(invoiceNo) — From invoice
' --------------------------------------------------------------------------
Public Sub GenerateReceipt(invoiceNo As String)
    On Error GoTo ErrHandler
    
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    Dim wsRcpt As Worksheet
    Set wsRcpt = SafeSheetRef("Receipt_Template")
    
    modUtilities.UnprotectSheet wsRcpt.Name
    
    Dim transRow As Long: transRow = 0
    Dim i As Long
    For i = 2 To wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
        If CStr(wsTrans.Cells(i, 1).Value) = invoiceNo Then transRow = i: Exit For
    Next i
    If transRow = 0 Then MsgBox "Invoice not found.", vbExclamation: Exit Sub
    
    wsRcpt.Range("B8").Value = modNumbering.GetNextReceiptNumber()
    wsRcpt.Range("B9").Value = Date
    wsRcpt.Range("B10").Value = invoiceNo ' Invoice Reference
    wsRcpt.Range("B11").Value = wsTrans.Cells(transRow, 3).Value ' Customer
    wsRcpt.Range("B12").Value = "" ' Tax ID
    wsRcpt.Range("B16").Value = wsTrans.Cells(transRow, 9).Value ' Amount Due
    wsRcpt.Range("B17").Value = wsTrans.Cells(transRow, 10).Value ' Amount Paid
    wsRcpt.Range("B18").Value = "" ' Method
    wsRcpt.Range("B19").Value = "" ' Ref
    wsRcpt.Range("B20").Value = wsTrans.Cells(transRow, 11).Value
    
    AuditLog "RECEIPT", wsRcpt.Range("B8").Value & " for " & invoiceNo
    
    modUtilities.ProtectSheet wsRcpt.Name
    
    wsRcpt.Activate
    MsgBox "Receipt " & wsRcpt.Range("B8").Value & " generated!", vbInformation
    Exit Sub
ErrHandler:
    modUtilities.ProtectSheet wsRcpt.Name
    ErrorHandler "GenerateReceipt", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' GenerateReceiptFromPayment(paymentID)
' --------------------------------------------------------------------------
Public Sub GenerateReceiptFromPayment(paymentID As String)
    On Error GoTo ErrHandler
    
    Dim wsPay As Worksheet
    Set wsPay = SafeSheetRef("PaymentLog")
    Dim wsRcpt As Worksheet
    Set wsRcpt = SafeSheetRef("Receipt_Template")
    
    Dim payRow As Long: payRow = 0
    Dim i As Long
    For i = 2 To wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row
        If CStr(wsPay.Cells(i, 1).Value) = paymentID Then payRow = i: Exit For
    Next i
    If payRow = 0 Then MsgBox "Payment not found.", vbExclamation: Exit Sub

    modUtilities.UnprotectSheet wsRcpt.Name

    wsRcpt.Range("B8").Value = modNumbering.GetNextReceiptNumber()
    wsRcpt.Range("B9").Value = Date
    wsRcpt.Range("B10").Value = wsPay.Cells(payRow, 2).Value
    wsRcpt.Range("B11").Value = wsPay.Cells(payRow, 3).Value
    wsRcpt.Range("B16").Value = wsPay.Cells(payRow, 5).Value
    wsRcpt.Range("B17").Value = wsPay.Cells(payRow, 5).Value
    wsRcpt.Range("B18").Value = wsPay.Cells(payRow, 6).Value
    wsRcpt.Range("B19").Value = wsPay.Cells(payRow, 7).Value
    wsRcpt.Range("B20").Value = 0

    modUtilities.ProtectSheet wsRcpt.Name

    AuditLog "RECEIPT", wsRcpt.Range("B8").Value & " for payment " & paymentID
    wsRcpt.Activate
    MsgBox "Receipt generated!", vbInformation
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet "Receipt_Template"
    On Error GoTo 0
    ErrorHandler "GenerateReceiptFromPayment", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' ClearReceiptTemplate
' --------------------------------------------------------------------------
Public Sub ClearReceiptTemplate()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Receipt_Template")
    modUtilities.UnprotectSheet ws.Name
    ws.Range("B8:B12").ClearContents
    ws.Range("B16:B20").ClearContents
    modUtilities.ProtectSheet ws.Name
    On Error GoTo 0
End Sub
