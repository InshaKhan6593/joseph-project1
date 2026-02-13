Attribute VB_Name = "modForms"
Option Explicit

' ==============================================================================
' Module: modForms.bas
' Purpose: Replaces UserForms with simple InputBox/MsgBox dialogs
'          This avoids the compilation errors from dynamically-created controls
'          that can't have event handlers.
'
' IMPORTANT: This module replaces frmCustomerSelect, frmProductSelect,
'            and frmPaymentEntry. Delete those forms from the VBA Editor.
' ==============================================================================

' --------------------------------------------------------------------------
' 1. ShowCustomerPicker() — Replaces frmCustomerSelect
'    Returns customer ID or "" if cancelled
' --------------------------------------------------------------------------
Public Function ShowCustomerPicker() As String
    On Error GoTo ErrHandler
    ShowCustomerPicker = ""
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Customers")
    If ws Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No customers found.", vbExclamation: Exit Function
    
    ' Build list string from sheet data directly
    Dim list As String
    list = "Available Customers:" & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) <> "" Then
            Dim status As String
            status = LCase(CStr(ws.Cells(i, 11).Value))
            If status = "active" Or status = "" Then
                list = list & CStr(ws.Cells(i, 1).Value) & " - " & CStr(ws.Cells(i, 2).Value) & vbCrLf
            End If
        End If
    Next i
    
    list = list & vbCrLf & "Enter Customer ID (e.g. C001):"
    
    Dim custID As String
    custID = InputBox(list, "Select Customer")
    
    If custID = "" Then Exit Function
    
    ' Validate
    custID = UCase(Trim(custID))
    For i = 2 To lastRow
        If UCase(CStr(ws.Cells(i, 1).Value)) = custID Then
            ShowCustomerPicker = CStr(ws.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
    
    MsgBox "Customer ID '" & custID & "' not found. Please try again.", vbExclamation
    ShowCustomerPicker = ShowCustomerPicker() ' Retry
    Exit Function
ErrHandler:
    ErrorHandler "ShowCustomerPicker", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' 2. ShowProductPicker(wsInv) — Replaces frmProductSelect
'    Lets user add line items to invoice one at a time
' --------------------------------------------------------------------------
Public Sub ShowProductPicker(wsInv As Worksheet)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Products")
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then MsgBox "No products found.", vbExclamation: Exit Sub
    
    Dim lineNum As Long
    lineNum = 1
    
    Do
        ' Build product list from sheet data directly
        Dim list As String
        list = "Available Products:" & vbCrLf & vbCrLf
        Dim i As Long
        For i = 2 To lastRow
            If CStr(ws.Cells(i, 1).Value) <> "" Then
                Dim pStatus As String
                pStatus = LCase(CStr(ws.Cells(i, 8).Value))
                If pStatus = "active" Or pStatus = "" Then
                    list = list & CStr(ws.Cells(i, 1).Value) & " - " & CStr(ws.Cells(i, 2).Value) & _
                           " (" & modUtilities.FormatCurrency(CDbl(Val(ws.Cells(i, 5).Value))) & "/" & CStr(ws.Cells(i, 6).Value) & ")" & vbCrLf
                End If
            End If
        Next i
        
        list = list & vbCrLf & "Enter Product SKU (or leave blank to finish):"
        
        Dim sku As String
        sku = InputBox(list, "Add Product - Line " & lineNum)
        
        If sku = "" Then Exit Do
        If lineNum > 15 Then MsgBox "Maximum 15 line items.", vbExclamation: Exit Do
        
        ' Get quantity
        Dim qtyStr As String
        qtyStr = InputBox("Enter quantity for " & sku & ":", "Quantity")
        If qtyStr = "" Then Exit Do
        
        Dim qty As Double
        qty = CDbl(Val(qtyStr))
        If qty <= 0 Then qty = 1
        
        ' Add line item
        modProduct.AddLineItem wsInv, lineNum, UCase(Trim(sku)), qty, 0
        lineNum = lineNum + 1
    Loop
    Exit Sub
ErrHandler:
    ErrorHandler "ShowProductPicker", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 3. ShowPaymentEntry(invoiceNo) — Replaces frmPaymentEntry
' --------------------------------------------------------------------------
Public Sub ShowPaymentEntry(Optional invoiceNo As String = "")
    On Error GoTo ErrHandler
    
    ' Get invoice number
    If invoiceNo = "" Then
        invoiceNo = InputBox("Enter Invoice Number (e.g. INV-2026-0001):", "Record Payment")
        If invoiceNo = "" Then Exit Sub
    End If
    
    ' Get amount
    Dim amtStr As String
    amtStr = InputBox("Enter payment amount:", "Payment Amount")
    If amtStr = "" Then Exit Sub
    
    Dim amount As Double
    amount = CDbl(Val(amtStr))
    If amount <= 0 Then MsgBox "Amount must be > 0", vbExclamation: Exit Sub
    
    ' Get payment method
    Dim method As String
    method = InputBox("Enter payment method:" & vbCrLf & _
                      "Cash, M-Pesa, Bank Transfer, Credit Card, Debit Card, Cheque, Other", _
                      "Payment Method")
    If method = "" Then method = "Cash"
    
    ' Get reference
    Dim refNo As String
    refNo = InputBox("Enter reference number (optional):", "Reference")
    
    ' Record payment
    modPayment.RecordPayment invoiceNo, amount, method, refNo, ""
    Exit Sub
ErrHandler:
    ErrorHandler "ShowPaymentEntry", Err.Number, Err.Description
End Sub
