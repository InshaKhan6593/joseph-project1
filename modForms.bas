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
' 1. ShowCustomerPicker() — Activates Customer Sheet
' --------------------------------------------------------------------------
Public Sub ShowCustomerPicker()
    On Error GoTo ErrHandler
    g_selectingForInvoice = True
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Customers")
    If ws Is Nothing Then Exit Sub
    
    ws.Activate
    MsgBox "Please DOUBLE-CLICK the customer you want to select.", vbInformation, "Select Customer"
    Exit Sub
ErrHandler:
    ErrorHandler "ShowCustomerPicker", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 1b. SelectCustomerFromSheet(row) — Called by DoubleClick Event
' --------------------------------------------------------------------------
Public Sub SelectCustomerFromSheet(row As Long)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Customers")
    
    Dim custID As String
    custID = CStr(ws.Cells(row, 1).Value)
    
    If custID = "" Then Exit Sub
    
    If g_selectingForInvoice Then
        g_selectedCustomerID = custID
        ' Return to Invoice
        Dim wsInv As Worksheet
        Set wsInv = SafeSheetRef("Invoice_Template")
        wsInv.Activate
        
        ' Populate
        PopulateInvoiceCustomer custID
        
        g_selectingForInvoice = False
    Else
        MsgBox "Customer: " & custID
    End If
    Exit Sub
ErrHandler:
    ErrorHandler "SelectCustomerFromSheet", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 2. ShowProductPicker(wsInv) — Replaces frmProductSelect
'    Lets user add line items to invoice one at a time
' --------------------------------------------------------------------------
' --------------------------------------------------------------------------
' 2. ShowProductPicker(wsInv) — Activates Products Sheet
' --------------------------------------------------------------------------
Public Sub ShowProductPicker(wsInv As Worksheet)
    On Error GoTo ErrHandler
    g_selectingForInvoice = True
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Products")
    If ws Is Nothing Then Exit Sub
    
    ws.Activate
    MsgBox "Please DOUBLE-CLICK the product you want to add.", vbInformation, "Select Product"
    Exit Sub
ErrHandler:
    ErrorHandler "ShowProductPicker", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 2b. SelectProductFromSheet(row) — Called by DoubleClick
' --------------------------------------------------------------------------
Public Sub SelectProductFromSheet(row As Long)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Products")
    
    Dim sku As String
    sku = CStr(ws.Cells(row, 1).Value)
    
    If sku = "" Then Exit Sub
    
    If g_selectingForInvoice Then
        ' Ask for quantity
        Dim qtyAsStr As String
        qtyAsStr = InputBox("Enter quantity for " & sku & ":", "Quantity", "1")
        If qtyAsStr = "" Then Exit Sub
        
        Dim qty As Double
        qty = CDbl(Val(qtyAsStr))
        
        ' Return to Invoice
        Dim wsInv As Worksheet
        Set wsInv = SafeSheetRef("Invoice_Template")
        wsInv.Activate
        
        ' Add item
        modProduct.AddLineItem wsInv, modProduct.GetNextLineItemRow(wsInv), sku, qty, 0
        
        ' Ask if more
        If MsgBox("Product added. Do you want to add another?", vbYesNo + vbQuestion, "Add More?") = vbYes Then
            ws.Activate
        Else
            g_selectingForInvoice = False
        End If
    Else
        MsgBox "Product: " & sku
    End If
    Exit Sub
ErrHandler:
    ErrorHandler "SelectProductFromSheet", Err.Number, Err.Description
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
    
    ' 3. Get payment method using Selection Form
    Dim method As String
    Dim items As New Collection
    Dim rngMethods As Range
    Dim cell As Range
    
    On Error Resume Next
    Set rngMethods = ThisWorkbook.Names("rngPaymentMethods").RefersToRange
    On Error GoTo ErrHandler
    
    If Not rngMethods Is Nothing Then
        For Each cell In rngMethods
            If Trim(cell.Value) <> "" Then items.Add cell.Value
        Next cell
    Else
        ' Fallback
        items.Add "Cash": items.Add "M-Pesa": items.Add "Bank Transfer": items.Add "Cheque"
    End If
    
    method = modFormBuilder.ShowSelectionDialog("Select Payment Method", items)
    If method = "" Then Exit Sub ' Cancelled


    
    ' Get reference
    Dim refNo As String
    refNo = InputBox("Enter reference number (optional):", "Reference")
    
    ' Record payment
    modPayment.RecordPayment invoiceNo, amount, method, refNo, ""
    Exit Sub
ErrHandler:
    ErrorHandler "ShowPaymentEntry", Err.Number, Err.Description
End Sub
