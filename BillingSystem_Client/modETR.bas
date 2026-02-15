Attribute VB_Name = "modETR"
Option Explicit

' ==============================================================================
' Module: modETR.bas — ETR thermal receipt (Kenya) (NO PROTECTION)
'
' ETR_Template layout (from Plan A):
'   A1  = Company Name
'   A4  = KRA PIN
'   A7  = Receipt No (after "Receipt No:" label)
'   A8  = Date
'   A9  = Cashier
'   A13:C27 = Line items (item | qty | amount)
'   C29 = Subtotal
'   C30 = VAT
'   C32 = Total
'   C34 = Payment method
'   C35 = Amount Tendered
'   C36 = Change
'   A39 = ETR Serial
' ==============================================================================

' --------------------------------------------------------------------------
' GenerateETR(invoiceNo)
' --------------------------------------------------------------------------
Public Sub GenerateETR(invoiceNo As String)
    On Error GoTo ErrHandler
    
    If LCase(GetSetting("Jurisdiction")) <> "kenya" Then
        MsgBox "ETR is only available for Kenya.", vbExclamation
        Exit Sub
    End If
    
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    Dim wsETR As Worksheet
    Set wsETR = SafeSheetRef("ETR_Template")
    Dim wsInv As Worksheet
    Set wsInv = SafeSheetRef("Invoice_Template")
    
    ' Find invoice
    Dim transRow As Long: transRow = 0
    Dim i As Long
    For i = 2 To wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
        If CStr(wsTrans.Cells(i, 1).Value) = invoiceNo Then transRow = i: Exit For
    Next i
    If transRow = 0 Then MsgBox "Invoice not found.", vbExclamation: Exit Sub
    
    ClearETRTemplate

    modUtilities.UnprotectSheet wsETR.Name

    ' Header info — rows 1-4 are formula-linked to Settings,
    ' so we only write rows 7-9 and line items
    Dim etrNum As String
    etrNum = modNumbering.GetNextETRNumber()

    wsETR.Cells(7, 1).Value = "Receipt No: " & etrNum
    wsETR.Cells(8, 1).Value = "Date: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    wsETR.Cells(9, 1).Value = "Cashier: " & Application.UserName

    ' Copy line items from invoice template (rows 15-29) to ETR (rows 13-27)
    Dim etrRow As Long: etrRow = 13
    For i = 15 To 29
        If wsInv.Cells(i, 3).Value <> "" Then
            wsETR.Cells(etrRow, 1).Value = wsInv.Cells(i, 3).Value  ' Item name
            wsETR.Cells(etrRow, 2).Value = wsInv.Cells(i, 4).Value  ' Qty
            wsETR.Cells(etrRow, 3).Value = wsInv.Cells(i, 8).Value  ' Amount
            etrRow = etrRow + 1
        End If
    Next i

    ' Totals
    Dim subtot As Double, tax As Double, total As Double
    subtot = CDbl(Val(wsTrans.Cells(transRow, 6).Value))
    tax = CDbl(Val(wsTrans.Cells(transRow, 7).Value))
    total = CDbl(Val(wsTrans.Cells(transRow, 9).Value))

    wsETR.Cells(29, 3).Value = subtot
    wsETR.Cells(30, 3).Value = tax
    wsETR.Cells(32, 3).Value = total

    ' ETR Serial
    wsETR.Cells(39, 1).Value = "ETR Serial: " & etrNum

    modUtilities.ProtectSheet wsETR.Name

    AuditLog "ETR_GENERATED", etrNum & " for " & invoiceNo
    wsETR.Activate
    MsgBox "ETR Receipt " & etrNum & " generated!", vbInformation
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet "ETR_Template"
    On Error GoTo 0
    ErrorHandler "GenerateETR", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' ClearETRTemplate
' --------------------------------------------------------------------------
Public Sub ClearETRTemplate()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SafeSheetRef("ETR_Template")
    modUtilities.UnprotectSheet ws.Name
    ws.Range("A13:C27").ClearContents
    ws.Range("C29").ClearContents
    ws.Range("C30").ClearContents
    ws.Range("C32").ClearContents
    ws.Range("C34:C36").ClearContents
    modUtilities.ProtectSheet ws.Name
    On Error GoTo 0
End Sub
