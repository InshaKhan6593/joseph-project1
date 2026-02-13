Attribute VB_Name = "modPayment"
Option Explicit

' ==============================================================================
' Module: modPayment.bas — Payment recording (NO PROTECTION)
' PaymentLog columns: A=PaymentID, B=InvNo, C=CustID, D=Date, E=Amount,
'                     F=Method, G=Reference, H=ReceivedBy, I=Notes
' ==============================================================================

' --------------------------------------------------------------------------
' RecordPayment — Main payment recording function
' --------------------------------------------------------------------------
Public Sub RecordPayment(invoiceNo As String, amount As Double, _
    paymentMethod As String, Optional refNo As String = "", Optional notes As String = "")
    On Error GoTo ErrHandler
    
    If amount <= 0 Then MsgBox "Amount must be > 0", vbExclamation: Exit Sub
    
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    Dim wsPay As Worksheet
    Set wsPay = SafeSheetRef("PaymentLog")
    
    ' Find invoice in Transactions
    Dim transRow As Long: transRow = 0
    Dim i As Long
    For i = 2 To wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
        If CStr(wsTrans.Cells(i, 1).Value) = invoiceNo Then transRow = i: Exit For
    Next i
    If transRow = 0 Then MsgBox "Invoice not found.", vbExclamation: Exit Sub
    
    TogglePerformance True
    
    ' Write to PaymentLog
    Dim payRow As Long
    payRow = modUtilities.GetNextRow(wsPay, 1)
    
    Dim paymentID As String
    paymentID = "PAY-" & Year(Date) & "-" & Format(payRow - 1, "0000")
    
    With wsPay
        .Cells(payRow, 1).Value = paymentID
        .Cells(payRow, 2).Value = invoiceNo
        .Cells(payRow, 3).Value = wsTrans.Cells(transRow, 2).Value  ' CustID
        .Cells(payRow, 4).Value = Date
        .Cells(payRow, 5).Value = amount
        .Cells(payRow, 6).Value = paymentMethod
        .Cells(payRow, 7).Value = refNo
        .Cells(payRow, 8).Value = Application.UserName
        .Cells(payRow, 9).Value = notes
    End With
    
    ' Update Transactions
    Dim currentPaid As Double
    currentPaid = CDbl(Val(wsTrans.Cells(transRow, 10).Value))
    Dim newPaid As Double
    newPaid = currentPaid + amount
    Dim grandTotal As Double
    grandTotal = CDbl(Val(wsTrans.Cells(transRow, 9).Value))
    
    wsTrans.Cells(transRow, 10).Value = newPaid
    wsTrans.Cells(transRow, 11).Value = grandTotal - newPaid
    
    ' Update status
    If newPaid >= grandTotal Then
        wsTrans.Cells(transRow, 12).Value = "Paid"
    ElseIf newPaid > 0 Then
        wsTrans.Cells(transRow, 12).Value = "Partial"
    End If
    
    AuditLog "PAYMENT_RECORDED", paymentID & " - " & invoiceNo & " - " & FormatCurrency(amount)
    TogglePerformance False
    MsgBox "Payment " & paymentID & " recorded!" & vbCrLf & "Balance: " & FormatCurrency(grandTotal - newPaid), vbInformation
    Exit Sub
ErrHandler:
    TogglePerformance False
    ErrorHandler "RecordPayment", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' GetPaymentHistory
' --------------------------------------------------------------------------
Public Function GetPaymentHistory(invoiceNo As String) As Collection
    On Error GoTo ErrHandler
    Dim col As New Collection
    Dim wsPay As Worksheet
    Set wsPay = SafeSheetRef("PaymentLog")
    Dim lastRow As Long
    lastRow = wsPay.Cells(wsPay.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsPay.Cells(i, 2).Value) = invoiceNo Then
            Dim info As String
            info = CStr(wsPay.Cells(i, 1).Value) & " | " & _
                   FormatDate(CDate(wsPay.Cells(i, 4).Value)) & " | " & _
                   FormatCurrency(CDbl(wsPay.Cells(i, 5).Value)) & " | " & _
                   CStr(wsPay.Cells(i, 6).Value)
            col.Add info
        End If
    Next i
    Set GetPaymentHistory = col
    Exit Function
ErrHandler:
    ErrorHandler "GetPaymentHistory", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' ShowPaymentForm
' --------------------------------------------------------------------------
Public Sub ShowPaymentForm(Optional invoiceNo As String = "")
    On Error GoTo ErrHandler
    modForms.ShowPaymentEntry invoiceNo
    Exit Sub
ErrHandler:
    ErrorHandler "ShowPaymentForm", Err.Number, Err.Description
End Sub
