Attribute VB_Name = "modCustomer"
Option Explicit

' ==============================================================================
' Module: modCustomer.bas — Customer management (NO PROTECTION, NO TABLES)
'
' Customers sheet layout (row 1 = headers, data starts row 2):
'   A=Cust_ID  B=Company Name  C=Contact Person  D=Email  E=Phone
'   F=Billing Address  G=City  H=Country  I=Tax ID  J=Credit Terms
'   K=Status  L=Notes
' ==============================================================================

Public g_selectedCustomerID As String

' --------------------------------------------------------------------------
' LookupCustomer — Search by ID or Name using direct cell refs
' --------------------------------------------------------------------------
Public Function LookupCustomer(identifier As String) As Object
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Customers")
    If ws Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    Dim i As Long, foundRow As Long
    foundRow = 0
    For i = 2 To lastRow
        If LCase(CStr(ws.Cells(i, 1).Value)) = LCase(identifier) Or _
           LCase(CStr(ws.Cells(i, 2).Value)) = LCase(identifier) Then
            foundRow = i
            Exit For
        End If
    Next i
    
    If foundRow > 0 Then
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "ID", CStr(ws.Cells(foundRow, 1).Value)
        dict.Add "Name", CStr(ws.Cells(foundRow, 2).Value)
        dict.Add "Contact", CStr(ws.Cells(foundRow, 3).Value)
        dict.Add "Email", CStr(ws.Cells(foundRow, 4).Value)
        dict.Add "Phone", CStr(ws.Cells(foundRow, 5).Value)
        dict.Add "Address", CStr(ws.Cells(foundRow, 6).Value)
        dict.Add "City", CStr(ws.Cells(foundRow, 7).Value)
        dict.Add "Country", CStr(ws.Cells(foundRow, 8).Value)
        dict.Add "TaxID", CStr(ws.Cells(foundRow, 9).Value)
        dict.Add "Terms", CStr(ws.Cells(foundRow, 10).Value)
        dict.Add "Status", CStr(ws.Cells(foundRow, 11).Value)
        dict.Add "Notes", CStr(ws.Cells(foundRow, 12).Value)
        Set LookupCustomer = dict
    End If
    Exit Function
ErrHandler:
    ErrorHandler "LookupCustomer", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' PopulateInvoiceCustomer — Writes customer info to Invoice_Template
' --------------------------------------------------------------------------
Public Sub PopulateInvoiceCustomer(custID As String)
    On Error GoTo ErrHandler
    Dim custData As Object
    Set custData = LookupCustomer(custID)
    If custData Is Nothing Then
        MsgBox "Customer not found.", vbExclamation
        Exit Sub
    End If
    
    Dim wsInv As Worksheet
    Set wsInv = SafeSheetRef("Invoice_Template")

    modUtilities.UnprotectSheet wsInv.Name

    wsInv.Range("E9").Value = custData("Name")
    wsInv.Range("E10").Value = custData("Address") & ", " & custData("City")
    wsInv.Range("E11").Value = "Tax ID: " & custData("TaxID")

    ' Set payment terms from customer record
    If custData("Terms") <> "" Then wsInv.Range("B11").Value = custData("Terms")

    modUtilities.ProtectSheet wsInv.Name
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet "Invoice_Template"
    On Error GoTo 0
    ErrorHandler "PopulateInvoiceCustomer", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' GetCustomerBalance
' --------------------------------------------------------------------------
Public Function GetCustomerBalance(custID As String) As Double
    On Error GoTo ErrHandler
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    If wsTrans Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
    Dim total As Double: total = 0
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsTrans.Cells(i, 2).Value) = custID Then
            If CStr(wsTrans.Cells(i, 12).Value) <> "Paid" And CStr(wsTrans.Cells(i, 12).Value) <> "Cancelled" Then
                total = total + CDbl(Val(wsTrans.Cells(i, 11).Value))
            End If
        End If
    Next i
    GetCustomerBalance = total
    Exit Function
ErrHandler:
    ErrorHandler "GetCustomerBalance", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' ListActiveCustomers — Returns collection of "ID - Name" strings
' --------------------------------------------------------------------------
Public Function ListActiveCustomers() As Collection
    On Error GoTo ErrHandler
    Dim col As New Collection
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Customers")
    If ws Is Nothing Then Set ListActiveCustomers = col: Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Set ListActiveCustomers = col: Exit Function
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) <> "" Then
            col.Add CStr(ws.Cells(i, 1).Value) & " - " & CStr(ws.Cells(i, 2).Value)
        End If
    Next i
    Set ListActiveCustomers = col
    Exit Function
ErrHandler:
    ErrorHandler "ListActiveCustomers", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' ValidateCustomerTaxID
' --------------------------------------------------------------------------
Public Function ValidateCustomerTaxID(taxID As String, jurisdiction As String) As Boolean
    ValidateCustomerTaxID = (Len(taxID) > 0)
End Function

' --------------------------------------------------------------------------
' ShowCustomerSelector
' --------------------------------------------------------------------------
Public Sub ShowCustomerSelector()
    On Error GoTo ErrHandler
    g_selectedCustomerID = modForms.ShowCustomerPicker()
    Exit Sub
ErrHandler:
    ErrorHandler "ShowCustomerSelector", Err.Number, Err.Description
End Sub
