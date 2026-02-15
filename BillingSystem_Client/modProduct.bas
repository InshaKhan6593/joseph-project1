Attribute VB_Name = "modProduct"
Option Explicit

' ==============================================================================
' Module: modProduct.bas — Product management (NO PROTECTION, NO TABLES)
'
' Products sheet layout (row 1 = headers, data starts row 2):
'   A=SKU  B=Product/Service Name  C=Description  D=Category
'   E=Unit Price  F=Unit  G=Tax Category  H=Status
' ==============================================================================

Public g_selectedProducts As Collection

' --------------------------------------------------------------------------
' LookupProduct — Search by SKU or Name using direct cell refs
' --------------------------------------------------------------------------
Public Function LookupProduct(identifier As String) As Object
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Products")
    If ws Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    Dim i As Long
    For i = 2 To lastRow
        If LCase(CStr(ws.Cells(i, 1).Value)) = LCase(identifier) Or _
           LCase(CStr(ws.Cells(i, 2).Value)) = LCase(identifier) Then
            Dim dict As Object
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "SKU", CStr(ws.Cells(i, 1).Value)
            dict.Add "Name", CStr(ws.Cells(i, 2).Value)
            dict.Add "Description", CStr(ws.Cells(i, 3).Value)
            dict.Add "Category", CStr(ws.Cells(i, 4).Value)
            dict.Add "UnitPrice", CDbl(Val(ws.Cells(i, 5).Value))
            dict.Add "Unit", CStr(ws.Cells(i, 6).Value)
            dict.Add "TaxCategory", CStr(ws.Cells(i, 7).Value)
            dict.Add "Status", CStr(ws.Cells(i, 8).Value)
            Set LookupProduct = dict
            Exit Function
        End If
    Next i
    Exit Function
ErrHandler:
    ErrorHandler "LookupProduct", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' AddLineItem — Adds a product to Invoice_Template (rows 15-29)
' Only writes to columns A-G. Column H has formulas.
' --------------------------------------------------------------------------
Public Sub AddLineItem(wsInv As Worksheet, lineNum As Long, sku As String, _
                       qty As Double, Optional discPct As Double = 0)
    On Error GoTo ErrHandler
    If lineNum < 1 Or lineNum > 15 Then MsgBox "Max 15 line items", vbExclamation: Exit Sub
    
    Dim prod As Object
    Set prod = LookupProduct(sku)
    If prod Is Nothing Then MsgBox "Product not found: " & sku, vbExclamation: Exit Sub

    modUtilities.UnprotectSheet wsInv.Name

    Dim r As Long
    r = 14 + lineNum  ' Row 15 = line 1, Row 29 = line 15

    wsInv.Cells(r, 1).Value = lineNum           ' A: #
    wsInv.Cells(r, 2).Value = prod("SKU")       ' B: SKU
    wsInv.Cells(r, 3).Value = prod("Name")      ' C: Description
    wsInv.Cells(r, 4).Value = qty               ' D: Qty
    wsInv.Cells(r, 5).Value = prod("UnitPrice") ' E: Unit Price

    ' Apply Default Discount if none provided
    If discPct = 0 Then
        Dim defDisc As Double
        defDisc = Val(modUtilities.GetSetting("Default Discount %"))
        ' Value 0.05 (5%) -> becomes 5 for the column
        If defDisc > 0 Then discPct = defDisc * 100
    End If

    wsInv.Cells(r, 6).Value = discPct           ' F: Discount%
    wsInv.Cells(r, 7).Value = prod("TaxCategory") ' G: Tax Category
    ' H column has formula =IF(D15="","",D15*E15*(1-F15/100)) — DO NOT overwrite!

    ' Set tax amount in H33 (VBA-managed cell)
    ApplyTax wsInv

    modUtilities.ProtectSheet wsInv.Name
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet "Invoice_Template"
    On Error GoTo 0
    ErrorHandler "AddLineItem", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' ApplyTax — Sets H33 (tax) based on subtotal. H31 and H35 are formulas.
' --------------------------------------------------------------------------
Public Sub ApplyTax(wsInv As Worksheet)
    On Error GoTo ErrHandler

    modUtilities.UnprotectSheet wsInv.Name

    Dim subtot As Double
    subtot = Val(wsInv.Range("H31").Value)

    Dim taxRate As Double
    taxRate = modTax.GetTaxRate()
    wsInv.Range("H33").Value = subtot * taxRate
    ' H35 formula (=H31+H33) auto-updates

    modUtilities.ProtectSheet wsInv.Name
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet wsInv.Name
    On Error GoTo 0
    ErrorHandler "ApplyTax", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' ShowProductSelector
' --------------------------------------------------------------------------
Public Sub ShowProductSelector()
    On Error GoTo ErrHandler
    Dim wsInv As Worksheet
    Set wsInv = SafeSheetRef("Invoice_Template")
    modForms.ShowProductPicker wsInv
    Exit Sub
ErrHandler:
    ErrorHandler "ShowProductSelector", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' ListActiveProducts — Returns "SKU - Name (Price)" collection
' --------------------------------------------------------------------------
Public Function ListActiveProducts() As Collection
    On Error GoTo ErrHandler
    Dim col As New Collection
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Products")
    If ws Is Nothing Then Set ListActiveProducts = col: Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Set ListActiveProducts = col: Exit Function
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) <> "" Then
            col.Add CStr(ws.Cells(i, 1).Value) & " - " & CStr(ws.Cells(i, 2).Value) & _
                    " (" & FormatCurrency(CDbl(Val(ws.Cells(i, 5).Value))) & ")"
        End If
    Next i
    Set ListActiveProducts = col
    Exit Function
ErrHandler:
    ErrorHandler "ListActiveProducts", Err.Number, Err.Description
End Function
