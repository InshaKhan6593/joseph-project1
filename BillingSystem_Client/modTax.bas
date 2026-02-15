Attribute VB_Name = "modTax"
Option Explicit

' ==============================================================================
' Module: modTax.bas — Tax calculations (NO PROTECTION)
' Tax table: Settings A16:E24
' ==============================================================================

' --------------------------------------------------------------------------
' GetTaxRate() — Returns the standard tax rate for current jurisdiction
' --------------------------------------------------------------------------
Public Function GetTaxRate() As Double
    On Error GoTo ErrHandler
    
    ' Get Jurisdiction from Named Range
    Dim jurisdiction As String
    On Error Resume Next
    jurisdiction = LCase(CStr(ThisWorkbook.Names("rngJurisdiction").RefersToRange.Value))
    On Error GoTo ErrHandler
    
    ' Get Tax Table from Named Range
    Dim rngTable As Range
    On Error Resume Next
    Set rngTable = ThisWorkbook.Names("rngTaxTable").RefersToRange
    On Error GoTo ErrHandler
    
    ' Fallback if ranges missing (e.g. before build)
    If rngTable Is Nothing Then
        GetTaxRate = 0.16
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To rngTable.Rows.Count
        If LCase(CStr(rngTable.Cells(i, 1).Value)) = jurisdiction Then
            If InStr(1, LCase(CStr(rngTable.Cells(i, 2).Value)), "standard") > 0 Then
                GetTaxRate = CDbl(Val(Replace(CStr(rngTable.Cells(i, 3).Value), "%", ""))) / 100
                Exit Function
            End If
        End If
    Next i
    GetTaxRate = 0.16 ' Default Kenya VAT
    Exit Function
ErrHandler:
    GetTaxRate = 0.16
End Function

' --------------------------------------------------------------------------
' CalculateInvoiceTax — Calculates tax and writes to H33
' --------------------------------------------------------------------------
Public Function CalculateInvoiceTax() As Double
    On Error GoTo ErrHandler
    Dim wsInv As Worksheet
    Set wsInv = SafeSheetRef("Invoice_Template")
    
    modUtilities.UnprotectSheet wsInv.Name
    
    Dim subtot As Double
    subtot = CDbl(Val(wsInv.Range("H31").Value))
    Dim taxRate As Double
    taxRate = GetTaxRate()
    Dim totalTax As Double
    totalTax = subtot * taxRate
    wsInv.Range("H33").Value = totalTax
    ' Do NOT overwrite H35 — it has a formula =H31-H32+H33
    
    modUtilities.ProtectSheet wsInv.Name
    
    CalculateInvoiceTax = totalTax
    Exit Function
ErrHandler:
    modUtilities.ProtectSheet wsInv.Name
    ErrorHandler "CalculateInvoiceTax", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' GenerateTaxSummary — Populates TaxSummary sheet
' --------------------------------------------------------------------------
Public Sub GenerateTaxSummary()
    On Error GoTo ErrHandler
    Dim wsSum As Worksheet
    Set wsSum = SafeSheetRef("TaxSummary")
    Dim wsTrans As Worksheet
    Set wsTrans = SafeSheetRef("Transactions")
    If wsSum Is Nothing Or wsTrans Is Nothing Then Exit Sub

    modUtilities.UnprotectSheet "TaxSummary"

    ' Clear data area (rows 7-18)
    On Error Resume Next
    wsSum.Range("A7:G18").ClearContents
    On Error GoTo ErrHandler

    Dim lastRow As Long
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        modUtilities.ProtectSheet "TaxSummary"
        Exit Sub
    End If

    Dim sumRow As Long: sumRow = 7
    Dim totalRev As Double, totalTax As Double, count As Long
    Dim i As Long
    For i = 2 To lastRow
        If CStr(wsTrans.Cells(i, 12).Value) <> "Cancelled" Then
            totalRev = totalRev + CDbl(Val(wsTrans.Cells(i, 9).Value))
            totalTax = totalTax + CDbl(Val(wsTrans.Cells(i, 7).Value))
            count = count + 1
        End If
    Next i

    wsSum.Cells(sumRow, 1).Value = Format(Date, "mmmm yyyy")
    wsSum.Cells(sumRow, 2).Value = modUtilities.GetSetting("Jurisdiction")
    wsSum.Cells(sumRow, 3).Value = totalRev
    wsSum.Cells(sumRow, 4).Value = totalTax
    wsSum.Cells(sumRow, 5).Value = Format(GetTaxRate(), "0%")
    wsSum.Cells(sumRow, 6).Value = count

    Dim totalOut As Double
    For i = 2 To lastRow
        If CStr(wsTrans.Cells(i, 12).Value) <> "Cancelled" And CStr(wsTrans.Cells(i, 12).Value) <> "Paid" Then
            totalOut = totalOut + CDbl(Val(wsTrans.Cells(i, 11).Value))
        End If
    Next i
    wsSum.Cells(sumRow, 7).Value = totalOut

    modUtilities.ProtectSheet "TaxSummary"

    AuditLog "TAX_SUMMARY", "Generated for " & Format(Date, "mmmm yyyy")
    MsgBox "Tax summary updated!", vbInformation
    Exit Sub
ErrHandler:
    On Error Resume Next
    modUtilities.ProtectSheet "TaxSummary"
    On Error GoTo 0
    ErrorHandler "GenerateTaxSummary", Err.Number, Err.Description
End Sub
