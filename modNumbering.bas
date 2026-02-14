Attribute VB_Name = "modNumbering"
Option Explicit

' ==============================================================================
' Module: modNumbering.bas — Auto-numbering (NO PROTECTION)
' ==============================================================================

' --------------------------------------------------------------------------
' GetNextInvoiceNumber() As String — Format: INV-2026-0001
' --------------------------------------------------------------------------
Public Function GetNextInvoiceNumber() As String
    On Error GoTo ErrHandler
    
    Dim rngYear As Range
    Dim rngCount As Range
    
    ' Use Named Ranges
    Set rngYear = ThisWorkbook.Names("rngYearPrefix").RefersToRange
    Set rngCount = ThisWorkbook.Names("rngLastInvoice").RefersToRange
    
    If rngYear Is Nothing Or rngCount Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Set ws = rngYear.Parent ' Get the sheet these ranges belong to
    modUtilities.UnprotectSheet ws.Name
    
    ' Year rollover check
    Dim currentYear As Long
    currentYear = Year(Date)
    If CLng(Val(rngYear.Value)) <> currentYear Then
        rngCount.Value = 0
        ThisWorkbook.Names("rngLastReceipt").RefersToRange.Value = 0
        ThisWorkbook.Names("rngLastETR").RefersToRange.Value = 0
        rngYear.Value = currentYear
    End If
    
    ' Increment counter
    Dim nextCount As Long
    nextCount = CLng(Val(rngCount.Value)) + 1
    rngCount.Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextInvoiceNumber = "INV-" & currentYear & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    If Not ws Is Nothing Then modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextInvoiceNumber", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' GetNextReceiptNumber() As String — Format: RCPT-2026-0001
' --------------------------------------------------------------------------
Public Function GetNextReceiptNumber() As String
    On Error GoTo ErrHandler
    
    Dim rngCount As Range
    Set rngCount = ThisWorkbook.Names("rngLastReceipt").RefersToRange
    If rngCount Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Set ws = rngCount.Parent
    modUtilities.UnprotectSheet ws.Name
    
    Dim nextCount As Long
    nextCount = CLng(Val(rngCount.Value)) + 1
    rngCount.Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextReceiptNumber = "RCPT-" & Year(Date) & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    If Not ws Is Nothing Then modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextReceiptNumber", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' GetNextETRNumber() As String — Format: ETR-2026-0001
' --------------------------------------------------------------------------
Public Function GetNextETRNumber() As String
    On Error GoTo ErrHandler
    
    Dim rngCount As Range
    Set rngCount = ThisWorkbook.Names("rngLastETR").RefersToRange
    If rngCount Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Set ws = rngCount.Parent
    modUtilities.UnprotectSheet ws.Name
    
    Dim nextCount As Long
    nextCount = CLng(Val(rngCount.Value)) + 1
    rngCount.Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextETRNumber = "ETR-" & Year(Date) & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    If Not ws Is Nothing Then modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextETRNumber", Err.Number, Err.Description
End Function
