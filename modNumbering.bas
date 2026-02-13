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
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Settings")
    If ws Is Nothing Then Exit Function
    
    modUtilities.UnprotectSheet ws.Name
    
    ' Year rollover check (B29 = Year Prefix)
    Dim currentYear As Long
    currentYear = Year(Date)
    If CLng(Val(ws.Range("B29").Value)) <> currentYear Then
        ws.Range("B26").Value = 0
        ws.Range("B27").Value = 0
        ws.Range("B28").Value = 0
        ws.Range("B29").Value = currentYear
    End If
    
    ' Increment counter (B26 = Last Invoice Number)
    Dim nextCount As Long
    nextCount = CLng(Val(ws.Range("B26").Value)) + 1
    ws.Range("B26").Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextInvoiceNumber = "INV-" & currentYear & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextInvoiceNumber", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' GetNextReceiptNumber() As String — Format: RCPT-2026-0001
' --------------------------------------------------------------------------
Public Function GetNextReceiptNumber() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Settings")
    
    modUtilities.UnprotectSheet ws.Name
    
    Dim nextCount As Long
    nextCount = CLng(Val(ws.Range("B27").Value)) + 1
    ws.Range("B27").Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextReceiptNumber = "RCPT-" & Year(Date) & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextReceiptNumber", Err.Number, Err.Description
End Function

' --------------------------------------------------------------------------
' GetNextETRNumber() As String — Format: ETR-2026-0001
' --------------------------------------------------------------------------
Public Function GetNextETRNumber() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Settings")
    
    modUtilities.UnprotectSheet ws.Name
    
    Dim nextCount As Long
    nextCount = CLng(Val(ws.Range("B28").Value)) + 1
    ws.Range("B28").Value = nextCount
    
    modUtilities.ProtectSheet ws.Name
    
    GetNextETRNumber = "ETR-" & Year(Date) & "-" & Format(nextCount, "0000")
    Exit Function
ErrHandler:
    modUtilities.ProtectSheet ws.Name
    ErrorHandler "GetNextETRNumber", Err.Number, Err.Description
End Function
