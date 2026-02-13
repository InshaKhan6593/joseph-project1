Attribute VB_Name = "modUtilities"
Option Explicit

' ==============================================================================
' Module: modUtilities.bas
' Purpose: Core utility functions used by all other modules
' NO PROTECTION — all protection removed for reliability
' ==============================================================================

' --------------------------------------------------------------------------
' 1. FormatCurrency(amount As Double) As String
' --------------------------------------------------------------------------
Public Function FormatCurrency(amount As Double) As String
    On Error GoTo ErrHandler
    
    Dim sym As String
    sym = GetSetting("Currency Symbol")
    If sym = "" Then sym = "KES"
    
    Dim rounded As Double
    rounded = WorksheetFunction.Round(amount, 2)
    
    Select Case UCase(sym)
        Case "USD", "$":  FormatCurrency = "$" & Format(rounded, "#,##0.00")
        Case "GBP", "£":  FormatCurrency = "£" & Format(rounded, "#,##0.00")
        Case Else:        FormatCurrency = sym & " " & Format(rounded, "#,##0.00")
    End Select
    Exit Function
ErrHandler:
    FormatCurrency = Format(amount, "#,##0.00")
End Function

' --------------------------------------------------------------------------
' 2. FormatDate(dt As Date, Optional style As String = "standard") As String
' --------------------------------------------------------------------------
Public Function FormatDate(dt As Date, Optional style As String = "standard") As String
    On Error GoTo ErrHandler
    Select Case LCase(style)
        Case "etr":      FormatDate = Format(dt, "dd/mm/yyyy HH:MM")
        Case "file":     FormatDate = Format(dt, "yyyy-mm-dd")
        Case Else:       FormatDate = Format(dt, "dd-mmm-yyyy")
    End Select
    Exit Function
ErrHandler:
    FormatDate = CStr(dt)
End Function

' --------------------------------------------------------------------------
' 3. GetSetting(settingName As String) As String
'    Searches Settings column A for label, returns column B value
' --------------------------------------------------------------------------
Public Function GetSetting(settingName As String) As String
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("Settings")
    If ws Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = LCase(Trim(settingName)) Then
            GetSetting = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    
    GetSetting = ""
    Exit Function
ErrHandler:
    GetSetting = ""
End Function

' --------------------------------------------------------------------------
' 4. ValidateInput(value, dataType, Optional minVal, Optional maxVal) As Boolean
' --------------------------------------------------------------------------
Public Function ValidateInput(value As Variant, dataType As String, _
    Optional minVal As Variant, Optional maxVal As Variant) As Boolean
    On Error GoTo ErrHandler
    ValidateInput = False
    Select Case LCase(dataType)
        Case "number"
            If Not IsNumeric(value) Then Exit Function
            If CDbl(value) < 0 Then Exit Function
            If Not IsMissing(minVal) Then If CDbl(value) < CDbl(minVal) Then Exit Function
            If Not IsMissing(maxVal) Then If CDbl(value) > CDbl(maxVal) Then Exit Function
            ValidateInput = True
        Case "text":  ValidateInput = (Len(Trim(CStr(value))) > 0)
        Case "date":  ValidateInput = IsDate(value)
        Case "email"
            Dim s As String: s = CStr(value)
            ValidateInput = (InStr(s, "@") > 0 And InStr(s, ".") > 0)
    End Select
    Exit Function
ErrHandler:
    ValidateInput = False
End Function

' --------------------------------------------------------------------------
' 5. ErrorHandler(procName, errNum, errDesc)
' --------------------------------------------------------------------------
Public Sub ErrorHandler(procName As String, errNum As Long, errDesc As String)
    AuditLog "ERROR", "Proc: " & procName & " | #" & errNum & " | " & errDesc
    MsgBox "Error in " & procName & ": " & errDesc, vbCritical, "System Error"
End Sub

' --------------------------------------------------------------------------
' 6. TogglePerformance(turnOn As Boolean)
' --------------------------------------------------------------------------
Public Sub TogglePerformance(turnOn As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = Not turnOn
    If turnOn Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
    Application.EnableEvents = Not turnOn
End Sub

' --------------------------------------------------------------------------
' 7. GetNextRow(ws As Worksheet, col As Long) As Long
' --------------------------------------------------------------------------
Public Function GetNextRow(ws As Worksheet, col As Long) As Long
    On Error GoTo ErrHandler
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    GetNextRow = lastRow + 1
    Exit Function
ErrHandler:
    GetNextRow = 2
End Function

' --------------------------------------------------------------------------
' 8. SafeSheetRef(sheetName As String) As Worksheet
' --------------------------------------------------------------------------
Public Function SafeSheetRef(sheetName As String) As Worksheet
    On Error Resume Next
    Set SafeSheetRef = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' 9. SafeRange(ws As Worksheet, rangeName As String) As Range
' --------------------------------------------------------------------------
Public Function SafeRange(ws As Worksheet, rangeName As String) As Range
    On Error Resume Next
    Set SafeRange = ws.Range(rangeName)
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' 10. ClearSafe(ws As Worksheet, rangeName As String)
' --------------------------------------------------------------------------
Public Sub ClearSafe(ws As Worksheet, rangeName As String)
    On Error Resume Next
    UnprotectSheet ws.Name
    ws.Range(rangeName).ClearContents
    ProtectSheet ws.Name
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' 11. UnprotectSheet(sheetName)
' --------------------------------------------------------------------------
Public Sub UnprotectSheet(sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SafeSheetRef(sheetName)
    If Not ws Is Nothing Then
        ws.Unprotect "admin2026"
    End If
End Sub

' --------------------------------------------------------------------------
' 12. ProtectSheet(sheetName)
' --------------------------------------------------------------------------
Public Sub ProtectSheet(sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SafeSheetRef(sheetName)
    If Not ws Is Nothing Then
        ' Protect with UserInterfaceOnly to allow VBA to work without constant unprotecting
        ws.Protect Password:="admin2026", UserInterfaceOnly:=True, _
                   AllowSorting:=True, AllowFiltering:=True
    End If
End Sub

' --------------------------------------------------------------------------
' 11. AuditLog(action As String, details As String)
' --------------------------------------------------------------------------
Public Sub AuditLog(action As String, details As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = SafeSheetRef("AuditLog")
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "AuditLog"
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Action"
        ws.Cells(1, 3).Value = "Details"
        ws.Cells(1, 4).Value = "User"
        ws.Rows(1).Font.Bold = True
    End If
    
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = action
    ws.Cells(nextRow, 3).Value = details
    ws.Cells(nextRow, 4).Value = Application.UserName
    On Error GoTo 0
End Sub
