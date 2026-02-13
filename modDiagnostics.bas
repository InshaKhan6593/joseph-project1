Attribute VB_Name = "modDiagnostics"
Option Explicit

Public Sub SystemDiagnostics()
    On Error Resume Next
    Dim out As String
    out = "--- SYSTEM DIAGNOSTICS ---" & vbCrLf
    
    ' 1. Check Sheets
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    If ws Is Nothing Then
        out = out & "[FAIL] Dashboard sheet not found!" & vbCrLf
    Else
        out = out & "[OK] Dashboard sheet found." & vbCrLf
    End If
    
    ' 2. Check Shapes on Dashboard
    Dim shp As Shape
    Dim foundNewInv As Boolean
    out = out & "--- BUTTONS FOUND ---" & vbCrLf
    For Each shp In ws.Shapes
        out = out & "Shape: '" & shp.Name & "' | Macro: " & shp.OnAction & vbCrLf
        If shp.Name = "btnNewInvoice" Then foundNewInv = True
    Next shp
    
    If Not foundNewInv Then 
        out = out & "[CRITICAL] Shape named 'btnNewInvoice' was NOT found!" & vbCrLf
    End If
    
    MsgBox out, vbInformation, "Diagnostic Results"
End Sub
