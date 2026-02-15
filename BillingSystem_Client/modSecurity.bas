Attribute VB_Name = "modSecurity"
Option Explicit

' ==============================================================================
' Module: modSecurity.bas
' Purpose: Workbook setup (PROTECTION REMOVED for development)
' Created: 2026-02-13
' Dependencies: modUtilities, modDashboard
' ==============================================================================

' --------------------------------------------------------------------------
' 1. SetupWorkbook() — MASTER SETUP (Run after importing all modules)
' --------------------------------------------------------------------------
Public Sub SetupWorkbook()
    On Error GoTo ErrHandler
    
    modUtilities.TogglePerformance True
    
    ' Step 1: Protect ALL sheets
    ProtectAllSheets
    
    ' Step 2: Create AuditLog if missing
    modUtilities.AuditLog "SETUP", "Workbook initialized and protected"
    
    ' Step 3: Refresh Dashboard
    modDashboard.RefreshDashboard
    
    ' Step 4: Navigate to Dashboard
    modDashboard.NavigateTo "Dashboard"
    
    modUtilities.TogglePerformance False
    MsgBox "Setup complete! All sheets are now protected with 'admin2026'.", vbInformation
    Exit Sub
ErrHandler:
    modUtilities.TogglePerformance False
    modUtilities.ErrorHandler "SetupWorkbook", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 2. ProtectAllSheets() — Protects EVERY sheet with standard settings
' --------------------------------------------------------------------------
Public Sub ProtectAllSheets()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        modUtilities.ProtectSheet ws.Name
    Next ws
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' 3. RemoveAllProtection() — Unprotects EVERY sheet
' --------------------------------------------------------------------------
Public Sub RemoveAllProtection()
    On Error Resume Next
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        modUtilities.UnprotectSheet ws.Name
    Next ws
    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' 3. ValidateUser() — Simple password check for admin functions
' --------------------------------------------------------------------------
Public Function ValidateUser(Optional requiredRole As String = "") As Boolean
    ValidateUser = True ' Skip validation during development
End Function

' --------------------------------------------------------------------------
' 4. ProtectVBAProject() — Instructions only
' --------------------------------------------------------------------------
Public Sub ProtectVBAProject()
    MsgBox "To protect VBA Code:" & vbCrLf & _
           "1. Tools -> VBAProject Properties" & vbCrLf & _
           "2. Protection Tab" & vbCrLf & _
           "3. Check 'Lock project for viewing'" & vbCrLf & _
           "4. Set password", vbInformation
End Sub
