Attribute VB_Name = "modModuleImporter"
Option Explicit

' ==============================================================================
' Module: modModuleImporter.bas
' Purpose: Automatically import all .bas module files into the workbook
' Usage: Run ImportAllModules() after building the workbook structure
' ==============================================================================

' --------------------------------------------------------------------------
' MAIN ENTRY POINT: ImportAllModules()
' Imports all .bas files from the project folder
' --------------------------------------------------------------------------
Public Sub ImportAllModules()
    On Error GoTo ErrHandler

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim importCount As Long
    Dim skipCount As Long
    Dim existingModules As String

    ' Get the folder where this workbook is saved
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first, then run this procedure.", vbExclamation
        Exit Sub
    End If

    folderPath = ThisWorkbook.Path
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' List of modules to import (in dependency order)
    Dim modules As Variant
    modules = Array( _
        "modUtilities.bas", _
        "modNumbering.bas", _
        "modCustomer.bas", _
        "modProduct.bas", _
        "modTax.bas", _
        "modInvoice.bas", _
        "modPayment.bas", _
        "modReceipt.bas", _
        "modETR.bas", _
        "modExport.bas", _
        "modDashboard.bas", _
        "modSecurity.bas", _
        "modForms.bas", _
        "modDiagnostics.bas", _
        "modWorkbookBuilder.bas" _
    )

    ' Check existing modules
    existingModules = GetExistingModules()

    ' Confirm with user
    Dim msg As String
    msg = "This will import " & (UBound(modules) + 1) & " VBA modules from:" & vbCrLf & vbCrLf & _
          folderPath & vbCrLf & vbCrLf & _
          "Existing modules will be REPLACED." & vbCrLf & vbCrLf & _
          "Continue?"

    If MsgBox(msg, vbQuestion + vbYesNo, "Import Modules") <> vbYes Then
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Import each module
    Dim i As Long
    For i = 0 To UBound(modules)
        fileName = modules(i)
        fullPath = folderPath & fileName

        ' Check if file exists
        If Dir(fullPath) <> "" Then
            ' Remove existing module if present
            RemoveModule GetModuleNameFromFile(fileName)

            ' Import the module
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import fullPath

            If Err.Number = 0 Then
                importCount = importCount + 1
            Else
                skipCount = skipCount + 1
                Debug.Print "Failed to import: " & fileName & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            skipCount = skipCount + 1
            Debug.Print "File not found: " & fullPath
        End If
    Next i

    Application.ScreenUpdating = True

    ' Show summary
    msg = "Module Import Complete!" & vbCrLf & vbCrLf & _
          "Imported: " & importCount & vbCrLf & _
          "Skipped: " & skipCount & vbCrLf & vbCrLf & _
          "Note: You may need to enable 'Trust access to the VBA project object model'" & vbCrLf & _
          "in Excel Options → Trust Center → Trust Center Settings → Macro Settings."

    MsgBox msg, vbInformation, "Import Complete"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error importing modules: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

' --------------------------------------------------------------------------
' Helper: Remove Module if it exists
' --------------------------------------------------------------------------
Private Sub RemoveModule(moduleName As String)
    On Error Resume Next

    Dim vbc As Object ' VBComponent
    Set vbc = ThisWorkbook.VBProject.VBComponents(moduleName)

    If Not vbc Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbc
    End If

    On Error GoTo 0
End Sub

' --------------------------------------------------------------------------
' Helper: Get module name from filename
' --------------------------------------------------------------------------
Private Function GetModuleNameFromFile(fileName As String) As String
    ' Remove .bas extension
    If Right(LCase(fileName), 4) = ".bas" Then
        GetModuleNameFromFile = Left(fileName, Len(fileName) - 4)
    Else
        GetModuleNameFromFile = fileName
    End If
End Function

' --------------------------------------------------------------------------
' Helper: Get list of existing modules
' --------------------------------------------------------------------------
Private Function GetExistingModules() As String
    On Error Resume Next

    Dim vbc As Object
    Dim moduleList As String

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Type = 1 Then ' vbext_ct_StdModule = 1
            moduleList = moduleList & vbc.Name & ", "
        End If
    Next vbc

    If Len(moduleList) > 0 Then
        moduleList = Left(moduleList, Len(moduleList) - 2)
    End If

    GetExistingModules = moduleList
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' Alternative: Import Single Module
' --------------------------------------------------------------------------
Public Sub ImportSingleModule(Optional moduleName As String = "")
    On Error GoTo ErrHandler

    If moduleName = "" Then
        moduleName = InputBox("Enter the module filename (e.g., modUtilities.bas):", "Import Module")
        If moduleName = "" Then Exit Sub
    End If

    Dim folderPath As String
    folderPath = ThisWorkbook.Path
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Dim fullPath As String
    fullPath = folderPath & moduleName

    If Dir(fullPath) = "" Then
        MsgBox "File not found: " & fullPath, vbExclamation
        Exit Sub
    End If

    ' Remove existing
    RemoveModule GetModuleNameFromFile(moduleName)

    ' Import
    ThisWorkbook.VBProject.VBComponents.Import fullPath

    MsgBox "Module imported: " & moduleName, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' --------------------------------------------------------------------------
' Export All Modules (backup)
' --------------------------------------------------------------------------
Public Sub ExportAllModules()
    On Error GoTo ErrHandler

    Dim folderPath As String
    folderPath = ThisWorkbook.Path
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Create exports subfolder
    Dim exportFolder As String
    exportFolder = folderPath & "ModuleBackup_" & Format(Now, "yyyymmdd_hhnnss") & "\"
    MkDir exportFolder

    Dim vbc As Object
    Dim exportCount As Long

    For Each vbc In ThisWorkbook.VBProject.VBComponents
        If vbc.Type = 1 Then ' Standard module
            vbc.Export exportFolder & vbc.Name & ".bas"
            exportCount = exportCount + 1
        End If
    Next vbc

    MsgBox "Exported " & exportCount & " modules to:" & vbCrLf & exportFolder, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error exporting: " & Err.Description, vbCritical
End Sub

' --------------------------------------------------------------------------
' Check VBA Access
' --------------------------------------------------------------------------
Public Function CheckVBAAccess() As Boolean
    On Error Resume Next

    Dim testCount As Long
    testCount = ThisWorkbook.VBProject.VBComponents.Count

    If Err.Number <> 0 Then
        CheckVBAAccess = False
        MsgBox "VBA Project access is disabled." & vbCrLf & vbCrLf & _
               "To enable:" & vbCrLf & _
               "1. Go to File → Options → Trust Center" & vbCrLf & _
               "2. Click 'Trust Center Settings'" & vbCrLf & _
               "3. Go to 'Macro Settings'" & vbCrLf & _
               "4. Check 'Trust access to the VBA project object model'" & vbCrLf & _
               "5. Click OK and restart Excel", vbExclamation, "VBA Access Required"
        Err.Clear
    Else
        CheckVBAAccess = True
    End If

    On Error GoTo 0
End Function
