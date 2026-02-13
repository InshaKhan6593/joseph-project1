Attribute VB_Name = "modExport"
Option Explicit

' ==============================================================================
' Module: modExport.bas — PDF export (NO PROTECTION, NO NAMED RANGES)
'
' Cell References:
'   Invoice_Template: B8=InvNo, E9=CustName
'   Receipt_Template: B8=RcptNo, B11=Customer
'   ETR_Template:     A7=ReceiptNo
' ==============================================================================

' --------------------------------------------------------------------------
' 1. ExportToPDF(docType, docNumber)
' --------------------------------------------------------------------------
Public Sub ExportToPDF(docType As String, Optional docNumber As String = "")
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim basePath As String
    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim custName As String
    
    basePath = GetSetting("PDF Export Path")
    If basePath = "" Then basePath = ThisWorkbook.Path & "\PDFs"
    
    Select Case LCase(docType)
        Case "invoice"
            Set ws = SafeSheetRef("Invoice_Template")
            If docNumber = "" Then docNumber = CStr(ws.Range("B8").Value)
            custName = CStr(ws.Range("E9").Value)
            
        Case "receipt"
            Set ws = SafeSheetRef("Receipt_Template")
            If docNumber = "" Then docNumber = CStr(ws.Range("B8").Value)
            custName = CStr(ws.Range("B11").Value)
            
        Case "etr"
            Set ws = SafeSheetRef("ETR_Template")
            If docNumber = "" Then
                Dim etrText As String
                etrText = CStr(ws.Range("A7").Value)
                docNumber = Replace(etrText, "Receipt No: ", "")
            End If
            custName = "Cash"
            
        Case Else
            MsgBox "Unknown document type.", vbCritical
            Exit Sub
    End Select
    
    ' Create folder structure: Base\Invoices\2026\02\
    folderPath = CreateFolderStructure(basePath, docType)
    
    ' Generate filename
    fileName = GenerateFileName(docType, docNumber, custName)
    fullPath = folderPath & "\" & fileName
    
    ' Export as PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    AuditLog "PDF_EXPORT", "Exported " & docType & " to " & fullPath
    
    If MsgBox("Export successful!" & vbCrLf & fullPath & vbCrLf & vbCrLf & "Open PDF now?", vbYesNo + vbQuestion) = vbYes Then
        ThisWorkbook.FollowHyperlink fullPath
    End If
    Exit Sub
ErrHandler:
    ErrorHandler "ExportToPDF", Err.Number, Err.Description
End Sub

' --------------------------------------------------------------------------
' 2. CreateFolderStructure — Creates directory hierarchy
' --------------------------------------------------------------------------
Public Function CreateFolderStructure(basePath As String, docType As String) As String
    On Error Resume Next
    
    Dim typeFolder As String
    Dim yearFolder As String
    Dim monthFolder As String
    
    typeFolder = basePath & "\" & StrConv(docType, vbProperCase) & "s"
    yearFolder = typeFolder & "\" & Format(Date, "yyyy")
    monthFolder = yearFolder & "\" & Format(Date, "mm")
    
    ' Create each level if not exists
    If Dir(basePath, vbDirectory) = "" Then MkDir basePath
    If Dir(typeFolder, vbDirectory) = "" Then MkDir typeFolder
    If Dir(yearFolder, vbDirectory) = "" Then MkDir yearFolder
    If Dir(monthFolder, vbDirectory) = "" Then MkDir monthFolder
    
    CreateFolderStructure = monthFolder
    On Error GoTo 0
End Function

' --------------------------------------------------------------------------
' 3. GenerateFileName — Clean filename for PDF
' --------------------------------------------------------------------------
Public Function GenerateFileName(docType As String, docNumber As String, customerName As String) As String
    Dim cleanName As String
    cleanName = CleanString(customerName)
    GenerateFileName = docNumber & "_" & cleanName & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
End Function

Private Function CleanString(inputStr As String) As String
    Dim badChars As String
    Dim i As Integer
    badChars = "/\:*?""<>|"
    CleanString = inputStr
    For i = 1 To Len(badChars)
        CleanString = Replace(CleanString, Mid(badChars, i, 1), "")
    Next i
    CleanString = Trim(CleanString)
End Function

' --------------------------------------------------------------------------
' 4. BatchExport — Placeholder
' --------------------------------------------------------------------------
Public Sub BatchExport(docType As String, fromDate As Date, toDate As Date)
    MsgBox "Batch export not yet implemented.", vbInformation
End Sub

' --------------------------------------------------------------------------
' 5. ExportToEmail — Uses Outlook to email PDF
' --------------------------------------------------------------------------
Public Sub ExportToEmail(filePath As String, recipientEmail As String)
    On Error GoTo ErrHandler
    
    Dim OutApp As Object
    Dim OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    With OutMail
        .To = recipientEmail
        .Subject = "Document Attachment"
        .Body = "Please find attached document." & vbCrLf & vbCrLf & "Sent from Billing System"
        .Attachments.Add filePath
        .Display
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    Exit Sub
ErrHandler:
    MsgBox "Could not create email. Ensure Outlook is installed.", vbExclamation
End Sub
