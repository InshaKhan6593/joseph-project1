Attribute VB_Name = "modFormBuilder"
Option Explicit

' ==============================================================================
' Module: modFormBuilder.bas
' Purpose: Dynamically builds a Searchable UserForm for selecting items
'          Avoids the need for binary .frx files by building controls at runtime
' ==============================================================================

' --------------------------------------------------------------------------
' EnsureSelectionFormExists - Checks if frmSelection exists, builds if not
' --------------------------------------------------------------------------
Public Sub EnsureSelectionFormExists()
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check if it exists
    Dim comp As Object
    Set comp = vbProj.VBComponents("frmSelection")
    
    If comp Is Nothing Then
        ' Create it
        Set comp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm
        comp.Name = "frmSelection"
        comp.Properties("Caption") = "Select Item"
        comp.Properties("Width") = 240
        comp.Properties("Height") = 300
        
        ' Inject Code
        Dim code As String
        code = GetFormCode()
        comp.CodeModule.AddFromString code
    End If
End Sub

' --------------------------------------------------------------------------
' GetFormCode - Returns the VBA code for the UserForm
' Notes: Uses string concatenation to avoid recursion/line-limit issues
' --------------------------------------------------------------------------
Private Function GetFormCode() As String
    Dim s As String
    s = "Option Explicit" & vbCrLf
    s = s & "Public SelectedValue As String" & vbCrLf
    s = s & "Public IsCancelled As Boolean" & vbCrLf
    s = s & "Private AllItems As Collection" & vbCrLf & vbCrLf
    
    s = s & "Private WithEvents lstItems As MSForms.ListBox" & vbCrLf
    s = s & "Private WithEvents txtSearch As MSForms.TextBox" & vbCrLf
    s = s & "Private WithEvents btnOK As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents btnCancel As MSForms.CommandButton" & vbCrLf
    s = s & "Private WithEvents lblSearch As MSForms.Label" & vbCrLf & vbCrLf
    
    s = s & "Private Sub UserForm_Initialize()" & vbCrLf
    s = s & "    Me.Caption = ""Select Item""" & vbCrLf
    s = s & "    Me.Width = 300" & vbCrLf
    s = s & "    Me.Height = 350" & vbCrLf
    s = s & "    SelectedValue = """"" & vbCrLf
    s = s & "    IsCancelled = True" & vbCrLf
    s = s & "    Set AllItems = New Collection" & vbCrLf
    s = s & "    BuildControls" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub BuildControls()" & vbCrLf
    s = s & "    Set lblSearch = Me.Controls.Add(""Forms.Label.1"", ""lblSearch"")" & vbCrLf
    s = s & "    With lblSearch" & vbCrLf
    s = s & "        .Top = 10: .Left = 10: .Width = 260: .Height = 15" & vbCrLf
    s = s & "        .Caption = ""Search (Type to filter):""" & vbCrLf
    s = s & "    End With" & vbCrLf
    
    s = s & "    Set txtSearch = Me.Controls.Add(""Forms.TextBox.1"", ""txtSearch"")" & vbCrLf
    s = s & "    With txtSearch" & vbCrLf
    s = s & "        .Top = 30: .Left = 10: .Width = 260: .Height = 20" & vbCrLf
    s = s & "    End With" & vbCrLf
    
    s = s & "    Set lstItems = Me.Controls.Add(""Forms.ListBox.1"", ""lstItems"")" & vbCrLf
    s = s & "    With lstItems" & vbCrLf
    s = s & "        .Top = 60: .Left = 10: .Width = 260: .Height = 200" & vbCrLf
    s = s & "    End With" & vbCrLf
    
    s = s & "    Set btnOK = Me.Controls.Add(""Forms.CommandButton.1"", ""btnOK"")" & vbCrLf
    s = s & "    With btnOK" & vbCrLf
    s = s & "        .Top = 270: .Left = 150: .Width = 60: .Height = 25" & vbCrLf
    s = s & "        .Caption = ""OK"": .Default = True" & vbCrLf
    s = s & "    End With" & vbCrLf
    
    s = s & "    Set btnCancel = Me.Controls.Add(""Forms.CommandButton.1"", ""btnCancel"")" & vbCrLf
    s = s & "    With btnCancel" & vbCrLf
    s = s & "        .Top = 270: .Left = 215: .Width = 60: .Height = 25" & vbCrLf
    s = s & "        .Caption = ""Cancel"": .Cancel = True" & vbCrLf
    s = s & "    End With" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Public Sub LoadItems(items As Collection)" & vbCrLf
    s = s & "    Set AllItems = items" & vbCrLf
    s = s & "    FilterList """"" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub txtSearch_Change()" & vbCrLf
    s = s & "    FilterList txtSearch.Text" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub FilterList(criteria As String)" & vbCrLf
    s = s & "    lstItems.Clear" & vbCrLf
    s = s & "    Dim itm As Variant" & vbCrLf
    s = s & "    For Each itm In AllItems" & vbCrLf
    s = s & "        If criteria = """" Or InStr(1, itm, criteria, vbTextCompare) > 0 Then" & vbCrLf
    s = s & "            lstItems.AddItem itm" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next itm" & vbCrLf
    s = s & "    If lstItems.ListCount > 0 Then lstItems.ListIndex = 0" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub btnOK_Click()" & vbCrLf
    s = s & "    If lstItems.ListIndex >= 0 Then" & vbCrLf
    s = s & "        SelectedValue = lstItems.List(lstItems.ListIndex)" & vbCrLf
    s = s & "        IsCancelled = False" & vbCrLf
    s = s & "        Me.Hide" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub btnCancel_Click()" & vbCrLf
    s = s & "    IsCancelled = True" & vbCrLf
    s = s & "    Me.Hide" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    
    s = s & "Private Sub lstItems_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    s = s & "    btnOK_Click" & vbCrLf
    s = s & "End Sub"
    
    GetFormCode = s
End Function

' --------------------------------------------------------------------------
' ShowSelectionDialog - Main Entry Point
' Returns selected string OR empty string if cancelled
' --------------------------------------------------------------------------
Public Function ShowSelectionDialog(title As String, items As Collection) As String
    ' 1. Ensure form exists
    EnsureSelectionFormExists
    
    ' 2. Load and Show
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmSelection")
    
    frm.Caption = title
    frm.LoadItems items
    
    frm.Show
    
    If Not frm.IsCancelled Then
        ShowSelectionDialog = frm.SelectedValue
    Else
        ShowSelectionDialog = ""
    End If
    
    Unload frm
End Function
