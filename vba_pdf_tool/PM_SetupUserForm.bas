Attribute VB_Name = "PM_SetupUserForm"
Option Explicit

'Automates creation of the PDF Manager userform
Sub PM_CreatePDFManagerUserForm()
    Dim comp As VBIDE.VBComponent
    Dim frm As Object
    Dim ctrl As MSForms.Control

    ' Add userform
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    comp.Properties("Caption") = "PDF Manager"
    comp.Properties("Width") = 400
    comp.Properties("Height") = 300
    comp.Name = "frmPM_PDFManager"
    Set frm = comp.Designer

    ' Label: PDF path
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblPM_Path"
        .Caption = "PDF File:"
        .Top = 20
        .Left = 10
        .Width = 70
    End With

    ' Textbox: path
    Set ctrl = frm.Controls.Add("Forms.TextBox.1")
    With ctrl
        .Name = "txtPM_PDFPath"
        .Top = 20
        .Left = 80
        .Width = 200
    End With

    ' Button: browse
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnPM_Browse"
        .Caption = "Browse..."
        .Top = 20
        .Left = 290
        .Width = 80
    End With

    ' Label: action
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblPM_Action"
        .Caption = "Action:"
        .Top = 60
        .Left = 10
        .Width = 70
    End With

    ' ComboBox: action list
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cmbPM_Action"
        .Top = 60
        .Left = 80
        .Width = 290
        .AddItem "Export Sheet to PDF"
        .AddItem "Merge PDFs"
        .AddItem "Split PDF"
        .AddItem "Delete Pages"
    End With

    ' Run button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnPM_Run"
        .Caption = "Run"
        .Top = 100
        .Left = 80
        .Width = 100
    End With

    ' Close button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnPM_Close"
        .Caption = "Close"
        .Top = 100
        .Left = 190
        .Width = 100
    End With
End Sub
    ' Add userform code
    With comp.CodeModule
        .InsertLines .CountOfLines + 1, _
            "Private Sub btnPM_Browse_Click()" & vbCrLf & _
            "    Dim fd As Office.FileDialog" & vbCrLf & _
            "    Set fd = Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf & _
            "    fd.Title = \"Select PDF File\"" & vbCrLf & _
            "    fd.Filters.Clear" & vbCrLf & _
            "    fd.Filters.Add \"PDF Files\", \"*.pdf\"" & vbCrLf & _
            "    If fd.Show Then" & vbCrLf & _
            "        Me.txtPM_PDFPath.Text = fd.SelectedItems(1)" & vbCrLf & _
            "    End If" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "" & vbCrLf & _
            "Private Sub btnPM_Run_Click()" & vbCrLf & _
            "    Dim action As String" & vbCrLf & _
            "    action = Me.cmbPM_Action.Value" & vbCrLf & _
            "    Select Case action" & vbCrLf & _
            "        Case \"Export Sheet to PDF\"" & vbCrLf & _
            "            PM_ExportActiveSheet Me.txtPM_PDFPath.Text" & vbCrLf & _
            "        Case \"Merge PDFs\"" & vbCrLf & _
            "            PM_MergeSelectedPDFs Me.txtPM_PDFPath.Text" & vbCrLf & _
            "        Case \"Split PDF\"" & vbCrLf & _
            "            PM_SplitSelectedPDF Me.txtPM_PDFPath.Text" & vbCrLf & _
            "        Case \"Delete Pages\"" & vbCrLf & _
            "            PM_DeletePagesFromSelectedPDF Me.txtPM_PDFPath.Text" & vbCrLf & _
            "    End Select" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "" & vbCrLf & _
            "Private Sub btnPM_Close_Click()" & vbCrLf & _
            "    Unload Me" & vbCrLf & _
            "End Sub"
    End With
