Attribute VB_Name = "PXFormBuilder"
Option Explicit

'Automates creation of the PDFX Manager userform and wires up events
Public Sub PX_BuildManagerForm()
    Dim comp As VBIDE.VBComponent
    Dim frm As Object
    Dim ctrl As MSForms.Control

    ' Add userform
    Set comp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    comp.Properties("Caption") = "PDFX Manager"
    comp.Properties("Width") = 400
    comp.Properties("Height") = 300
    comp.Name = "frmPXManager"
    Set frm = comp.Designer

    ' Label: PDF path
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblPXPath"
        .Caption = "PDF File:"
        .Top = 20
        .Left = 10
        .Width = 70
    End With

    ' Textbox: path
    Set ctrl = frm.Controls.Add("Forms.TextBox.1")
    With ctrl
        .Name = "txtPXPath"
        .Top = 20
        .Left = 80
        .Width = 200
    End With

    ' Button: browse
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdPXBrowse"
        .Caption = "Browse..."
        .Top = 20
        .Left = 290
        .Width = 80
    End With

    ' Label: action
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblPXAction"
        .Caption = "Action:"
        .Top = 60
        .Left = 10
        .Width = 70
    End With

    ' ComboBox: action list
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboPXAction"
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
        .Name = "cmdPXExecute"
        .Caption = "Run"
        .Top = 100
        .Left = 80
        .Width = 100
    End With

    ' Close button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdPXClose"
        .Caption = "Close"
        .Top = 100
        .Left = 190
        .Width = 100
    End With

    ' Add userform code
    With comp.CodeModule
        .InsertLines .CountOfLines + 1, _
            "Private Sub cmdPXBrowse_Click()" & vbCrLf & _
            "    Dim fd As Office.FileDialog" & vbCrLf & _
            "    Set fd = Application.FileDialog(msoFileDialogFilePicker)" & vbCrLf & _
            "    fd.Title = \"Select PDF File\"" & vbCrLf & _
            "    fd.Filters.Clear" & vbCrLf & _
            "    fd.Filters.Add \"PDF Files\", \"*.pdf\"" & vbCrLf & _
            "    If fd.Show Then" & vbCrLf & _
            "        Me.txtPXPath.Text = fd.SelectedItems(1)" & vbCrLf & _
            "    End If" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "" & vbCrLf & _
            "Private Sub cmdPXExecute_Click()" & vbCrLf & _
            "    Dim action As String" & vbCrLf & _
            "    action = Me.cboPXAction.Value" & vbCrLf & _
            "    Select Case action" & vbCrLf & _
            "        Case \"Export Sheet to PDF\"" & vbCrLf & _
            "            PX_ExportActiveSheet Me.txtPXPath.Text" & vbCrLf & _
            "        Case \"Merge PDFs\"" & vbCrLf & _
            "            PX_MergePDFs Me.txtPXPath.Text" & vbCrLf & _
            "        Case \"Split PDF\"" & vbCrLf & _
            "            PX_SplitPDF Me.txtPXPath.Text" & vbCrLf & _
            "        Case \"Delete Pages\"" & vbCrLf & _
            "            PX_DeletePDFPages Me.txtPXPath.Text" & vbCrLf & _
            "    End Select" & vbCrLf & _
            "End Sub" & vbCrLf & _
            "" & vbCrLf & _
            "Private Sub cmdPXClose_Click()" & vbCrLf & _
            "    Unload Me" & vbCrLf & _
            "End Sub"
    End With
End Sub
