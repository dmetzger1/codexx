VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "frmPXManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPXBrowse_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select PDF File"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf"
    If fd.Show Then
        Me.txtPXPath.Text = fd.SelectedItems(1)
    End If
End Sub

Private Sub cmdPXExecute_Click()
    Dim action As String
    action = Me.cboPXAction.Value
    Select Case action
        Case "Export Sheet to PDF"
            PX_ExportActiveSheet Me.txtPXPath.Text
        Case "Merge PDFs"
            PX_MergePDFs Me.txtPXPath.Text
        Case "Split PDF"
            PX_SplitPDF Me.txtPXPath.Text
        Case "Delete Pages"
            PX_DeletePDFPages Me.txtPXPath.Text
    End Select
End Sub

Private Sub cmdPXClose_Click()
    Unload Me
End Sub
