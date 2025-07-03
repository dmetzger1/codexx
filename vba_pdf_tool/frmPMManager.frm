VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "frmPMManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub pm_btnBrowse_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select PDF File"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf"
    If fd.Show Then
        Me.pm_txtPDFPath.Text = fd.SelectedItems(1)
    End If
End Sub

Private Sub pm_btnRun_Click()
    Dim action As String
    action = Me.pm_cmbAction.Value
    Select Case action
        Case "Export Sheet to PDF"
            PM_ExportActiveSheet Me.pm_txtPDFPath.Text
        Case "Merge PDFs"
            PM_MergeSelectedPDFs Me.pm_txtPDFPath.Text
        Case "Split PDF"
            PM_SplitSelectedPDF Me.pm_txtPDFPath.Text
        Case "Delete Pages"
            PM_DeletePagesFromPDF Me.pm_txtPDFPath.Text
    End Select
End Sub

Private Sub pm_btnClose_Click()
    Unload Me
End Sub
