VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "frmPM_PDFManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPM_Browse_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select PDF File"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf"
    If fd.Show Then
        Me.txtPM_PDFPath.Text = fd.SelectedItems(1)
    End If
End Sub

Private Sub btnPM_Run_Click()
    Dim action As String
    action = Me.cmbPM_Action.Value
    Select Case action
        Case "Export Sheet to PDF"
            PM_ExportActiveSheet Me.txtPM_PDFPath.Text
        Case "Merge PDFs"
            PM_MergeSelectedPDFs Me.txtPM_PDFPath.Text
        Case "Split PDF"
            PM_SplitSelectedPDF Me.txtPM_PDFPath.Text
        Case "Delete Pages"
            PM_DeletePagesFromSelectedPDF Me.txtPM_PDFPath.Text
    End Select
End Sub

Private Sub btnPM_Close_Click()
    Unload Me
End Sub
