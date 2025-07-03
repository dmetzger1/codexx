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

Private Sub pmBtnBrowse_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select PDF File"
    fd.Filters.Clear
    fd.Filters.Add "PDF Files", "*.pdf"
    If fd.Show Then
        Me.pmTxtPath.Text = fd.SelectedItems(1)
    End If
End Sub

Private Sub pmBtnRun_Click()
    Dim action As String
    action = Me.pmCboAction.Value
    Select Case action
        Case "Export Sheet to PDF"
            PM_ExportActiveSheet Me.pmTxtPath.Text
        Case "Merge PDFs"
            PM_MergeSelectedPDFs Me.pmTxtPath.Text
        Case "Split PDF"
            PM_SplitSelectedPDF Me.pmTxtPath.Text
        Case "Delete Pages"
            PM_DeletePagesFromSelectedPDF Me.pmTxtPath.Text
    End Select
End Sub

Private Sub pmBtnClose_Click()
    Unload Me
End Sub
