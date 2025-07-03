Attribute VB_Name = "PXPDFOperations"
Option Explicit

'Exports the active sheet to PDF using the provided path
Public Sub PX_ExportActiveSheet(ByVal outputPath As String)
    If Len(outputPath) = 0 Then
        MsgBox "Please specify a PDF file path.", vbExclamation
        Exit Sub
    End If
    On Error GoTo ErrHandler
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=outputPath, Quality:=xlQualityStandard
    MsgBox "Export complete: " & outputPath, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Error exporting to PDF: " & Err.Description, vbCritical
End Sub

'--- Advanced operations below require Adobe Acrobat ---
'Ensure the Adobe Acrobat library is referenced in your VBA project.

'Merge multiple PDFs into one file using Acrobat PDDoc objects
Public Sub PX_MergePDFs(ByVal firstPDF As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim srcDoc As Acrobat.AcroPDDoc
    Dim fd As Office.FileDialog
    Dim item As Variant
    Dim outputPath As String

    If Len(firstPDF) = 0 Then
        MsgBox "Please specify the first PDF.", vbExclamation
        Exit Sub
    End If

    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(firstPDF) = False Then
        MsgBox "Unable to open: " & firstPDF, vbCritical
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select PDFs to merge"
    fd.AllowMultiSelect = True
    If fd.Show Then
        For Each item In fd.SelectedItems
            Set srcDoc = CreateObject("AcroExch.PDDoc")
            If srcDoc.Open(CStr(item)) Then
                pdDoc.InsertPages pdDoc.GetNumPages - 1, srcDoc, 0, srcDoc.GetNumPages, True
                srcDoc.Close
            End If
        Next item
    End If

    outputPath = Application.GetSaveAsFilename(InitialFileName:="Merged.pdf", _
                        FileFilter:="PDF Files (*.pdf), *.pdf")
    If outputPath <> "False" Then
        pdDoc.Save PDSaveFull, outputPath
        MsgBox "Merge complete: " & outputPath, vbInformation
    End If
    pdDoc.Close
End Sub

'Split the selected PDF into individual pages
Public Sub PX_SplitPDF(ByVal pdfPath As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim i As Long
    Dim newDoc As Acrobat.AcroPDDoc
    Dim outputFolder As String
    Dim outputPath As String

    If Len(pdfPath) = 0 Then Exit Sub

    outputFolder = PX_GetFolder("Select folder for split pages")
    If Len(outputFolder) = 0 Then Exit Sub

    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(pdfPath) = False Then
        MsgBox "Unable to open file", vbCritical
        Exit Sub
    End If

    For i = 0 To pdDoc.GetNumPages - 1
        Set newDoc = CreateObject("AcroExch.PDDoc")
        newDoc.Create
        newDoc.InsertPages -1, pdDoc, i, 1, True
        outputPath = outputFolder & Application.PathSeparator & _
                        "page_" & (i + 1) & ".pdf"
        newDoc.Save PDSaveFull, outputPath
        newDoc.Close
    Next i
    pdDoc.Close
    MsgBox "Split complete", vbInformation
End Sub

'Delete pages from the selected PDF
Public Sub PX_DeletePDFPages(ByVal pdfPath As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim response As String
    Dim parts() As String
    Dim i As Long

    response = InputBox("Pages to delete (comma separated):", "Delete Pages")
    If Len(response) = 0 Then Exit Sub

    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(pdfPath) = False Then
        MsgBox "Unable to open file", vbCritical
        Exit Sub
    End If

    parts = Split(response, ",")
    For i = UBound(parts) To LBound(parts) Step -1
        If IsNumeric(parts(i)) Then
            pdDoc.DeletePages CLng(parts(i)) - 1, CLng(parts(i)) - 1
        End If
    Next i
    pdDoc.Save PDSaveFull, pdfPath
    pdDoc.Close
    MsgBox "Delete pages complete", vbInformation
End Sub

'Helper to show a folder picker and return the selected path
Private Function PX_GetFolder(prompt As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = prompt
    If fd.Show Then PX_GetFolder = fd.SelectedItems(1)
End Function
