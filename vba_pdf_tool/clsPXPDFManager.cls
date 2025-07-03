VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "clsPXPDFManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPXFilePath As String

Public Property Get PXFilePath() As String
    PXFilePath = pPXFilePath
End Property

Public Property Let PXFilePath(ByVal value As String)
    pPXFilePath = value
End Property

'Exports a worksheet to PDF using the stored PXFilePath
Public Sub ExportWorksheetPX(ByVal ws As Worksheet)
    If Len(pPXFilePath) = 0 Then
        MsgBox "PXFilePath not set.", vbExclamation
        Exit Sub
    End If
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pPXFilePath, _
        Quality:=xlQualityStandard
End Sub

'Merge this PDF with additional ones and save to outputPath
Public Sub MergeWithPX(ByVal additionalPDFs As Collection, ByVal outputPath As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim srcDoc As Acrobat.AcroPDDoc
    Dim item As Variant

    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(pPXFilePath) = False Then
        MsgBox "Unable to open file", vbCritical
        Exit Sub
    End If

    For Each item In additionalPDFs
        Set srcDoc = CreateObject("AcroExch.PDDoc")
        If srcDoc.Open(CStr(item)) Then
            pdDoc.InsertPages pdDoc.GetNumPages - 1, srcDoc, 0, srcDoc.GetNumPages, True
            srcDoc.Close
        End If
    Next item
    pdDoc.Save PDSaveFull, outputPath
    pdDoc.Close
End Sub

'Split the PDF into individual pages in outputFolder
Public Sub SplitIntoPX(ByVal outputFolder As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim newDoc As Acrobat.AcroPDDoc
    Dim i As Long

    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(pPXFilePath) = False Then
        MsgBox "Unable to open file", vbCritical
        Exit Sub
    End If

    For i = 0 To pdDoc.GetNumPages - 1
        Set newDoc = CreateObject("AcroExch.PDDoc")
        newDoc.Create
        newDoc.InsertPages -1, pdDoc, i, 1, True
        newDoc.Save PDSaveFull, outputFolder & Application.PathSeparator & _
            "page_" & (i + 1) & ".pdf"
        newDoc.Close
    Next i
    pdDoc.Close
End Sub

'Delete specific pages from the PDF
Public Sub DeletePagesPX(ByVal pages As String)
    Dim pdDoc As Acrobat.AcroPDDoc
    Dim arr() As String
    Dim i As Long

    arr = Split(pages, ",")
    Set pdDoc = CreateObject("AcroExch.PDDoc")
    If pdDoc.Open(pPXFilePath) = False Then
        MsgBox "Unable to open file", vbCritical
        Exit Sub
    End If

    For i = UBound(arr) To LBound(arr) Step -1
        If IsNumeric(arr(i)) Then
            pdDoc.DeletePages CLng(arr(i)) - 1, CLng(arr(i)) - 1
        End If
    Next i
    pdDoc.Save PDSaveFull, pPXFilePath
    pdDoc.Close
End Sub
