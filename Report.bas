Attribute VB_Name = "Report"
Option Explicit

Type GanYuXueXiaoZhengTiQingKuang
    dengJiBaiFengBiYouXiu As String
    dengJiBaiFengBiLiangHao As String
    dengJiBaiFengBiJiGe As String
    dengJiBaiFengBiBuJiGe As String

    shenTiXingTaiBaiFenBiZhengChang As String
    shenTiXingTaiBaiFenBiDiTiZhong As String
    shenTiXingTaiBaiFenBiChaoZhong As String
    shenTiXingTaiBaiFenBiFeiPang As String

    shenTiJiNengYouXiu As String
    stjnLh As String
    stjnJg As String
    stjnBjg As String
End Type

Type EntryType
    hangHao As Long
    yunDongChuFangSheetName As String
End Type


Sub main()
    Dim gyxxztqk As GanYuXueXiaoZhengTiQingKuang
    Dim entry As EntryType
    Dim ret As Integer
    Dim row As Long

    ret = calcGyxxztqk(gyxxztqk)

    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    wordApp.Visible = True

    row = 2
    Do While True
        If Range("B" & row) = "" Then Exit Do
        ret = processOne(row, gyxxztqk, wordApp)
        row = row + 1
        'Exit Do
    Loop

    wordApp.Quit
End Sub

Function calcGyxxztqk(ByRef gyxxztqk As GanYuXueXiaoZhengTiQingKuang)

End Function

Function processOne(row As Long, _
    ByRef gyxxztqk As GanYuXueXiaoZhengTiQingKuang, _
    ByRef wordApp As Word.Application)
    Dim wordDoc As Word.Document
    Dim docTmplFullName As String
    Dim docFullName As String
    Dim docPath As String

    docPath = ThisWorkbook.Path & "\" _
        & Range("D" & row) & "\" & Range("F" & row) & "\" & Range("G" & row)
    Dim abc As String
    abc = Dir(docPath, vbDirectory)
    If Dir(docPath, vbDirectory) = Empty Then
        docPath = ThisWorkbook.Path & "\" & Range("D" & row)
    abc = Dir(docPath, vbDirectory)
        If Dir(docPath, vbDirectory) = Empty Then MkDir (docPath)
        docPath = docPath & "\" & Range("F" & row)
    abc = Dir(docPath, vbDirectory)
        If Dir(docPath, vbDirectory) = Empty Then MkDir (docPath)
        docPath = docPath & "\" & Range("G" & row)
    abc = Dir(docPath, vbDirectory)
        If Dir(docPath, vbDirectory) = Empty Then MkDir (docPath)
    End If
    docFullName = docPath & "\" _
        & Range("D" & row) & "_" & Range("F" & row) & "_" & Range("G" & row) & "_" & Range("B" & row) & ".docx"

    docTmplFullName = ThisWorkbook.Path & "\abc.docx"
    Set wordDoc = wordApp.Documents.Add(docTmplFullName)
    wordDoc.SaveAs docFullName
    wordDoc.Close wdSaveChanges
    Set wordDoc = Nothing
End Function
