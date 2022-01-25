Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form1
	' 测试数据Excel信息

	Dim 测项起始列号 As UInt32 = 21

	Dim 测项名称信息 As String() = { _
	      "50米跑（单位：秒）", _
	      "坐位体前屈（单位：厘米）", _
	      "一分钟跳绳（单位：次）", _
	      "一分钟仰卧起坐（单位：次）", _
	      "50米×8往返跑（单位：秒）", _
	      "立定跳远（单位：厘米）", _
	      "800米跑（单位：秒）", _
	      "1000米跑（单位：秒）", _
	      "引体向上（单位：次）"}

	' Excel模板信息

	Dim 工作表名称信息 As String() = { _
	      "1112", _
	      "1314", _
	      "1516女", _
	      "1516男", _
	      "212223女", _
	      "212223男", _
	      "313233女", _
	      "313233男"}

	Dim 学生评价等级的建议起始行号 As UInt32 = 2

	Dim 学生身体形态的建议起始行号 As UInt32 = 15

	Dim 学生肺活量的建议起始行号 As UInt32 = 28

	Dim 整体运动建议起始行号 As UInt32 = 41

	Dim 学生测试项信息起始行号 As UInt32 = 63

	Dim 整体运动建议信息 As UInt32() = {5, 5, 7, 7, 7, 7, 7, 7}

	' 单项指标
	' 学生身体素质测试结果的建议
	Dim 学生测试项信息 As UInt32() = { _
	      3, 0, 1, 2, 0, 0, _
	      4, 0, 1, 3, 2, 0, _
	      5, 0, 1, 3, 2, 4, _
	      5, 0, 1, 3, 2, 4, _
	      5, 0, 1, 3, 5, 6, _
	      5, 0, 1, 8, 5, 7, _
	      5, 0, 1, 3, 5, 6, _
	      5, 0, 1, 8, 5, 7}

	' 当前信息

	Dim 当前行号 As UInt32 = 1
	Dim 当前类别 As UInt32 = 1
	Dim 身高体重等级 As UInt32

	Dim 各等级百分比(4) As UInt32
	Dim 各身体形态百分比(4) As UInt32
	Dim 各身体机能百分比(4) As UInt32

	' 资源信息

	Dim wordApp As Word.Application
	Dim excelApp As Excel.Application

	Dim wordDocTmpl As Word.Document
	Dim wordDoc As Word.Document

	Dim excelWbTmpl As Excel.Workbook
	Dim excelWsTmpl As Excel.Worksheet
	Dim excelWb As Excel.Workbook
	Dim excelWs As Excel.Worksheet

	' 配置
	Const displayExcel As Boolean = True
	Const displayWord As Boolean = True

	Private Sub 装载应用()
		If wordApp Is Nothing Then
			wordApp = New Word.Application
			If displayWord Then wordApp.Visible = True
		End If
		If excelApp Is Nothing Then
			excelApp = New Excel.Application
			If displayExcel Then excelApp.Visible = True
		End If
	End Sub

	Private Sub 卸载应用()
		If Not wordApp Is Nothing Then
			wordApp.Quit()
			wordApp = Nothing
		End If
		If Not excelApp Is Nothing Then
			excelApp.Quit()
			excelApp = Nothing
		End If
	End Sub

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		' 装载应用()
	End Sub

	Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
		卸载应用()
	End Sub

	Private Sub 计算类别()
		Dim 年级 As String
		Dim 性别 As String
		年级 = excelWs.Range("E" & 当前行号).Text
		性别 = excelWs.Range("G" & 当前行号).Text
		Select Case 年级
			Case "一年级", "二年级"
				当前类别 = 0
			Case "三年级", "四年级"
				当前类别 = 1
			Case "五年级", "六年级"
				If 性别 = "男" Then
					当前类别 = 3
				Else
					当前类别 = 2
				End If
			Case "初一", "初二", "初三", "七年级", "八年级", "九年级"
				If 性别 = "男" Then
					当前类别 = 5
				Else
					当前类别 = 4
				End If
			Case "高一", "高二", "高三"
				If 性别 = "男" Then
					当前类别 = 7
				Else
					当前类别 = 6
				End If
		End Select
	End Sub

	Private Function 计算等级(ByVal 评价 As String)
		Select Case 评价
			Case "优秀"
				计算等级 = 0
			Case "良好"
				计算等级 = 1
			Case "及格"
				计算等级 = 2
			Case "不及格"
				计算等级 = 3
			Case Else
				计算等级 = 3
		End Select
	End Function

	Private Function 计算身体形态等级(ByVal 评价 As String)
		Select Case 评价
			Case "正常"
				计算身体形态等级 = 0
			Case "低体重"
				计算身体形态等级 = 1
			Case "超重"
				计算身体形态等级 = 2
			Case "肥胖"
				计算身体形态等级 = 3
			Case Else
				计算身体形态等级 = 1
		End Select
	End Function

	Private Function 计算百分比(ByRef 计数() As UInt32, ByRef 百分比() As UInt32)
		Dim 余数 As UInt32
		Dim 百分比和 As UInt32
		Dim i As UInt32

		For i = 0 To 4
			百分比(i) = 0
		Next

		If 计数(0) = 0 Then GoTo out

		百分比和 = 0
		For i = 1 To 3
			百分比(i - 1) = 10000 * (计数(i)) / 计数(0)
			余数 = 百分比(i - 1) Mod 10
			百分比(i - 1) = Int(百分比(i - 1) / 10)
			If 余数 >= 5 Then 百分比(i - 1) = 百分比(i - 1) + 1
			百分比和 = 百分比和 + 百分比(i - 1)
		Next
		百分比(3) = 1000 - 百分比和

out:
		计算百分比 = 0
	End Function

	Private Function 计算学校整体情况()
		Dim 各等级计数(5) As UInt32
		Dim 各身体形态计数(5) As UInt32
		Dim 各身体机能计数(5) As UInt32
		Dim 行号 As UInt32
		Dim 类别 As UInt32

		行号 = 2
		Do While True
			If excelWs.Range("B" & 行号).Text = "" Then Exit Do

			类别 = 计算等级(excelWs.Range("I" & 行号).Text)
			各等级计数(类别 + 1) = 各等级计数(类别 + 1) + 1
			各等级计数(0) = 各等级计数(0) + 1
			类别 = 计算身体形态等级(excelWs.Range("Q" & 行号).Text)
			各身体形态计数(类别 + 1) = 各身体形态计数(类别 + 1) + 1
			各身体形态计数(0) = 各身体形态计数(0) + 1
			类别 = 计算等级(excelWs.Range("T" & 行号).Text)
			各身体机能计数(类别 + 1) = 各身体机能计数(类别 + 1) + 1
			各身体机能计数(0) = 各身体机能计数(0) + 1

			行号 = 行号 + 1
		Loop

		计算百分比(各等级计数, 各等级百分比)
		计算百分比(各身体形态计数, 各身体形态百分比)
		计算百分比(各身体机能计数, 各身体机能百分比)

		计算学校整体情况 = 0
	End Function

	Private Sub 处理数据(ByRef 待处理文件 As String)
		Try
			excelWb = excelApp.Workbooks.Add(待处理文件)
			excelWs = excelWb.Sheets(1)
		Catch e As Exception
			GoTo out
		End Try

		If excelWs.Range("A" & 1).Text <> "ID" Then
			MsgBox("不识别的待处理文件:" & 待处理文件)
			GoTo out
		End If

		Try
			计算学校整体情况()

			当前行号 = 2
			Do While True
				If excelWs.Range("B" & 当前行号).Text = "" Then Exit Do

				计算类别()

				打开报告()

				生成报告()

				关闭报告()

				当前行号 = 当前行号 + 1

				If 当前行号 > 1 Then Exit Do
			Loop
		Catch e As Exception
			MsgBox("处理数据:" & e.Message & "\n" & e.StackTrace)
		End Try

out:
		If Not wordDoc Is Nothing Then
			wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
			wordDoc = Nothing
		End If
		excelWs = Nothing
		If Not excelWb Is Nothing Then
			excelWb.Close()
			excelWb = Nothing
		End If
	End Sub

	Private Sub 处理事件()
		Dim 对话框 As New System.Windows.Forms.OpenFileDialog
		Dim 待处理文件列表() As String

		With 对话框
			.InitialDirectory = Application.StartupPath
			.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
			.Multiselect = True
			.ShowDialog()
			待处理文件列表 = .FileNames
		End With

		If 待处理文件列表 Is Nothing Or 待处理文件列表.Length = 0 Then
			Exit Sub
		End If

		'Exit Sub

		装载应用()

		Try
			excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\Tmpl.xlsx")
		Catch e As Exception
			MsgBox("打开Excel模板: " & e.Message)
			GoTo out
		End Try

		' 循环处理所有Excel
		Try
			For i = 0 To 待处理文件列表.Length - 1
				' 处理一个Excel
				处理数据(待处理文件列表(i))
			Next
		Catch e As Exception
			MsgBox("处理文件列表: " & e.Message)
		End Try

		If Not excelWbTmpl Is Nothing Then
			excelWbTmpl.Close(False)
			excelWbTmpl = Nothing
		End If

		If Not wordDocTmpl Is Nothing Then
			wordDocTmpl.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
			wordDoc = Nothing
		End If

out:
		卸载应用()
	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		Try
			处理事件()
		Catch ex As Exception
			MsgBox("Click: " & ex.Message)
		End Try
	End Sub

	Function 打开报告()
		Dim docFullName As String
		Dim docPath As String

		docPath = Application.StartupPath & "\" & excelWs.Range("D" & 当前行号).Value2 & "\" & excelWs.Range("E" & 当前行号).Value2 & "\" & excelWs.Range("F" & 当前行号).Value2
		If Not Directory.Exists(docPath) Then
			docPath = Application.StartupPath & "\" & excelWs.Range("D" & 当前行号).Value2
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
			docPath = docPath & "\" & excelWs.Range("E" & 当前行号).Value2
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
			docPath = docPath & "\" & excelWs.Range("F" & 当前行号).Value2
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		End If
		docFullName = docPath & "\" _
		    & excelWs.Range("D" & 当前行号).Value2 & "_" & excelWs.Range("E" & 当前行号).Value2 & "_" & excelWs.Range("F" & 当前行号).Value2 & "_" & excelWs.Range("A" & 当前行号).Value2 & ".docx"

		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\Tmpl.docx")
		wordDoc.SaveAs(docFullName)
		wordDoc.Application.Activate()

		打开报告 = 0
	End Function

	Function 关闭报告()
		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		关闭报告 = 0
	End Function

	Function 生成首页()
		' 学校名称
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(2).Range.Text = excelWs.Range("D" & 当前行号).Value2
		' 年级班级
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(5).Range.Text = excelWs.Range("F" & 当前行号).Value2
		' 学生姓名
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(8).Range.Text = excelWs.Range("C" & 当前行号).Value2

		生成首页 = 0
	End Function

	Function 生成学生情况()
		' 学生识别号
		wordDoc.Tables(1).Cell(1, 2).Range.Text = excelWs.Range("A" & 当前行号).Text
		' 性别
		wordDoc.Tables(1).Cell(1, 4).Range.Text = excelWs.Range("G" & 当前行号).Text
		' 年级
		wordDoc.Tables(1).Cell(1, 6).Range.Text = excelWs.Range("E" & 当前行号).Text
		' 班级
		wordDoc.Tables(1).Cell(2, 2).Range.Text = excelWs.Range("F" & 当前行号).Text
		' 测试成绩
		wordDoc.Tables(1).Cell(2, 4).Range.Text = excelWs.Range("J" & 当前行号).Text
		' 测试等级
		wordDoc.Tables(1).Cell(2, 6).Range.Text = excelWs.Range("K" & 当前行号).Text
		' 综合成绩
		wordDoc.Tables(1).Cell(3, 2).Range.Text = excelWs.Range("H" & 当前行号).Text
		' 综合等级
		wordDoc.Tables(1).Cell(3, 4).Range.Text = excelWs.Range("I" & 当前行号).Text
		' 所在学校
		wordDoc.Tables(1).Cell(4, 2).Range.Text = excelWs.Range("D" & 当前行号).Text

		生成学生情况 = 0
	End Function

	Function 生成单项指标()
		Dim 表格位置 As UInt32 = 2
		Dim 测项序号 As UInt32 = 0
		Dim 内容 As String
		Dim idx As UInt32 = 0
		Dim i As UInt32 = 0

		' 身体形态
		内容 = excelWs.Range("O" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 2).Range.Text = 内容
		内容 = excelWs.Range("P" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 3).Range.Text = 内容
		内容 = excelWs.Range("Q" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 4).Range.Text = 内容
		' 身体机能
		内容 = excelWs.Range("R" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 2).Range.Text = 内容
		内容 = excelWs.Range("S" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 3).Range.Text = 内容
		内容 = excelWs.Range("T" & 当前行号).Text
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 4).Range.Text = 内容

		' 动态项
		idx = 当前类别 * 6
		For i = 1 To 学生测试项信息(idx)
			wordDoc.Tables(表格位置).Rows.Add()
			测项序号 = 学生测试项信息(idx + i)
			wordDoc.Tables(表格位置).Cell(3 + i, 1).Range.Text = 测项名称信息(测项序号)
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 0).Text
			If 内容 = "X" Then 内容 = ""
			wordDoc.Tables(表格位置).Cell(3 + i, 2).Range.Text = 内容
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 1).Text
			If 内容 = "X" Then 内容 = ""
			wordDoc.Tables(表格位置).Cell(3 + i, 3).Range.Text = 内容
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 2).Text
			If 内容 = "X" Then 内容 = ""
			wordDoc.Tables(表格位置).Cell(3 + i, 4).Range.Text = 内容
		Next

		生成单项指标 = 0
	End Function

	Function 生成各指标得分图表()
		Dim 表格位置 As UInt32 = 2
		Dim 图表工作表 As Excel.Worksheet
		Dim 测项序号 As UInt32 = 0
		Dim idx As UInt32 = 0
		Dim i As UInt32

		Try
			图表工作表 = excelWbTmpl.Sheets(工作表名称信息(当前类别) & "图表")
			图表工作表.Activate()
			图表工作表.Cells(1, 2).Value2 = excelWs.Range("P" & 当前行号).Text
			图表工作表.Cells(2, 2).Value2 = excelWs.Range("S" & 当前行号).Text

			' 动态项
			idx = 当前类别 * 6
			For i = 1 To 学生测试项信息(idx)
				测项序号 = 学生测试项信息(idx + i)
				图表工作表.Cells(2 + i, 2).Value2 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 1).Text
			Next

			图表工作表.Shapes.SelectAll()
			'图表工作表.Activate()
			图表工作表.Application.Selection.copy()
		Catch e As Exception
			MsgBox("生成各指标得分图表:" & e.Message)
			GoTo out
		End Try

		'excelWb.Activate()
		wordDoc.Tables(表格位置).Select()
		wordDoc.Application.Selection.MoveDown()
		wordDoc.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture)

out:
		生成各指标得分图表 = 0
	End Function

	Function 转换百分比(ByVal 数值 As UInt32)
		If 数值 = 0 Then
			转换百分比 = ""
		ElseIf 数值 < 10 Then
			转换百分比 = "0." & 数值
		Else
			转换百分比 = Int(数值 / 10) & "." & (数值 Mod 10)
		End If
	End Function

	Function 生成学校整体情况()
		Dim 表格位置 As UInt32 = 3
		Dim i As UInt32 = 0

		For i = 0 To 3
			wordDoc.Tables(表格位置).Cell(i + 2, 2).Range.Text = 转换百分比(各等级百分比(i))
			wordDoc.Tables(表格位置).Cell(i + 2, 4).Range.Text = 转换百分比(各身体形态百分比(i))
			wordDoc.Tables(表格位置).Cell(i + 2, 6).Range.Text = 转换百分比(各身体机能百分比(i))
		Next

		生成学校整体情况 = 0
	End Function

	Function 生成运动处方()
		Dim idx As UInt32
		Dim i As UInt32

		Dim wordFind As Word.Find
		wordFind = wordDoc.Application.Selection.Find
		wordFind.ClearFormatting()
		wordFind.Text = "运动处方"
		wordFind.Replacement.Text = ""
		wordFind.Forward = True
		wordFind.Wrap = Word.WdFindWrap.wdFindContinue
		wordFind.Format = False
		wordFind.MatchCase = False
		wordFind.MatchWholeWord = False
		wordFind.MatchByte = True
		wordFind.MatchWildcards = False
		wordFind.MatchSoundsLike = False
		wordFind.MatchAllWordForms = False
		wordDoc.Application.Selection.Find.Execute()
		wordDoc.Application.Selection.MoveDown(Word.WdUnits.wdLine, 1)

		excelWsTmpl = excelWbTmpl.Sheets(工作表名称信息(当前类别))

		'Dim 评价 As String
		Dim 等级 As UInt32

		' 学生评价等级的建议
		等级 = 计算等级(excelWsTmpl.Range("I" & 当前行号).Text)
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（一）" & excelWsTmpl.Range("A1").Text)
		wordDoc.Application.Selection.TypeParagraph()
		wordDoc.Application.Selection.Style = "正文1"
		wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (1 + 1 + 等级 * 3)).Text)
		wordDoc.Application.Selection.TypeParagraph()

		'评价 = excelWsTmpl.Range("Q" & 当前行号).Text
		Select Case excelWs.Range("Q" & 当前行号).Text
			Case "正常"
				身高体重等级 = 0
				等级 = 2
			Case "超重"
				身高体重等级 = 1
				等级 = 1
			Case "低体重"
				身高体重等级 = 2
				等级 = 3
			Case "肥胖"
				身高体重等级 = 3
				等级 = 0
		End Select

		' 学生身体形态的建议
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（二）" & excelWsTmpl.Range("A14").Text)
		wordDoc.Application.Selection.TypeParagraph()
		wordDoc.Application.Selection.Style = "正文1"
		wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (14 + 1 + 等级 * 3)).Text)
		wordDoc.Application.Selection.TypeParagraph()

		' 学生肺活量的建议
		等级 = 计算等级(excelWsTmpl.Range("T" & 当前行号).Text)
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（三）" & excelWsTmpl.Range("A27").Text)
		wordDoc.Application.Selection.TypeParagraph()
		wordDoc.Application.Selection.Style = "正文1"
		wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (27 + 1 + 等级 * 3)).Text)
		wordDoc.Application.Selection.TypeParagraph()

		' 整体运动建议
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（四）" & excelWsTmpl.Range("A40").Text)
		wordDoc.Application.Selection.TypeParagraph()
		For i = 0 To 整体运动建议信息(当前类别) - 1
			wordDoc.Application.Selection.Style = "主标题2"
			wordDoc.Application.Selection.TypeText(i + 1 & "." & excelWsTmpl.Range("A" & (40 + 1 + i * 3)).Text)
			wordDoc.Application.Selection.TypeParagraph()
			wordDoc.Application.Selection.Style = "正文1"
			wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (40 + 1 + i * 3)).Text)
			wordDoc.Application.Selection.TypeParagraph()
		Next

		'学生身体素质测试结果的建议
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（五）" & excelWsTmpl.Range("A62").Text)
		wordDoc.Application.Selection.TypeParagraph()
		For i = 0 To 学生测试项信息(当前类别 * 6) - 1
			' 第i个测项
			' 4个等级, 4个小等级，每个小等级3行
			idx = 62 + 1 + i * 4 * 4 * 3
			Dim 测项列号 As UInt32
			测项列号 = 测项起始列号 + 学生测试项信息(当前类别 * 6 + 1) * 3 + 2
			等级 = 计算等级(excelWs.Cells(当前行号, 测项列号).Text)
			idx = idx + 等级 * 4 * 3 + 身高体重等级

			' 需要加粗
			wordDoc.Application.Selection.Style = "主标题2"
			wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("A" & idx).Text)
			'wordDoc.Application.Selection.Paragraphs.First.
			wordDoc.Application.Selection.TypeParagraph()
			wordDoc.Application.Selection.Style = "正文1"
			wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & idx).Text)
			wordDoc.Application.Selection.TypeParagraph()
		Next

		生成运动处方 = 0
	End Function

	Function 生成报告()
		生成首页()

		生成学生情况()

		生成单项指标()

		生成各指标得分图表()

		生成学校整体情况()

		生成运动处方()

		生成报告 = 0
	End Function

	Public Sub New()

		' 此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()

		' 在 InitializeComponent() 调用之后添加任何初始化。

	End Sub
End Class


Public Class GanYuXueXiaoZhengTiQingKuang
	Dim dengJiBaiFengBiYouXiu As String
	Dim dengJiBaiFengBiLiangHao As String
	Dim dengJiBaiFengBiJiGe As String
	Dim dengJiBaiFengBiBuJiGe As String

	Dim shenTiXingTaiBaiFenBiZhengChang As String
	Dim shenTiXingTaiBaiFenBiDiTiZhong As String
	Dim shenTiXingTaiBaiFenBiChaoZhong As String
	Dim shenTiXingTaiBaiFenBiFeiPang As String

	Dim shenTiJiNengYouXiu As String
	Dim shenTiJiNengLh As String
	Dim shenTiJiNengJg As String
	Dim shenTiJiNengBjg As String
End Class

Public Class EntryType
	Dim hangHao As Long
	Dim yunDongChuFangSheetName As String


End Class