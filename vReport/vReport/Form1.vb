Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop

Public Class Form1
	' 测试数据Excel信息

	Dim 测项起始列号 As UInt32 = 21

	Dim 测项名称信息() As String = { _
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

	Dim 工作表名称信息() As String = { _
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

	Dim 整体运动建议信息() As UInt32 = {5, 5, 7, 7, 7, 7, 7, 7}

	' 单项指标
	' 学生身体素质测试结果的建议
	Dim 学生测试项信息() As UInt32 = { _
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

	Dim 学校名称 As String
	Dim 年级班级 As String
	Dim 学生姓名 As String

	Dim 各等级百分比(4) As UInt32
	Dim 各身体形态百分比(4) As UInt32
	Dim 各身体机能百分比(4) As UInt32

	Dim 待处理文件列表() As String

	' 资源信息

	Dim wk As Thread
	Dim wkStart As UInt32
	Dim wkDone As UInt32
	Dim wkExiting As UInt32

	Dim wordApp As Word.Application
	Dim excelApp As Excel.Application

	Dim wordDocTmpl As Word.Document
	Dim wordDoc As Word.Document

	Dim excelWbTmpl As Excel.Workbook
	Dim excelWsTmpl As Excel.Worksheet
	Dim excelWb As Excel.Workbook
	Dim excelWs As Excel.Worksheet

	' 配置
	Dim displayExcel As Boolean = False
	Dim displayWord As Boolean = False

	' 日志
	Dim logger As StreamWriter

	Private Sub 装载应用()
		logW("装载应用")
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
		logW("卸载应用")
		If Not wordApp Is Nothing Then
			wordApp.Quit()
			wordApp = Nothing
		End If
		If Not excelApp Is Nothing Then
			excelApp.Quit()
			excelApp = Nothing
		End If
	End Sub

	Private Delegate Sub logToUIDelegate(ByRef msg As String)
	Dim dlgt As New logToUIDelegate(AddressOf logtoUI)

	Sub logtoUI(ByRef msg As String)
		RichTextBox1.Text = msg & Chr(13) & Chr(10) & RichTextBox1.Text
	End Sub

	Protected Sub log(ByVal tag As String, ByVal msg As String)
		Dim m As String
		m = String.Format("[{0}] {1} {2}", Now, tag, msg)
		If tag <> "信息" Then RichTextBox1.BeginInvoke(dlgt, m)
		If Not logger Is Nothing Then logger.WriteLine(m)
	End Sub

	Protected Sub logFlush()
		If Not logger Is Nothing Then logger.Flush()
	End Sub

	Protected Sub logI(ByVal msg As String)
		log("信息", msg)
	End Sub

	Protected Sub logW(ByVal msg As String)
		log("警告", msg)
	End Sub

	Protected Sub logE(ByVal msg As String)
		log("错误", msg)
		logflush()
	End Sub

	Protected Sub logF(ByVal msg As String)
		log("致命", msg)
	End Sub

	Private Sub cancelReport(ByVal wait As UInt32)
		If wk Is Nothing Then Exit Sub
		Thread.VolatileWrite(wkExiting, 1)

retry:
		If Thread.VolatileRead(wkDone) = 1 Then
			wk.Join()
			wk = Nothing
			wkStart = 0
			Exit Sub
		End If

		If wait = 1 Then GoTo retry
	End Sub

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim ts As Date
		ts = Now
		logger = New StreamWriter(String.Format("log_{0:04}{1,2:d2}{2,2:d2}{3,2:d2}{4,2:d2}{5,2:d2}.txt", ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, ts.Second), True)
		logW("程序启动")

		' 装载应用()
	End Sub

	Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
		cancelReport(1)

		卸载应用()

		logW("程序终止")
		Try
			logger.Flush()
			logger.Close()
			logger = Nothing
		Catch ex As Exception
			MsgBox(ex.Message)
		Finally
			logger = Nothing
		End Try
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
		logW("年级: " & 年级)
		logW("性别: " & 性别)
		logW("计算类别: " & 当前类别)
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
		logI("评价: " & 计算等级)
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
		logI("计算身体形态等级: " & 计算身体形态等级)
	End Function

	Private Function 计算百分比(ByRef 计数() As UInt32, ByRef 百分比() As UInt32)
		Dim 余数 As UInt32
		Dim 百分比和 As UInt32
		Dim i As UInt32

		For i = 0 To 4
			百分比(i) = 0
		Next

		If 计数(0) = 0 Then GoTo out

		logI("计数(0): " & 计数(0))
		百分比和 = 0
		For i = 1 To 3
			百分比(i - 1) = 10000 * (计数(i)) / 计数(0)
			余数 = 百分比(i - 1) Mod 10
			百分比(i - 1) = Int(百分比(i - 1) / 10)
			logI("百分比" & (i - 1) & ": " & 百分比(i - 1))
			If 余数 >= 5 Then 百分比(i - 1) = 百分比(i - 1) + 1
			百分比和 = 百分比和 + 百分比(i - 1)
		Next
		百分比(3) = 1000 - 百分比和
		logI("百分比3: " & 百分比(3))

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

		Dim i As UInt32
		For i = 0 To 5
			logI("各等级计数" & i & ": " & 各等级计数(i))
			logI("各身体形态计数" & i & ": " & 各身体形态计数(i))
			logI("各身体机能计数" & i & ": " & 各身体机能计数(i))
		Next

		计算百分比(各等级计数, 各等级百分比)
		计算百分比(各身体形态计数, 各身体形态百分比)
		计算百分比(各身体机能计数, 各身体机能百分比)

		计算学校整体情况 = 0
	End Function

	Private Sub 处理数据(ByRef 待处理文件 As String)
		logI("开始 - 处理数据 " & 待处理文件)

		Try
			excelWb = excelApp.Workbooks.Add(待处理文件)
			excelWs = excelWb.Sheets(1)
		Catch e As Exception
			logE(e.Message)
			logE(e.StackTrace)
			GoTo out
		End Try

		If excelWs.Range("A" & 1).Text <> "ID" Then
			logE("不识别的待处理文件:" & 待处理文件)
			MsgBox("不识别的待处理文件:" & 待处理文件)
			GoTo out
		End If

		Try
			计算学校整体情况()

			当前行号 = 2
			Do While True
				logW("当前行号 " & 当前行号)

				If excelWs.Range("B" & 当前行号).Text = "" Then Exit Do

				计算类别()

				打开报告()

				生成报告()

				关闭报告()

				当前行号 = 当前行号 + 1

				If wkExiting Then Exit Do

				' Release版本2分钟处理42笔数据
				If 当前行号 > 10 Then Exit Do
			Loop
		Catch e As Exception
			logE("处理数据:" & e.Message)
			logE(e.StackTrace)
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
		logI("结束 - 处理数据 " & 待处理文件)
	End Sub

	Private Sub 处理事件()
		logI("开始 - 处理事件")

		logW("待处理文件列表: " & String.Join(",", 待处理文件列表))
		If 待处理文件列表 Is Nothing Or 待处理文件列表.Length = 0 Then
			logW("没有待处理文件列表")
			GoTo out
		End If

		'Exit Sub

		装载应用()

		Try
			excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\Tmpl.xlsx")
		Catch e As Exception
			logE("打开Excel模板: " & e.Message)
			logE(e.StackTrace)
			MsgBox("打开Excel模板: " & e.Message)
			GoTo out
		End Try

		logFlush()

		' 循环处理所有Excel
		Try
			For i = 0 To 待处理文件列表.Length - 1
				logW("开始 - 处理: " & 待处理文件列表(i))
				' 处理一个Excel
				处理数据(待处理文件列表(i))
				logW("结束 - 处理: " & 待处理文件列表(i))
				logFlush()
				If wkExiting Then Exit For
			Next
		Catch e As Exception
			logE("处理文件列表: " & e.Message)
			logE(e.StackTrace)
			MsgBox("处理文件列表: " & e.Message)
		End Try

		If Not excelWbTmpl Is Nothing Then
			logW("关闭excelWbTmpl")
			excelWbTmpl.Close(False)
			excelWbTmpl = Nothing
		End If

		If Not wordDocTmpl Is Nothing Then
			logW("关闭wordDocTmpl")
			wordDocTmpl.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
			wordDoc = Nothing
		End If

out:
		卸载应用()
		logI("结束 - 处理事件")
		logFlush()
	End Sub

	Private Sub Worker()
		Try
			处理事件()
		Catch ex As Exception
			logE(ex.Message)
			logE(ex.StackTrace)
		End Try

		Thread.VolatileWrite(wkDone, 1)
	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		If Thread.VolatileRead(wkStart) Then
			If Thread.VolatileRead(wkDone) = 0 Then
				MsgBox("正在处理中 ...")
				Exit Sub
			End If
			wk.Join()
			wk = Nothing
			wkStart = False
		End If

		Dim 对话框 As New System.Windows.Forms.OpenFileDialog

		logI("开始 - 处理点击事件")

		With 对话框
			.InitialDirectory = Application.StartupPath
			.Filter = "Excel files (*.xlsx;*.xls)|*.xlsx;*.xls"
			.Multiselect = True
			.ShowDialog()
			待处理文件列表 = .FileNames
		End With

		logW("待处理文件列表: " & String.Join(",", 待处理文件列表))
		If 待处理文件列表 Is Nothing Or 待处理文件列表.Length = 0 Then
			logW("没有待处理文件列表")
			GoTo out
		End If

		Thread.VolatileWrite(wkExiting, 0)
		Thread.VolatileWrite(wkDone, 0)
		wk = New Thread(AddressOf Worker)
		wk.Start()
		Thread.VolatileWrite(wkStart, 1)

out:
		logI("结束 - 处理点击事件")
	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
		cancelReport(0)
	End Sub

	Function 打开报告()
		Dim docFullName As String
		Dim docPath As String

		logW("开始 - 打开报告")

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

		logW("打开报告模板")
		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\Tmpl.docx")
		logW("保存报告: " & docFullName)
		wordDoc.SaveAs(docFullName)
		If displayWord Then wordDoc.Application.Activate()

		logW("结束 - 打开报告")

		打开报告 = 0
	End Function

	Function 关闭报告()
		logW("开始 - 关闭报告")

		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		logW("结束 - 关闭报告")

		关闭报告 = 0
	End Function

	Function 生成首页()
		logI("开始 - 生成首页")

		' 学校名称
		学校名称 = excelWs.Range("D" & 当前行号).Value2
		' 年级班级
		年级班级 = excelWs.Range("F" & 当前行号).Value2
		' 学生姓名
		学生姓名 = excelWs.Range("C" & 当前行号).Value2

		' 学校名称
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(2).Range.Text = 学校名称
		' 年级班级
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(5).Range.Text = 年级班级
		' 学生姓名
		wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(8).Range.Text = 学生姓名

		logW("学校名称 " & 学校名称)
		logW("年级班级 " & 年级班级)
		logW("学生姓名 " & 学生姓名)

		logI("结束 - 生成首页")

		生成首页 = 0
	End Function

	Function 生成学生情况()
		logI("开始 - 生成学生情况")

		' 姓名
		wordDoc.Tables(1).Cell(1, 2).Range.Text = excelWs.Range("C" & 当前行号).Text
		' 学生识别号
		wordDoc.Tables(1).Cell(1, 4).Range.Text = excelWs.Range("A" & 当前行号).Text
		' 性别
		wordDoc.Tables(1).Cell(1, 6).Range.Text = excelWs.Range("G" & 当前行号).Text
		' 年级
		wordDoc.Tables(1).Cell(2, 2).Range.Text = excelWs.Range("E" & 当前行号).Text
		' 班级
		wordDoc.Tables(1).Cell(2, 4).Range.Text = excelWs.Range("F" & 当前行号).Text
		' 测试成绩
		wordDoc.Tables(1).Cell(2, 6).Range.Text = excelWs.Range("J" & 当前行号).Text
		' 测试等级
		wordDoc.Tables(1).Cell(3, 2).Range.Text = excelWs.Range("K" & 当前行号).Text
		' 综合成绩
		wordDoc.Tables(1).Cell(3, 4).Range.Text = excelWs.Range("H" & 当前行号).Text
		' 综合等级
		wordDoc.Tables(1).Cell(3, 6).Range.Text = excelWs.Range("I" & 当前行号).Text
		' 所在学校
		wordDoc.Tables(1).Cell(4, 2).Range.Text = excelWs.Range("D" & 当前行号).Text

		logI("结束 - 生成学生情况")

		生成学生情况 = 0
	End Function

	Function 生成单项指标()
		Dim 表格位置 As UInt32 = 2
		Dim 测项序号 As UInt32 = 0
		Dim 内容 As String
		Dim idx As UInt32 = 0
		Dim i As UInt32 = 0

		logI("开始 - 生成单项指标")

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
			内容 = 测项名称信息(测项序号)
			logW(i & " " & idx & " 测项 " & 测项序号 & " " & 内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 1).Range.Text = 内容
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 0).Text
			If 内容 = "X" Then 内容 = ""
			wordDoc.Tables(表格位置).Cell(3 + i, 2).Range.Text = 内容
			logW(内容)
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 1).Text
			If 内容 = "X" Then 内容 = ""
			logW(内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 3).Range.Text = 内容
			内容 = excelWs.Cells(当前行号, 测项起始列号 + 3 * 测项序号 + 2).Text
			If 内容 = "X" Then 内容 = ""
			logW(内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 4).Range.Text = 内容
		Next

		logI("结束 - 生成单项指标")

		生成单项指标 = 0
	End Function

	Function 生成各指标得分图表()
		Dim 表格位置 As UInt32 = 2
		Dim 图表工作表 As Excel.Worksheet
		Dim 测项序号 As UInt32 = 0
		Dim idx As UInt32 = 0
		Dim i As UInt32

		logI("开始 - 生成各指标得分图表")

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
			logE("生成各指标得分图表:" & e.Message)
			logE(e.StackTrace)
			'MsgBox("生成各指标得分图表:" & e.Message)
			GoTo out
		End Try

		'excelWb.Activate()
		wordDoc.Tables(表格位置).Select()
		wordDoc.Application.Selection.MoveDown()
		wordDoc.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture)

out:
		logI("结束 - 生成各指标得分图表")

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
		logI("转换百分比: " & 数值 & " > " & 转换百分比)
	End Function

	Function 生成学校整体情况()
		logI("开始 - 生成学校整体情况")

		Dim 表格位置 As UInt32 = 3
		Dim i As UInt32 = 0

		For i = 0 To 3
			wordDoc.Tables(表格位置).Cell(i + 2, 2).Range.Text = 转换百分比(各等级百分比(i))
			wordDoc.Tables(表格位置).Cell(i + 2, 4).Range.Text = 转换百分比(各身体形态百分比(i))
			wordDoc.Tables(表格位置).Cell(i + 2, 6).Range.Text = 转换百分比(各身体机能百分比(i))
		Next

		logI("结束- 生成学校整体情况")

		生成学校整体情况 = 0
	End Function

	Function 生成运动处方()
		Dim idx As UInt32
		Dim i As UInt32

		logI("开始 - 生成运动处方")

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
		logW("等级 " & 等级)
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

		logI("结束 - 生成运动处方")

		生成运动处方 = 0
	End Function

	Function 生成报告()
		logI("开始 - 生成报告")

		生成首页()

		生成学生情况()

		生成单项指标()

		生成各指标得分图表()

		生成学校整体情况()

		生成运动处方()

		logI("结束 - 生成报告")

		生成报告 = 0
	End Function

	Public Sub New()

		' 此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()

		' 在 InitializeComponent() 调用之后添加任何初始化。

	End Sub

	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
End Class