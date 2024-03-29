﻿Imports System.IO
Imports System.Threading
Imports Microsoft.Office.Interop

Public Enum OpType
	StudentScore = 0
	StudentReport = 1
	AreaReport = 2
	SchoolReport = 3
	GradeReport = 4
	ClassReport = 5
	AllReport = 6
	AreaPercent = 7
End Enum

Public Class Form1
	' 测试数据Excel信息

	Dim QQQQ As Char = Chr(-24092)

	Dim 测项起始列号 As UInt32 = 0
	Dim 测项附加分起始列号 As UInt32 = 0

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
	Dim 学生测试项附加分信息() As UInt32 = { _
	  3, 0, 0, 1, 0, 0, _
	  4, 0, 0, 0, 1, 0, _
	  5, 0, 0, 0, 1, 0, _
	  5, 0, 0, 0, 1, 0, _
	  5, 0, 0, 1, 0, 1, _
	  5, 0, 0, 1, 0, 1, _
	  5, 0, 0, 1, 0, 1, _
	  5, 0, 0, 1, 0, 1}

	' 当前信息

	Dim 当前行号 As UInt32 = 1
	Dim 当前类别 As UInt32 = 1
	Dim 身高体重等级 As UInt32

	Dim 学校名称 As String
	Dim 年级班级 As String
	Dim 学生姓名 As String

	Dim 各等级百分比 As UInt32() = {0, 0, 0, 0}
	Dim 各身体形态百分比 As UInt32() = {0, 0, 0, 0}
	Dim 各身体机能百分比 As UInt32() = {0, 0, 0, 0}

	Dim 待处理文件列表() As String

	' 为了生成学校整体状态
	Dim 各等级计数 As UInt32() = {0, 0, 0, 0}
	Dim 各身体形态计数 As UInt32() = {0, 0, 0, 0}
	Dim 各身体机能计数 As UInt32() = {0, 0, 0, 0}

	Dim st As Student

	Dim 学校统计项() As 统计项 = { _
	   New 统计项(0, 0, "综合评定", New Int32() {1, 1, 1, 1}, "总"), _
	   New 统计项(1, 1, "身高体重等级", New Int32() {1, 1, 1, 1}, "形"), _
	   New 统计项(0, 2, "肺活量等级", New Int32() {1, 1, 1, 1}, "肺"), _
	   New 统计项(0, 3, "50米跑等级", New Int32() {1, 1, 1, 1}, "A50"), _
	   New 统计项(0, 4, "坐位体前屈等级", New Int32() {1, 1, 1, 1}, "屈"), _
	   New 统计项(0, 5, "一分钟跳绳等级", New Int32() {1, 0, 0, 1}, "绳"), _
	   New 统计项(0, 6, "一分钟仰卧起坐等级", New Int32() {1, 1, 1, 1}, "卧"), _
	   New 统计项(0, 7, "50米×8往返跑等级", New Int32() {1, 0, 0, 1}, "A8"), _
	   New 统计项(0, 8, "立定跳远等级", New Int32() {0, 1, 1, 1}, "立"), _
	   New 统计项(0, 9, "800米跑等级", New Int32() {0, 1, 1, 1}, "A800"), _
	   New 统计项(0, 10, "1000米跑等级", New Int32() {0, 1, 1, 1}, "A1000"), _
	   New 统计项(0, 11, "引体向上等级", New Int32() {0, 1, 1, 1}, "引") _
	}
	Dim 全区统计信息 As 统计信息 = New 统计信息()
	Dim 学校统计信息 As Dictionary(Of String, 统计信息) = New Dictionary(Of String, 统计信息)
	Dim 学校统计信息详细 As Dictionary(Of String, 学校信息详细) = New Dictionary(Of String, 学校信息详细)
	Dim schoolTemplateMap As New Dictionary(Of String, UInteger)

	Dim 学段模板 As String

	' 资源信息
	Dim wkType As Int32
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

	Dim asyncResultList As List(Of IAsyncResult) = New List(Of IAsyncResult)

	Dim 已经读取的行数 As UInt32
	Const 最大缓存行数 As UInt32 = 32
	Const 最大缓存列数 As UInt32 = 128
	Dim 数据缓存(最大缓存行数 - 1)() As String

	' 配置
	Dim displayExcel As Boolean = False
	Dim displayWord As Boolean = False

	Dim 列重命名0 As Dictionary(Of String, String)
	Dim 列重命名1 As Dictionary(Of String, String)
	Dim 列重命名2 As Dictionary(Of String, String)
	Dim 学校转学区表 As Dictionary(Of String, String)
	Dim 列名转列号表 As Dictionary(Of String, UInt32)

	Dim 转pdf As Int32

	Dim maxNumOfAdvise As Int32 = 6

	Dim useClipboard = 1

	' 日志
	Dim logger As StreamWriter

	Dim tmpName As String = String.Empty

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

	Private Delegate Sub logToUIDelegate(ByRef msgEntity As MsgEntity)
	Dim dlgt As New logToUIDelegate(AddressOf logtoUI)

	'Private sb As System.Text.StringBuilder = New System.Text.StringBuilder

	Sub logtoUI(ByRef msgEntity As MsgEntity)
		If msgEntity.type = MsgType.mtNormal Then
			RichTextBox1.Text = msgEntity.data & Chr(13) & Chr(10) & RichTextBox1.Text
			If RichTextBox1.Text.Length > 8192 Then
				RichTextBox1.Text = Strings.Left(RichTextBox1.Text, 4096)
			End If
		ElseIf msgEntity.type = MsgType.mtProgress Then
			ToolStripStatusLabel1.Text = msgEntity.data
		Else
		End If
	End Sub

	Sub purgeAsync()
		Dim res As IAsyncResult
		Dim cnt As UInt32 = 0

retry:
		res = Nothing
		SyncLock asyncResultList
			If asyncResultList.Count <> 0 Then
				If asyncResultList(0).IsCompleted Then
					res = asyncResultList(0)
					asyncResultList.RemoveAt(0)
				End If
			End If
		End SyncLock

		If Not res Is Nothing Then
			EndInvoke(res)
			cnt += 1
			GoTo retry
		End If

		'Debug.Print("Purge " & cnt)
	End Sub

	Sub sendInternal(ByVal type As MsgType, ByRef msg As String)
		Dim res As IAsyncResult
		Dim msgEntity As MsgEntity

		msgEntity = New MsgEntity()
		msgEntity.type = type
		msgEntity.data = msg
		res = Me.BeginInvoke(dlgt, msgEntity)
		SyncLock asyncResultList
			asyncResultList.Add(res)
		End SyncLock
	End Sub

	Sub sendMsg(ByRef msg As String)
		sendInternal(MsgType.mtNormal, msg)
	End Sub

	Sub sendProgress(ByRef msg As String)
		sendInternal(MsgType.mtProgress, msg)
	End Sub

	Protected Sub log(ByVal tag As String, ByVal msg As String)
		Dim fmtMsg As String
		fmtMsg = String.Format("[{0}] {1} {2}", Now, tag, msg)
		If tag = "错误" Or tag = "致命" Or tag = "报表" Then sendMsg(fmtMsg)
		If Not logger Is Nothing Then logger.WriteLine(fmtMsg)
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

	Protected Sub logR(ByVal msg As String)
		log("报表", msg)
	End Sub

	Protected Sub logE(ByVal msg As String)
		log("错误", msg)
		logFlush()
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
		logger = New StreamWriter(String.Format("log_{0}{1,2:d2}{2,2:d2}{3,2:d2}{4,2:d2}{5,2:d2}.txt", ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, ts.Second), True)
		logW("程序启动")
		logR(String.Format("生成时间: {0}", System.IO.File.GetLastWriteTime(Me.GetType().Assembly.Location)))

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
		年级 = 获取当前行数据("年级")
		性别 = 获取当前行数据("性别")
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
		'logR("当前行: " & 当前行号 & " 年级: " & 年级 & " 性别: " & 性别 & " 计算类别: " & 当前类别)
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
		'logI("评价: " & 计算等级)
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
				计算身体形态等级 = 0
		End Select
		'logI("计算身体形态等级: " & 计算身体形态等级)
	End Function

	Private Function 计算百分比(ByRef 计数() As UInt32, ByRef 百分比() As UInt32)
		Dim 百分比和 As UInt32 = 0
		Dim 总和 As UInt32 = 0
		Dim 余数(3) As UInt32
		Dim 系数 As UInt64 = 100000
		Dim idx As Int32
		Dim max As UInt32
		Dim i As UInt32
		Dim j As UInt32

		ReDim 余数(计数.Length - 1)

		For i = 0 To 计数.Length - 1
			百分比(i) = 0
			总和 += 计数(i)
		Next

		If 总和 = 0 Then GoTo out

		'logI("计数(0): " & 计数(0))

		百分比和 = 0
		For i = 0 To 计数.Length - 1
			百分比(i) = Int(系数 * (计数(i)) / 总和)
			余数(i) = 百分比(i) Mod 10
			百分比(i) = Int(百分比(i) / 10)
		Next

		总和 = 0
		For i = 0 To 计数.Length - 1
			总和 += 百分比(i)
		Next

		For i = 总和 To 10000
			If i >= 10000 Then Exit For
			idx = -1
			max = 0
			For j = 0 To 计数.Length - 1
				If 余数(j) > max Then
					idx = j
					max = 余数(j)
				End If
			Next
			If idx >= 0 Then
				百分比(idx) += 1
				余数(idx) = 0
			End If
		Next

		'For i = 0 To 计数.Length - 1
		'logI(String.Format("计数({0}): {1} {2}", i, 计数(i), 百分比(i)))
		'Next

out:
		计算百分比 = 0
	End Function

	Private Function 计算学校整体情况()
		'Dim 各等级计数(5) As UInt32
		'Dim 各身体形态计数(5) As UInt32
		'Dim 各身体机能计数(5) As UInt32
		'Dim 行号 As UInt32
		'Dim 类别 As UInt32

		'行号 = 2
		'Do While True
		'	If excelWs.Range("B" & 行号).Text = "" Then Exit Do

		'	类别 = 计算等级(excelWs.Range("I" & 行号).Text)
		'	各等级计数(类别 + 1) = 各等级计数(类别 + 1) + 1
		'	各等级计数(0) = 各等级计数(0) + 1
		'	类别 = 计算身体形态等级(excelWs.Range("Q" & 行号).Text)
		'	各身体形态计数(类别 + 1) = 各身体形态计数(类别 + 1) + 1
		'	各身体形态计数(0) = 各身体形态计数(0) + 1
		'	类别 = 计算等级(excelWs.Range("T" & 行号).Text)
		'	各身体机能计数(类别 + 1) = 各身体机能计数(类别 + 1) + 1
		'	各身体机能计数(0) = 各身体机能计数(0) + 1

		'	行号 = 行号 + 1
		'Loop

		'Dim i As UInt32
		'For i = 0 To 5
		'	logI("各等级计数" & i & ": " & 各等级计数(i))
		'	logI("各身体形态计数" & i & ": " & 各身体形态计数(i))
		'	logI("各身体机能计数" & i & ": " & 各身体机能计数(i))
		'Next

		'计算百分比(各等级计数, 各等级百分比)
		'计算百分比(各身体形态计数, 各身体形态百分比)
		'计算百分比(各身体机能计数, 各身体机能百分比)

		计算学校整体情况 = 0
	End Function

	Private Sub 处理单项等级(ByRef 项 As 统计项, ByVal 学段 As UInt32, ByRef 信息 As 统计信息)
		' 缺项不参与计算
		If 获取当前行数据("缺项数量") <> "0" And 项.序号 = 0 Then Exit Sub
		' 该学段没有此测项
		If 项.学段(学段) = 0 Then Exit Sub
		If 获取当前行数据(项.名称) = String.Empty Or 获取当前行数据(项.名称) = "X" Then Exit Sub

		Dim 等级 As UInt32
		If 项.类型 = 0 Then
			等级 = 计算等级(获取当前行数据(项.名称))
		Else
			等级 = 计算身体形态等级(获取当前行数据(项.名称))
		End If
		信息.等级(项.序号, 学段, 等级) += 1
		信息.等级(项.序号, 3, 等级) += 1
	End Sub

	Private Sub 处理统计信息(ByVal 统计学校 As Int32, ByRef 信息 As 统计信息)
		If 统计学校 = 1 Then
			Dim 区 As String

			区 = 获取当前行数据("所属区")
			If 区 <> String.Empty Then
				If 学校转学区表.ContainsKey(获取当前行数据("学校")) Then
					区 = 学校转学区表(获取当前行数据("学校"))
				End If
			End If
			信息.区 = 区
		End If

		信息.报名人数 += 1
		If 获取当前行数据("是否参测") = "是" Or 获取当前行数据("是否参测") = "缺项" Then
			信息.参测人数 += 1
			If 获取当前行数据("缺项数量") = "0" Then
				信息.完测人数 += 1
			End If
		End If
		If 获取当前行数据("附加分") <> String.Empty And 获取当前行数据("附加分") <> "0" And 获取当前行数据("附加分") <> "X" Then
			信息.加分人数 += 1
		End If

		Dim 学段 As UInt32
		Select Case 获取当前行数据("学段")
			Case "小学"
				学段 = 0
			Case "初中"
				学段 = 1
			Case "高中"
				学段 = 2
			Case Else
				学段 = 0
		End Select

		Dim i As UInt32
		For i = 0 To 学校统计项.Count() - 1
			处理单项等级(学校统计项(i), 学段, 信息)
		Next
	End Sub

	Private Sub 处理学校百分比(ByRef 信息 As 统计信息)
		Dim 源(3) As UInt32
		Dim 目(3) As UInt32

		源(2) = 0
		源(3) = 0

		源(0) = 信息.参测人数
		源(1) = 信息.报名人数 - 信息.参测人数
		计算百分比(源, 目)
		信息.参测比例 = 目(0)
		源(0) = 信息.完测人数
		源(1) = 信息.报名人数 - 信息.完测人数
		计算百分比(源, 目)
		信息.完测比例 = 目(0)

		Dim i As UInt32
		Dim j As UInt32
		Dim k As UInt32

		For i = 0 To 11
			For j = 0 To 3
				For k = 0 To 3
					源(k) = 信息.等级(i, j, k)
				Next
				计算百分比(源, 目)
				For k = 0 To 3
					信息.百分比(i, j, k) = 目(k)
				Next
			Next
		Next
	End Sub

	Function 打开学校报告(ByRef 区 As String, ByRef 学校 As String)
		Dim docFullName As String
		Dim docPath As String

		'logW("开始 - 打开报告")

		docPath = Application.StartupPath & "\" & 区
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docPath &= "\学校报告\"
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docPath &= "\" & 学校
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docFullName = docPath & "\" & 区 & "_" & 学校 & ".docx"

		logW("打开报告模板")
		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\学校报告模板" & 学段模板 & ".docx")
		logW("保存报告: " & docFullName)
		wordDoc.SaveAs(docFullName)
		If displayWord Then wordDoc.Application.Activate()

		'logW("结束 - 打开报告")

		打开学校报告 = 0
	End Function

	Function 关闭学校报告(ByRef 区 As String, ByRef 学校 As String)
		'logW("开始 - 关闭报告")

		If 转pdf = 1 Then
			Try
				Dim docFullName As String
				Dim docPath As String
				docPath = Application.StartupPath & "\" & 区 & "\学校报告\" & 学校
				docFullName = docPath & "\" & 区 & "_" & 学校 & ".pdf"

				wordDoc.SaveAs(docFullName, Word.WdSaveFormat.wdFormatPDF)
			Catch
			End Try
		End If
		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		'logW("结束 - 关闭报告")

		关闭学校报告 = 0
	End Function

	Function 打开年级报告(ByRef 区 As String, ByRef 学校 As String, ByRef 年级 As String)
		Dim docFullName As String
		Dim docPath As String

		'logW("开始 - 打开报告")

		docPath = Application.StartupPath & "\年级报告\" & 区
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docPath = docPath & "\" & 学校
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docFullName = docPath & "\" & 区 & "_" & 学校 & "_" & 年级 & ".docx"

		logW("打开报告模板")
		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\年级报告模板" & 学段模板 & ".docx")
		logW("保存报告: " & docFullName)
		wordDoc.SaveAs(docFullName)
		If displayWord Then wordDoc.Application.Activate()

		'logW("结束 - 打开报告")

		打开年级报告 = 0
	End Function

	Function 关闭年级报告(ByRef 区 As String, ByRef 学校 As String, ByRef 年级 As String)
		'logW("开始 - 关闭报告")

		If 转pdf = 1 Then
			Try
				Dim docFullName As String
				Dim docPath As String
				docPath = Application.StartupPath & "\年级报告\" & 区 & "\" & 学校
				docFullName = docPath & "\" & 区 & "_" & 学校 & "_" & 年级 & ".pdf"

				wordDoc.SaveAs(docFullName, Word.WdSaveFormat.wdFormatPDF)
			Catch
			End Try
		End If
		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		'logW("结束 - 关闭报告")

		关闭年级报告 = 0
	End Function

	Function 打开班级报告(ByRef 区 As String, ByRef 学校 As String, ByRef 年级 As String, ByRef 班级 As String)
		Dim docFullName As String
		Dim docPath As String

		'logW("开始 - 打开报告")

		docPath = Application.StartupPath & "\班级报告\" & 区
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docPath = docPath & "\" & 学校
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docPath = docPath & "\" & 年级
		If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		docFullName = docPath & "\" & 区 & "_" & 学校 & "_" & 年级 & "_" & 班级 & ".docx"

		logW("打开报告模板")
		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\班级报告模板" & 学段模板 & ".docx")
		logW("保存报告: " & docFullName)
		wordDoc.SaveAs(docFullName)
		If displayWord Then wordDoc.Application.Activate()

		'logW("结束 - 打开报告")

		打开班级报告 = 0
	End Function

	Function 关闭班级报告(ByRef 区 As String, ByRef 学校 As String, ByRef 年级 As String, ByRef 班级 As String)
		'logW("开始 - 关闭报告")

		If 转pdf = 1 Then
			Try
				Dim docFullName As String
				Dim docPath As String
				docPath = Application.StartupPath & "\班级报告\" & 区 & "\" & 学校 & "\" & 年级
				docFullName = docPath & "\" & 区 & "_" & 学校 & "_" & 年级 & "_" & 班级 & ".pdf"

				wordDoc.SaveAs(docFullName, Word.WdSaveFormat.wdFormatPDF)
			Catch
			End Try
		End If
		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		'logW("结束 - 关闭报告")

		关闭班级报告 = 0
	End Function

	Private Function 搜索(ByRef 关键字 As String)
		Dim wordFind As Word.Find
		wordFind = wordDoc.Application.Selection.Find
		wordFind.ClearFormatting()
		wordFind.Text = 关键字
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
		搜索 = wordDoc.Application.Selection.Find.Execute()
	End Function

	Private Function 比较(ByVal x As UInt32, ByVal y As UInt32)
		If x > y Then
			比较 = "大于"
		ElseIf x < y Then
			比较 = "小于"
		Else
			比较 = "等于"
		End If
	End Function

	Private 测试结果分析综述字体 As Word.Font = Nothing
	Private 测试结果分析文本字体 As Word.Font = Nothing
	Private 内容中的百分比字体 As Word.Font = Nothing

	Private Sub 应用字体(ByRef 字体 As Word.Font)
		wordDoc.Application.Selection.Font.Name = 字体.Name
		wordDoc.Application.Selection.Font.Bold = 字体.Bold
		wordDoc.Application.Selection.Font.Italic = 字体.Italic
		wordDoc.Application.Selection.Font.Size = 字体.Size
		wordDoc.Application.Selection.Font.Color = 字体.Color
	End Sub

	Private Sub 生成单个学校测试结果分析文本(ByRef 学校名称 As String, ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息, ByVal 综述 As Int32, ByRef 关键字 As String, ByRef 测项名称 As String, ByVal 项目 As UInt32, ByRef 类别名称 As String)
		Dim x() As UInt32 = {0, 0, 0, 0, 0, 0, 0, 0}
		Dim 字体 As Word.Font
		Dim 样式 As String

		If 综述 = 1 Then
			字体 = 测试结果分析综述字体
			样式 = "测试结果分析综述"
		Else
			字体 = 测试结果分析文本字体
			样式 = "测试结果分析文本"
		End If

		Dim 结果 As Boolean = 搜索(关键字)
		If 结果 = 0 Then
			Exit Sub
		End If
		x(0) = 学校信息.百分比(项目, 3, 0)
		x(1) = 学区信息.百分比(项目, 3, 0)
		x(2) = 学校信息.百分比(项目, 3, 3)
		x(3) = 学区信息.百分比(项目, 3, 3)
		wordDoc.Application.Selection.Paragraphs.First.Range.Text = ""
		wordDoc.Application.Selection.Style = 样式
		wordDoc.Application.Selection.TypeText(String.Format("{0}测试情况反馈，本{1}处于优秀等级的学生占比（", 测项名称, 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(0)))
		应用字体(字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(0), x(1))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(1)))
		应用字体(字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。本{0}处于不及格等级的学生占比（", 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(2)))
		应用字体(字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(2), x(3))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(3)))
		应用字体(字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。"))
	End Sub

	Private Sub 生成单个学校测试结果分析(ByRef 学校名称 As String, ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息, ByRef 名称 As String, ByRef 类别名称 As String)
		'Dim content As String
		Dim x() As UInt32 = {0, 0, 0, 0, 0, 0, 0, 0}
		'Dim d(7) As String

		搜索("结果综述")
		x(0) = 学校信息.百分比(0, 3, 0) + 学校信息.百分比(0, 3, 1)
		x(1) = 学区信息.百分比(0, 3, 0) + 学区信息.百分比(0, 3, 1)
		x(2) = 学校信息.百分比(0, 3, 3)
		x(3) = 学区信息.百分比(0, 3, 3)

		Dim 获取 As Int32 = 0
		If 测试结果分析综述字体 Is Nothing Then
			For Each 样式 As Word.Style In wordDoc.Styles
				If 样式.NameLocal = "测试结果分析综述" Then
					测试结果分析综述字体 = 样式.Font.Duplicate
					获取 += 1
				End If
				If 样式.NameLocal = "测试结果分析文本" Then
					测试结果分析文本字体 = 样式.Font.Duplicate
					获取 += 1
				End If
				If 样式.NameLocal = "内容中的百分比" Then
					内容中的百分比字体 = 样式.Font.Duplicate
					获取 += 1
				End If
				If 获取 >= 3 Then Exit For
			Next
		End If

		wordDoc.Application.Selection.Paragraphs.First.Range.Text = ""
		wordDoc.Application.Selection.Style = "测试结果分析综述"
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("本次体质健康标准测试情况反馈，{0}的学生综合成绩优良率（", 名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(0)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(0), x(1))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(1)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。学生综合成绩不及格率（"))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(2)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(2), x(3))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(3)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。"))

		搜索("身体形态综述")
		x(0) = 学校信息.百分比(1, 3, 0)
		x(1) = 学区信息.百分比(1, 3, 0)
		x(2) = 学校信息.百分比(1, 3, 1)
		x(3) = 学区信息.百分比(1, 3, 1)
		x(4) = 学校信息.百分比(1, 3, 2)
		x(5) = 学区信息.百分比(1, 3, 2)
		x(6) = 学校信息.百分比(1, 3, 3)
		x(7) = 学区信息.百分比(1, 3, 3)
		wordDoc.Application.Selection.Paragraphs.First.Range.Text = ""
		wordDoc.Application.Selection.Style = "测试结果分析综述"
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("体重指数测试情况反馈，本{0}处于正常体重水平的学生占比（", 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(0)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(0), x(1))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(1)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。本{0}处于低体重水平的学生占比（", 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(2)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(2), x(3))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(3)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。本{0}处于超重水平的学生占比（", 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(4)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(4), x(5))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(5)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。本{0}处于肥胖水平的学生占比（", 类别名称))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(6)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("），{0}总体测试数据平均水平（", 比较(x(6), x(7))))
		应用字体(内容中的百分比字体)
		wordDoc.Application.Selection.TypeText(转换完整百分比(x(7)))
		应用字体(测试结果分析综述字体)
		wordDoc.Application.Selection.TypeText(String.Format("）。"))

		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 1, "身体机能综述", "肺活量", 2, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "速度素质文本", "50米跑", 3, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "柔韧素质文本", "坐位体前屈", 4, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "力量素质文本1", "1分钟仰卧起坐", 6, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "力量素质文本2", "引体向上", 11, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "协调素质文本", "1分钟跳绳", 5, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "耐力素质文本小", "50米×8往返跑", 7, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "耐力素质文本女", "800米跑", 9, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "耐力素质文本男", "1000米跑", 10, 类别名称)
		生成单个学校测试结果分析文本(学校名称, 学校信息, 学区信息, 0, "爆发力素质文本", "立定跳远", 8, 类别名称)
	End Sub

	Private Sub 生成学校报告图表(ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息)
		Dim i As Int32

		excelWsTmpl.Cells(2, 2).Value2 = 学校信息.参测人数
		excelWsTmpl.Cells(3, 2).Value2 = 学校信息.报名人数 - 学校信息.参测人数
		excelWsTmpl.Cells(12, 2).Value2 = 学校信息.完测人数
		excelWsTmpl.Cells(13, 2).Value2 = 学校信息.参测人数 - 学校信息.完测人数
		excelWsTmpl.Cells(14, 2).Value2 = 学校信息.报名人数 - 学校信息.参测人数
		excelWsTmpl.Cells(32, 2).Value2 = 转换完整百分比(学校信息.百分比(0, 3, 3))
		excelWsTmpl.Cells(32, 3).Value2 = 转换完整百分比(学校信息.百分比(0, 3, 2))
		excelWsTmpl.Cells(32, 4).Value2 = 转换完整百分比(学校信息.百分比(0, 3, 1))
		excelWsTmpl.Cells(32, 5).Value2 = 转换完整百分比(学校信息.百分比(0, 3, 0))
		excelWsTmpl.Cells(33, 2).Value2 = 转换完整百分比(学区信息.百分比(0, 3, 3))
		excelWsTmpl.Cells(33, 3).Value2 = 转换完整百分比(学区信息.百分比(0, 3, 2))
		excelWsTmpl.Cells(33, 4).Value2 = 转换完整百分比(学区信息.百分比(0, 3, 1))
		excelWsTmpl.Cells(33, 5).Value2 = 转换完整百分比(学区信息.百分比(0, 3, 0))

		For i = 1 To wordDoc.InlineShapes.Count
			wordDoc.InlineShapes(i).Select()
			wordDoc.InlineShapes(i).Delete()
			If False Then
				Dim ct As Excel.Chart = excelWsTmpl.ChartObjects(i).Chart
				ct.Export(tmpName, "GIF")
				wordDoc.Application.Selection.InlineShapes.AddPicture(tmpName, False, True)
			Else
				Dim retry As Int32
				Dim suc As Int32
				suc = 0
				For retry = 1 To 20
					Try
						excelWsTmpl.Shapes().Item(i).Select() '.SelectAll()
						System.Windows.Forms.Clipboard.Clear()
						excelWsTmpl.Application.Selection.copy()
						'wordDoc.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture)
						wordDoc.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting)
						suc = 1
					Catch ex As Exception
						If retry = 20 Then
							Throw ex
						End If
						logW("重试 " & retry)
						logW(ex.Message)
						logW(ex.StackTrace)
					End Try
					If suc = 1 Then
						Exit For
					End If
				Next
			End If
		Next
	End Sub
	Private Function 获取学段索引(t As String)
		Dim i As Int32 = 0
		If t.Contains("小") Then
			i = 0
		ElseIf t.Contains("初") Then
			i = 1
		ElseIf t.Contains("高") Then
			i = 2
		Else
			i = 3
		End If
		获取学段索引 = i
	End Function

	Private Function 获取测项索引(t As String)
		Dim i As Int32 = 0
		Dim j As Int32 = 学校统计项.Length - 1
		Dim k As Int32 = -1
		For i = 0 To j
			If t.Contains(学校统计项(i).关键字) Then
				' not a perfect match
				If 学校统计项(i).关键字 = "A8" And t.Contains("A800") Then Continue For
				k = i
				Exit For
			End If
		Next
		获取测项索引 = k
	End Function

	Private Sub 生成单个学校报告(ByRef 学校名称 As String, ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息)
		Dim i As Int32
		Dim j As Int32
		Dim k As Int32
		Dim r As Int32
		Dim c As Int32
		Dim t As String

		logI("开始 - 打开学校报告")
		打开学校报告(学校信息.区, 学校名称)
		logI("结束 - 打开学校报告")

		' 生成学校名称
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("学校名称") Then
				wordDoc.Paragraphs(i).Range.Text = 学校名称
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		logI("结束 - 生成学校名称")

		' 生成表格
		For c = 1 To wordDoc.Tables.Count
			Dim table As Word.Table

			table = wordDoc.Tables(c)
			table.Select()
			t = table.Cell(1, 1).Range.Text

			'logI("处理表格" & i & "一共" & wordDoc.Tables.Count & "名称" & val)
			If t.Contains("学校") Then
				table.Cell(1, 2).Range.Text = 学校名称
				table.Cell(2, 3).Range.Text = 学校信息.报名人数
				table.Cell(2, 5).Range.Text = 学校信息.参测人数
				table.Cell(2, 7).Range.Text = 学校信息.完测人数
				table.Cell(3, 3).Range.Text = 学校信息.加分人数
				table.Cell(3, 5).Range.Text = 转换完整百分比(学校信息.参测比例)
				table.Cell(3, 7).Range.Text = 转换完整百分比(学校信息.完测比例)
				table.Cell(4, 3).Range.Text = 学区信息.报名人数
				table.Cell(4, 5).Range.Text = 学区信息.参测人数
				table.Cell(4, 7).Range.Text = 学区信息.完测人数
				table.Cell(5, 3).Range.Text = 学区信息.加分人数
				table.Cell(5, 5).Range.Text = 转换完整百分比(学区信息.参测比例)
				table.Cell(5, 7).Range.Text = 转换完整百分比(学区信息.完测比例)
			ElseIf t.Contains("综合评定等级") Then
				For r = 2 To table.Rows.Count
					j = 获取学段索引(table.Cell(r, 3).Range.Text)
					For k = 0 To 3
						table.Cell(r + 0, 3 + k).Range.Text = 转换完整百分比(学校信息.百分比(0, j, k))
						table.Cell(r + 1, 3 + k).Range.Text = 转换完整百分比(学区信息.百分比(0, j, k))
					Next
					r += 1
				Next
			ElseIf t.Contains("评价等级") Then
				j = 获取学段索引(table.Cell(2, 3).Range.Text)
				For k = 0 To 3
					table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(1, j, k))
				Next
				For r = 4 To table.Rows.Count
					i = 获取测项索引(table.Cell(r, 2).Range.Text)
					If i >= 0 Then
						For k = 0 To 3
							table.Cell(r, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, j, k))
						Next
					End If
				Next
			ElseIf t.Contains("等级") Then
				i = 获取测项索引(table.Cell(2, 3).Range.Text)
				If i >= 0 Then
					For r = 2 To table.Rows.Count
						j = 获取学段索引(table.Cell(r, 3).Range.Text)
						If 学校统计项(i).学段(j) <> 0 Then
							For k = 0 To 3
								table.Cell(r + 0, 3 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, j, k))
								table.Cell(r + 1, 3 + k).Range.Text = 转换完整百分比(学区信息.百分比(i, j, k))
							Next
							' 本校
							' 全区
							r += 1
						End If
					Next
				End If
			End If
		Next

		logI("结束 - 处理表格")
		' 生成图表
		生成学校报告图表(学校信息, 学区信息)

		logI("结束 - 生成学校图表")
		' 生成文本
		生成单个学校测试结果分析(学校名称, 学校信息, 学区信息, 学校名称, "校")

		logI("结束 - 生成学校结果分析")
		关闭学校报告(学校信息.区, 学校名称)

		logI("结束 - 关闭学校报告")
	End Sub

	Private Sub 生成单个年级报告(ByRef 学校名称 As String, ByVal 年级 As UInt32, ByRef 年级名称 As String, ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息)
		Dim i As Int32
		'Dim j As Int32
		Dim k As Int32
		Dim r As Int32
		Dim c As Int32
		Dim t As String

		logI("开始 - 打开年级报告")
		打开年级报告(学校信息.区, 学校名称, 年级名称)
		logI("结束 - 打开年级报告")

		' 生成学校名称 年级名称
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("学校名称") Then
				wordDoc.Paragraphs(i).Range.Select()
				wordDoc.Application.Selection.Delete()
				wordDoc.Application.Selection.TypeText(学校名称)
				wordDoc.Application.Selection.TypeParagraph()
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("年级名称") Then
				wordDoc.Paragraphs(i).Range.Select()
				wordDoc.Application.Selection.Delete()
				wordDoc.Application.Selection.TypeText(年级名称)
				wordDoc.Application.Selection.TypeParagraph()
				' 使用同样样式
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		logI("结束 - 生成学校名称年级名称")

		' 生成表格
		For c = 1 To wordDoc.Tables.Count
			Dim table As Word.Table

			table = wordDoc.Tables(c)
			table.Select()
			t = table.Cell(1, 1).Range.Text

			'logR("处理表格" & i & "一共" & wordDoc.Tables.Count & "名称" & val)
			If t.Contains("学校") Then
				table.Cell(1, 2).Range.Text = 学校名称
				table.Cell(2, 2).Range.Text = 年级名称
				table.Cell(3, 3).Range.Text = 学校信息.报名人数
				table.Cell(3, 5).Range.Text = 学校信息.参测人数
				table.Cell(3, 7).Range.Text = 学校信息.完测人数
				table.Cell(4, 3).Range.Text = 学校信息.加分人数
				table.Cell(4, 5).Range.Text = 转换完整百分比(学校信息.参测比例)
				table.Cell(4, 7).Range.Text = 转换完整百分比(学校信息.完测比例)
				table.Cell(5, 3).Range.Text = 学区信息.报名人数
				table.Cell(5, 5).Range.Text = 学区信息.参测人数
				table.Cell(5, 7).Range.Text = 学区信息.完测人数
				table.Cell(6, 3).Range.Text = 学区信息.加分人数
				table.Cell(6, 5).Range.Text = 转换完整百分比(学区信息.参测比例)
				table.Cell(6, 7).Range.Text = 转换完整百分比(学区信息.完测比例)
			ElseIf t.Contains("综合评定等级") Then
				' 优良中差
				For k = 0 To 3
					table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(0, 3, k))
					table.Cell(3, 2 + k).Range.Text = 转换完整百分比(学区信息.百分比(0, 3, k))
				Next
			ElseIf t.Contains("评价等级") Then
				For k = 0 To 3
					table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(1, 3, k))
				Next
				For r = 4 To table.Rows.Count
					i = 获取测项索引(table.Cell(r, 2).Range.Text)
					For k = 0 To 3
						table.Cell(r, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, 3, k))
					Next
				Next
			ElseIf t.Contains("等级") Then
				i = 获取测项索引(table.Cell(2, 3).Range.Text)
				If i >= 0 Then
					For k = 0 To 3
						table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, 3, k))
						table.Cell(3, 2 + k).Range.Text = 转换完整百分比(学区信息.百分比(i, 3, k))
					Next
				End If
			End If
		Next

		logI("结束 - 处理表格")
		' 生成图表
		生成学校报告图表(学校信息, 学区信息)

		logI("结束 - 生成年级图表")
		' 生成文本
		生成单个学校测试结果分析(学校名称, 学校信息, 学区信息, 年级名称, "年级")

		logI("结束 - 生成年级结果分析")
		关闭年级报告(学校信息.区, 学校名称, 年级名称)

		logI("结束 - 关闭年级报告")
	End Sub

	Private Sub 生成单个班级报告(ByRef 学校名称 As String, ByRef 年级名称 As String, ByRef 班级名称 As String, ByRef 学校信息 As 统计信息, ByRef 学区信息 As 统计信息)
		Dim i As Int32
		'Dim j As Int32
		Dim k As Int32
		Dim r As Int32
		Dim c As Int32
		Dim t As String

		logI("开始 - 打开班级报告")
		打开班级报告(学校信息.区, 学校名称, 年级名称, 班级名称)
		logI("结束 - 打开班级报告")

		Dim a As Int32 = 0
		Dim b As Int32 = 0

		' 生成学校名称 年级名称
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("学校名称") Then
				wordDoc.Paragraphs(i).Range.Select()
				wordDoc.Application.Selection.Delete()
				wordDoc.Application.Selection.TypeText(学校名称)
				wordDoc.Application.Selection.TypeParagraph()
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("年级名称") Then
				wordDoc.Paragraphs(i).Range.Select()
				wordDoc.Application.Selection.Delete()
				wordDoc.Application.Selection.TypeText(年级名称)
				wordDoc.Application.Selection.TypeParagraph()
				' 使用同样样式
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		For i = 1 To wordDoc.Paragraphs.Count
			t = wordDoc.Paragraphs(i).Range.Text()
			If t.Contains("班级名称") Then
				wordDoc.Paragraphs(i).Range.Select()
				wordDoc.Application.Selection.Delete()
				wordDoc.Application.Selection.TypeText(班级名称)
				wordDoc.Application.Selection.TypeParagraph()
				' 使用同样样式
				wordDoc.Paragraphs(i).Style = "学校名称"
				wordDoc.Paragraphs(i).Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
				Exit For
			End If
		Next
		logI("结束 - 生成学校名称年级名称")

		' 生成表格
		For c = 1 To wordDoc.Tables.Count
			Dim table As Word.Table

			table = wordDoc.Tables(c)
			table.Select()
			t = table.Cell(1, 1).Range.Text

			'logR("处理表格" & i & "一共" & wordDoc.Tables.Count & "名称" & val)
			If t.Contains("学校") Then
				table.Cell(1, 2).Range.Text = 学校名称
				table.Cell(2, 2).Range.Text = 年级名称
				table.Cell(3, 2).Range.Text = 班级名称
				table.Cell(4, 3).Range.Text = 学校信息.报名人数
				table.Cell(4, 5).Range.Text = 学校信息.参测人数
				table.Cell(4, 7).Range.Text = 学校信息.完测人数
				table.Cell(5, 3).Range.Text = 学校信息.加分人数
				table.Cell(5, 5).Range.Text = 转换完整百分比(学校信息.参测比例)
				table.Cell(5, 7).Range.Text = 转换完整百分比(学校信息.完测比例)
				table.Cell(6, 3).Range.Text = 学区信息.报名人数
				table.Cell(6, 5).Range.Text = 学区信息.参测人数
				table.Cell(6, 7).Range.Text = 学区信息.完测人数
				table.Cell(7, 3).Range.Text = 学区信息.加分人数
				table.Cell(7, 5).Range.Text = 转换完整百分比(学区信息.参测比例)
				table.Cell(7, 7).Range.Text = 转换完整百分比(学区信息.完测比例)
			ElseIf t.Contains("综合评定等级") Then
				' 优良中差
				For k = 0 To 3
					table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(0, 3, k))
					table.Cell(3, 2 + k).Range.Text = 转换完整百分比(学区信息.百分比(0, 3, k))
				Next
			ElseIf t.Contains("评价等级") Then
				For k = 0 To 3
					table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(1, 3, k))
				Next
				For r = 4 To table.Rows.Count
					i = 获取测项索引(table.Cell(r, 2).Range.Text)
					For k = 0 To 3
						table.Cell(r, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, 3, k))
					Next
				Next
			ElseIf t.Contains("等级") Then
				i = 获取测项索引(table.Cell(2, 3).Range.Text)
				If i >= 0 Then
					For k = 0 To 3
						table.Cell(2, 2 + k).Range.Text = 转换完整百分比(学校信息.百分比(i, 3, k))
						table.Cell(3, 2 + k).Range.Text = 转换完整百分比(学区信息.百分比(i, 3, k))
					Next
				End If
			End If
		Next

		logI("结束 - 处理表格")
		' 生成图表
		生成学校报告图表(学校信息, 学区信息)

		logI("结束 - 生成班级图表")
		' 生成文本
		生成单个学校测试结果分析(学校名称, 学校信息, 学区信息, 班级名称, "班")

		logI("结束 - 生成班级结果分析")
		关闭班级报告(学校信息.区, 学校名称, 年级名称, 班级名称)

		logI("结束 - 关闭班级报告")
	End Sub

	Private Sub SelectSchoolTemplate(ByRef school As String)
		Dim s As UInt32 = 0

		If schoolTemplateMap.ContainsKey(school) Then
			s = schoolTemplateMap(school)
		End If
		If s = 1 Then
			学段模板 = "（小学版）"
		ElseIf s = 2 Then
			学段模板 = "（中学版）"
		Else
			学段模板 = "（全学段版）"
		End If

		logR("学校名称:" & school & " 学段模板:" & 学段模板)
	End Sub

	Private Sub 生成学校报告(ByVal 生成何种数据 As UInt32, ByVal 共几个文件 As UInt32, ByVal 第几个文件 As UInt32, ByRef 待处理文件 As String, ByRef excelWs As Excel.Worksheet)
		' 处理Excel，生成报表
		Dim 默认目录 As String
		默认目录 = Path.GetFileNameWithoutExtension(待处理文件)

		全区统计信息 = New 统计信息()
		学校统计信息.Clear()
		学校统计信息详细.Clear()
		schoolTemplateMap.Clear()

		列名转列号表.Clear()
		当前行号 = 1
		已经读取的行数 = 0
		预取数据到缓存(excelWs)
		生成列信息表格(列重命名2)

		Try
			Dim s As Int32 = 0

			Do While True
				移动到下一行()
				预取数据到缓存(excelWs)

				If 获取当前行数据("姓名") = String.Empty Then
					logI(String.Format("共{0}个文件。当前处理第{1}个文件的第{2}行。处理完毕。", 共几个文件, 第几个文件, 当前行号))
					Exit Do
				End If

				logI("当前行: " & 当前行号 & " 姓名 " & 获取当前行数据("姓名") & " 年级: " & 获取当前行数据("年级") & " 性别: " & 获取当前行数据("性别"))
				Dim 学校名称 As String
				学校名称 = 获取当前行数据("学校")
				If 学校名称 = String.Empty Then
					logR("非法的学校名称")
					Continue Do
				End If

				处理统计信息(0, 全区统计信息)
				If Not 学校统计信息.ContainsKey(学校名称) Then
					学校统计信息.Add(学校名称, New 统计信息())
				End If
				处理统计信息(1, 学校统计信息(学校名称))
				If 学校统计信息(学校名称).区 = String.Empty Then
					学校统计信息(学校名称).区 = 默认目录
				End If

				' 年级
				Dim g As UInt32
				For g = 0 To gradeNameTbl.Count - 1
					If 获取当前行数据("年级") = gradeNameTbl(g) Then
						Exit For
					End If
				Next

				If Not schoolTemplateMap.ContainsKey(学校名称) Then
					schoolTemplateMap.Add(学校名称, 0)
				End If

				If g < 6 Then
					schoolTemplateMap(学校名称) = schoolTemplateMap(学校名称) Or 1
				Else
					schoolTemplateMap(学校名称) = schoolTemplateMap(学校名称) Or 2
				End If

				If Not 学校统计信息详细.ContainsKey(学校名称) Then 学校统计信息详细(学校名称) = New 学校信息详细()
				If 学校统计信息详细(学校名称).年级统计信息(g) Is Nothing Then 学校统计信息详细(学校名称).年级统计信息(g) = New 统计信息()
				处理统计信息(1, 学校统计信息详细(学校名称).年级统计信息(g))
				If Not 学校统计信息详细(学校名称).班级信息详细(g).ContainsKey(获取当前行数据("班级")) Then
					学校统计信息详细(学校名称).班级信息详细(g)(获取当前行数据("班级")) = New 统计信息()
				End If
				处理统计信息(1, 学校统计信息详细(学校名称).班级信息详细(g)(获取当前行数据("班级")))

				logI(String.Format("共{0}个文件。当前处理第{1}个文件的第{2}行", 共几个文件, 第几个文件, 当前行号))

				purgeAsync()

				If wkExiting Then GoTo out
			Loop

			Dim i As UInt32
			For i = 0 To 学校统计信息.Count - 1
				logR(String.Format("文件 {0}/{1}，学校 {2}/{3}，进行中 ...", 第几个文件, 共几个文件, i + 1, 学校统计信息.Count))
				logR("正在处理 " & 学校统计信息.ElementAt(i).Key)
				If 生成何种数据 = OpType.SchoolReport Then
					处理学校百分比(全区统计信息)
					处理学校百分比(学校统计信息.ElementAt(i).Value)
					SelectSchoolTemplate(学校统计信息.ElementAt(i).Key)
					生成单个学校报告(学校统计信息.ElementAt(i).Key, 学校统计信息.ElementAt(i).Value, 全区统计信息)
					If wkExiting Then GoTo out
				ElseIf 生成何种数据 = OpType.GradeReport Then
					Dim j As UInt32
					处理学校百分比(学校统计信息.ElementAt(i).Value)
					If 学校统计信息详细.ContainsKey(学校统计信息.ElementAt(i).Key) Then
						Dim gt As Int32 = 0
						Dim gc As Int32 = 0
						For j = 0 To 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息.Count - 1
							If Not 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j) Is Nothing Then
								gt += 1
							End If
						Next
						SelectSchoolTemplate(学校统计信息.ElementAt(i).Key)
						For j = 0 To 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息.Count - 1
							If Not 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j) Is Nothing Then
								logR("正在处理 " & gradeNameTbl(j))
								处理学校百分比(学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j))
								生成单个年级报告(学校统计信息.ElementAt(i).Key, j, gradeNameTbl(j), 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j), 学校统计信息.ElementAt(i).Value)
								logR(String.Format("文件 {0}/{1}，学校 {2}/{3}，年级 {4}/{5}，处理完", 第几个文件, 共几个文件,
								 i + 1, 学校统计信息.Count,
								 gc + 1, gt))
								gc += 1
							End If
							If wkExiting Then GoTo out
						Next
					End If
				ElseIf 生成何种数据 = OpType.ClassReport Then
					Dim j As UInt32
					If 学校统计信息详细.ContainsKey(学校统计信息.ElementAt(i).Key) Then
						Dim gt As Int32 = 0
						Dim gc As Int32 = 0
						For j = 0 To 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息.Count - 1
							If Not 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j) Is Nothing Then
								gt += 1
							End If
						Next
						SelectSchoolTemplate(学校统计信息.ElementAt(i).Key)
						For j = 0 To 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息.Count - 1
							If Not 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j) Is Nothing Then
								logR("正在处理 " & gradeNameTbl(j))
								处理学校百分比(学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j))
								Dim k As Int32
								For k = 0 To 学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).Count - 1
									logR("正在处理 " & 学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).ElementAt(k).Key)
									处理学校百分比(学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).ElementAt(k).Value)
									生成单个班级报告(学校统计信息.ElementAt(i).Key, gradeNameTbl(j), 学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).ElementAt(k).Key, 学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).ElementAt(k).Value, 学校统计信息详细(学校统计信息.ElementAt(i).Key).年级统计信息(j))
									logR(String.Format("文件 {0}/{1}，学校 {2}/{3}，年级 {4}/{5}，班级 {6}/{7}，处理完",
									 第几个文件, 共几个文件, i + 1, 学校统计信息.Count,
									 gc + 1, gt,
									 k + 1, 学校统计信息详细(学校统计信息.ElementAt(i).Key).班级信息详细(j).Count))
									If wkExiting Then GoTo out
								Next
								logR(String.Format("文件 {0}/{1}，学校 {2}/{3}，年级 {4}/{5}，处理完", 第几个文件, 共几个文件,
								 i + 1, 学校统计信息.Count,
								 gc + 1, gt))
								gc += 1
							End If
						Next
					End If
				End If
				logR(String.Format("文件 {0}/{1}，学校 {2}/{3}，处理完", 第几个文件, 共几个文件, i + 1, 学校统计信息.Count))
				purgeAsync()
				If wkExiting Then Exit For
				'Exit For
			Next
		Catch e As Exception
			logE("处理数据:" & e.Message)
			logE(e.StackTrace)
			GoTo out
		End Try

out:
		测试结果分析综述字体 = Nothing
		测试结果分析文本字体 = Nothing
		内容中的百分比字体 = Nothing
	End Sub

	Private Sub 处理区百分比(ByVal 共几个文件 As UInt32, ByVal 第几个文件 As UInt32, ByRef 待处理文件 As String, ByRef excelWb As Excel.Workbook)
		Dim i As Int32
		Dim x(3) As UInt32
		Dim y(3) As UInt32

		For i = 1 To excelWb.Sheets.Count
			excelWs = excelWb.Sheets(i)

			当前行号 = 1
			已经读取的行数 = 0
			预取数据到缓存(excelWs)

			Do While True
				移动到下一行()
				预取数据到缓存(excelWs)

				If 获取当前行数据(3) = String.Empty Then Exit Do
				If Not IsNumeric(获取当前行数据(3)) Then Continue Do

				sendProgress(String.Format("文件 {0}/{1}，工作表 {2}/{3}，行{4}开始 ...", 第几个文件, 共几个文件, i, excelWb.Sheets.Count, 当前行号))

				For j = 0 To 3
					x(j) = 获取当前行数据(4 + j * 2)
				Next
				计算百分比(x, y)
				For j = 0 To 3
					excelWs.Cells(当前行号, 5 + j * 2).Value2 = 转换完整百分比(y(j))
				Next

				sendProgress(String.Format("文件 {0}/{1}，工作表 {2}/{3}，行{4}结束 ...", 第几个文件, 共几个文件, i, excelWb.Sheets.Count, 当前行号))
			Loop
		Next
		excelWb.Close(Excel.XlSaveAction.xlSaveChanges)
		excelWb = Nothing
	End Sub

	Private Sub 处理数据(ByVal 共几个文件 As UInt32, ByVal 第几个文件 As UInt32, ByRef 待处理文件 As String)
		Dim 生成何种数据 As Int32 = -1

		logI("开始 - 处理数据 " & 待处理文件)

		Try
			' 学校报告
			If wkType = OpType.SchoolReport Or wkType = OpType.GradeReport Or wkType = OpType.ClassReport Then
				excelWb = excelApp.Workbooks.Open(待处理文件, Nothing, False)
			Else
				excelWb = excelApp.Workbooks.Open(待处理文件, Nothing, True)
			End If
			excelWs = Nothing
			For i = 1 To excelWb.Sheets.Count
				excelWs = excelWb.Sheets(i)
				If wkType = OpType.StudentScore And excelWs.Range("A1").Text = "学校名称" Then
					生成何种数据 = OpType.StudentScore
					Exit For
				End If
				If wkType = OpType.StudentReport And excelWs.Range("C1").Text = "ID" Then
					生成何种数据 = OpType.StudentReport
					Exit For
				End If
				If wkType = OpType.SchoolReport And excelWs.Range("C1").Text = "ID" Then
					生成何种数据 = OpType.SchoolReport
					Exit For
				End If
				If wkType = OpType.GradeReport And excelWs.Range("C1").Text = "ID" Then
					生成何种数据 = OpType.GradeReport
					Exit For
				End If
				If wkType = OpType.ClassReport And excelWs.Range("C1").Text = "ID" Then
					生成何种数据 = OpType.ClassReport
					Exit For
				End If
				If wkType = OpType.AreaPercent And excelWs.Range("C1").Text = "样本个数" Then
					生成何种数据 = OpType.AreaPercent
					Exit For
				End If
			Next
		Catch e As Exception
			logE(e.Message)
			logE(e.StackTrace)
			GoTo out
		End Try

		If 生成何种数据 = -1 Then
			logE("没有找到可用的工作表: " & 待处理文件)
			GoTo out
		End If

		'If useClipboard = 0 Then
		If tmpName = String.Empty Then
			tmpName = System.IO.Path.GetTempFileName()
			logR("获取临时文件: " & tmpName)
		End If
		'End If

		' 处理原始数据
		If 生成何种数据 = OpType.StudentScore Then
			Dim excelWbDst As Excel.Workbook
			Try
				excelWbDst = excelApp.Workbooks.Add()
				excelWbDst.Application.DisplayAlerts = False
				excelWbDst.SaveAs(Strings.Replace(待处理文件, ".xls", "-数据.xls"))
				excelWbDst.Application.DisplayAlerts = True
			Catch e As Exception
				logE("不能打开输出文件: " & e.Message)
				logE("不能打开输出文件: " & e.StackTrace)
				GoTo out
			End Try
			Try
				' Entry
				start(共几个文件, 第几个文件, excelWs, excelWbDst.Sheets(1))
			Catch e As Exception
				logE("处理Excel数据: " & e.Message)
				logE("处理Excel数据: " & e.StackTrace)
				excelWbDst.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
				excelWbDst = Nothing
				GoTo out
			End Try
			excelWbDst.Close(Excel.XlSaveAction.xlSaveChanges)
			excelWbDst = Nothing

			GoTo out
		End If

		' 学校/年级/班级报告
		If 生成何种数据 = OpType.SchoolReport Or 生成何种数据 = OpType.GradeReport Or 生成何种数据 = OpType.ClassReport Then
			' 第一次处理的时候打开Excel模板
			If excelWbTmpl Is Nothing Then
				Try
					If 生成何种数据 = OpType.SchoolReport Then
						excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\学校报告模板.xlsx")
					ElseIf 生成何种数据 = OpType.GradeReport Then
						excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\年级报告模板.xlsx")
					ElseIf 生成何种数据 = OpType.ClassReport Then
						excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\班级报告模板.xlsx")
					Else
						excelWbTmpl = Nothing
					End If
					excelWsTmpl = excelWbTmpl.Sheets(1)
				Catch e As Exception
					logE("打开Excel模板: " & e.Message)
					logE(e.StackTrace)
					'MsgBox("打开Excel模板: " & e.Message)
					GoTo out
				End Try
			End If

			Try
				' Entry
				生成学校报告(生成何种数据, 共几个文件, 第几个文件, 待处理文件, excelWs)
			Catch e As Exception
				logE("处理Excel数据: " & e.Message)
				logE("处理Excel数据: " & e.StackTrace)
			End Try

			GoTo out
		End If

		' 处理区百分比
		If 生成何种数据 = OpType.AreaPercent Then
			Try
				处理区百分比(共几个文件, 第几个文件, 待处理文件, excelWb)
			Catch e As Exception
				logE("处理Excel数据: " & e.Message)
				logE("处理Excel数据: " & e.StackTrace)
			End Try
			GoTo out
		End If

		' 处理Excel，生成报表

		列名转列号表.Clear()
		当前行号 = 1
		已经读取的行数 = 0
		预取数据到缓存(excelWs)
		生成列信息表格(列重命名1)

		If excelWbTmpl Is Nothing Then
			Try
				excelWbTmpl = excelApp.Workbooks.Add(Application.StartupPath & "\Tmpl.xlsx")
			Catch e As Exception
				logE("打开Excel模板: " & e.Message)
				logE(e.StackTrace)
				'MsgBox("打开Excel模板: " & e.Message)
				GoTo out
			End Try
		End If

		测项起始列号 = 0
		测项附加分起始列号 = 0

		If 列名转列号表.ContainsKey("50米跑成绩") Then
			测项起始列号 = 列名转列号表("50米跑成绩")
		End If
		If 列名转列号表.ContainsKey("是否有50米跑") Then
			测项附加分起始列号 = 列名转列号表("是否有50米跑")
		End If

		If 测项起始列号 = 0 Or 测项附加分起始列号 = 0 Then
			logE("测项起始列号 " & 测项起始列号 & " 测项附加分起始列号 " & 测项附加分起始列号)
		End If

		Try
			计算学校整体情况()

			Do While True
				移动到下一行()
				预取数据到缓存(excelWs)

				If 获取当前行数据("姓名") = String.Empty Then
					sendProgress(String.Format("共{0}个文件。当前处理第{1}个文件的第{2}行。处理完毕。", 共几个文件, 第几个文件, 当前行号))
					Exit Do
				End If

				logR("当前行: " & 当前行号 & " 姓名 " & 获取当前行数据("姓名") & " 年级: " & 获取当前行数据("年级") & " 性别: " & 获取当前行数据("性别"))

				计算类别()

				打开报告()

				生成报告()

				关闭报告()

				sendProgress(String.Format("共{0}个文件。当前处理第{1}个文件的第{2}行", 共几个文件, 第几个文件, 当前行号))

				purgeAsync()

				If wkExiting Then Exit Do
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
			excelWb.Close(Excel.XlSaveAction.xlDoNotSaveChanges)
			excelWb = Nothing
		End If
		If tmpName <> String.Empty Then
			Try
				If File.Exists(tmpName) Then
					File.Delete(tmpName)
					logR("删除临时文件: " & tmpName)
				End If
			Catch ex As Exception
				logR("删除临时文件失败: " & tmpName)
			End Try
			tmpName = String.Empty
		End If
		logI("结束 - 处理数据 " & 待处理文件)
	End Sub

	Private Sub 生成()
		logI("开始 - 生成")

		'logR("待处理文件列表: " & String.Join(",", 待处理文件列表))
		If 待处理文件列表 Is Nothing Or 待处理文件列表.Length = 0 Then
			logR("没有待处理文件列表")
			GoTo out
		End If

		'Exit Sub

		装载应用()

		logFlush()

		' 循环处理所有Excel
		Try
			For i = 0 To 待处理文件列表.Length - 1
				logR("开始 - 处理: " & 待处理文件列表(i))
				' 处理一个Excel
				处理数据(待处理文件列表.Length, i + 1, 待处理文件列表(i))
				logR("结束 - 处理: " & 待处理文件列表(i))
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
		logI("结束 - 生成")
		logFlush()
	End Sub

	Private Sub Worker()
		Dim type As Int32 = wkType
		Try
			If type = OpType.AllReport Then
				wkType = OpType.SchoolReport
				生成()
				wkType = OpType.GradeReport
				生成()
				wkType = OpType.ClassReport
				生成()
			Else
				生成()
			End If
		Catch ex As Exception
			logE(ex.Message)
			logE(ex.StackTrace)
		End Try

		Thread.VolatileWrite(wkDone, 1)
	End Sub

	Private Sub 读配置文件()
		Dim sr As StreamReader = Nothing
		Dim s As String
		Dim m As Int32

		转pdf = 0
		useClipboard = 1
		maxNumOfAdvise = 6
		列重命名0.Clear()
		列重命名1.Clear()
		学校转学区表.Clear()

		Try
			sr = New StreamReader("vReport.txt", True)

			m = 1
			While True
				s = sr.ReadLine()
				If s Is Nothing Then Exit While
				Dim a As String() = s.Split(" ")
				If a.Length = 0 Then
					m = 0
					Continue While
				End If
				If a(0) = "区" Then
					m = 1
					Continue While
				End If
				If a(0) = "列重命名0" Then
					m = 2
					Continue While
				End If
				If a(0) = "列重命名1" Then
					m = 3
					Continue While
				End If
				If a(0) = "转pdf" Then
					转pdf = 1
					Continue While
				End If
				If a(0) = "maxNumOfAdvise" Then
					m = 0
					If a.Length = 2 Then
						maxNumOfAdvise = Int(a(1))
						If maxNumOfAdvise < 1 Or maxNumOfAdvise > 256 Then
							maxNumOfAdvise = 6
						End If
					End If
					Continue While
				End If
				If a(0) = "noClipboard" Then
					m = 0
					useClipboard = 0
					Continue While
				End If

				If m = 1 And a.Length = 2 Then
					If Not 学校转学区表.ContainsKey(a(1)) Then 学校转学区表(a(1)) = a(0)
				End If
				If m = 2 And a.Length = 2 Then
					If Not 列重命名0.ContainsKey(a(0)) Then 列重命名0(a(0)) = a(1)
				End If
				If m = 3 And a.Length = 2 Then
					If Not 列重命名1.ContainsKey(a(0)) Then 列重命名1(a(0)) = a(1)
				End If
			End While
		Catch ex As Exception
			logE("打开配置文件vReport.txt失败")
		Finally
			If Not sr Is Nothing Then sr.Close()
		End Try
	End Sub

	Private Sub 点击事件(ByVal 类别 As Int32)
		Dim i As UInt32

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

		If 待处理文件列表 Is Nothing Or 待处理文件列表.Length = 0 Then
			logR("没有待处理文件列表")
			GoTo out
		End If

		For i = 0 To 待处理文件列表.Length - 1
			logR("第" & (i + 1) & "个待处理文件 " & 待处理文件列表(i))
		Next

		读配置文件()

		'Label1.Text = ""
		ToolStripStatusLabel1.Text = ""

		Thread.VolatileWrite(wkExiting, 0)
		Thread.VolatileWrite(wkDone, 0)
		wk = New Thread(AddressOf Worker)
		'MessageBox.Show(wk.GetApartmentState())
		wk.SetApartmentState(ApartmentState.STA)
		wkType = 类别
		wk.Start()
		Thread.VolatileWrite(wkStart, 1)

out:
		logI("结束 - 处理点击事件")
	End Sub

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
		点击事件(0)
	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
		cancelReport(0)
	End Sub

	Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
		点击事件(1)
	End Sub

	Function 打开报告()
		Dim docFullName As String
		Dim docPath As String

		'logW("开始 - 打开报告")

		docPath = Application.StartupPath & "\" & 获取当前行数据("学校") & "\" & 获取当前行数据("年级") & "\" & 获取当前行数据("班级")
		If Not Directory.Exists(docPath) Then
			docPath = Application.StartupPath & "\" & 获取当前行数据("学校")
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
			docPath = docPath & "\" & 获取当前行数据("年级")
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
			docPath = docPath & "\" & 获取当前行数据("班级")
			If Not Directory.Exists(docPath) Then Directory.CreateDirectory(docPath)
		End If
		docFullName = docPath & "\" _
		 & 获取当前行数据("学校") & "_" & 获取当前行数据("年级") & "_" & 获取当前行数据("班级") & "_" & 获取当前行数据("ID") & "_" & 获取当前行数据("姓名") & ".docx"

		logW("打开报告模板")
		wordDoc = wordApp.Documents.Add(Application.StartupPath & "\Tmpl.docx")
		logW("保存报告: " & docFullName)
		wordDoc.SaveAs(docFullName)
		If displayWord Then wordDoc.Application.Activate()

		'logW("结束 - 打开报告")

		打开报告 = 0
	End Function

	Function 关闭报告()
		'logW("开始 - 关闭报告")

		If 转pdf = 1 Then
			Try
				Dim docFullName As String
				Dim docPath As String
				docPath = Application.StartupPath & "\" & 获取当前行数据("学校") & "\" & 获取当前行数据("年级") & "\" & 获取当前行数据("班级")
				docFullName = docPath & "\" _
				 & 获取当前行数据("学校") & "_" & 获取当前行数据("年级") & "_" & 获取当前行数据("班级") & "_" & 获取当前行数据("ID") & "_" & 获取当前行数据("姓名") & ".pdf"

				wordDoc.SaveAs(docFullName, Word.WdSaveFormat.wdFormatPDF)
			Catch
			End Try
		End If
		wordDoc.Close(Word.WdSaveOptions.wdSaveChanges)
		wordDoc = Nothing

		'logW("结束 - 关闭报告")

		关闭报告 = 0
	End Function

	Function 生成首页()
		Dim token As String

		'logI("开始 - 生成首页")

		' 学校名称
		学校名称 = 获取当前行数据("学校")
		' 年级班级
		年级班级 = 获取当前行数据("年级") & " " & 获取当前行数据("班级")
		' 学生姓名
		学生姓名 = 获取当前行数据("姓名")

		For i = 1 To wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs.Count
			If i > wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs.Count Then
				logE("名称不能属于同一段落")
				Exit For
			End If
			token = wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text
			If token.Length >= 2 Then
				If token.Contains("XXMC") Then
					' 学校名称
					wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text = 学校名称
				ElseIf token.Contains("NJBJ") Then
					' 年级班级
					wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text = 年级班级
				ElseIf token.Contains("NJ") Then
					' 年级
					wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text = 获取当前行数据("年级")
				ElseIf token.Contains("BJ") Then
					' 班级
					wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text = 获取当前行数据("班级")
				ElseIf token.Contains("XSXM") Then
					' 学生姓名
					wordDoc.Shapes(1).TextFrame.TextRange.Paragraphs(i).Range.Text = 学生姓名
				Else
				End If
			End If
		Next

		'logR("学校名称 " & 学校名称 & " 年级班级 " & 年级班级 & " 学生姓名 " & 学生姓名)

		'logI("结束 - 生成首页")

		生成首页 = 0
	End Function

	Function 生成学生情况()
		'logI("开始 - 生成学生情况")

		wordDoc.Tables(1).Cell(1, 2).Range.Text = 获取当前行数据("姓名")
		wordDoc.Tables(1).Cell(1, 4).Range.Text = 获取当前行数据("ID")
		wordDoc.Tables(1).Cell(1, 6).Range.Text = 获取当前行数据("性别")
		wordDoc.Tables(1).Cell(2, 2).Range.Text = 获取当前行数据("年级")
		wordDoc.Tables(1).Cell(2, 4).Range.Text = 获取当前行数据("班级")
		wordDoc.Tables(1).Cell(2, 6).Range.Text = 获取当前行数据("测试成绩")
		wordDoc.Tables(1).Cell(3, 2).Range.Text = 获取当前行数据("测试成绩评定")
		wordDoc.Tables(1).Cell(3, 4).Range.Text = 获取当前行数据("综合成绩")
		wordDoc.Tables(1).Cell(3, 6).Range.Text = 获取当前行数据("综合评定")
		wordDoc.Tables(1).Cell(4, 2).Range.Text = 获取当前行数据("学校")

		'logI("结束 - 生成学生情况")

		生成学生情况 = 0
	End Function

	Function 生成单项指标()
		Dim 表格位置 As UInt32 = 2
		Dim 测项序号 As UInt32 = 0
		Dim 表格行 As UInt32 = 0
		Dim 内容 As String
		Dim idx As UInt32 = 0
		Dim i As UInt32 = 0
		Dim j As UInt32 = 0

		'logI("开始 - 生成单项指标")

		' 身体形态
		内容 = 获取当前行数据("身高体重指数")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 2).Range.Text = 内容
		内容 = 获取当前行数据("身高体重成绩")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 3).Range.Text = 内容
		内容 = 获取当前行数据("身高体重等级")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(2, 4).Range.Text = 内容
		wordDoc.Tables(表格位置).Cell(2, 5).Range.Text = "/"
		' 身体机能
		内容 = 获取当前行数据("肺活量成绩")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 2).Range.Text = 内容
		内容 = 获取当前行数据("肺活量得分")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 3).Range.Text = 内容
		内容 = 获取当前行数据("肺活量等级")
		If 内容 = "X" Then 内容 = ""
		wordDoc.Tables(表格位置).Cell(3, 4).Range.Text = 内容
		wordDoc.Tables(表格位置).Cell(3, 5).Range.Text = "/"

		' 动态项
		idx = 当前类别 * 6
		For i = 1 To 学生测试项信息(idx)
			'' 在加分指标前面加一行
			'If i <> 学生测试项信息(idx) Then
			'	wordDoc.Tables(表格位置).Rows.Add(wordDoc.Tables(表格位置).Rows(wordDoc.Tables(表格位置).Rows.Count - 2))
			'End If
			wordDoc.Tables(表格位置).Rows.Add()
			测项序号 = 学生测试项信息(idx + i)

			内容 = 测项名称信息(测项序号)
			'logW(i & " " & idx & " 测项 " & 测项序号 & " " & 内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 1).Range.Text = 内容

			内容 = 获取当前行数据(测项起始列号 + 3 * 测项序号 + 0)
			If 内容 = "X" Then
				内容 = ""
			Else
				' to '
				内容 = 内容.Replace(Convert.ToChar(8216), Convert.ToChar(39))
				内容 = 内容.Replace(Convert.ToChar(8242), Convert.ToChar(39))
			End If
			'Dim a As Int32
			'Dim b As Int32
			'a = Convert.ToInt32(内容(1))
			'b = Convert.ToInt32("'"(0))
			'logW(内容)
			If i = 1 Then
				' 50米跑留一位小数
				If Not 内容.Contains(".") Then
					内容 = 内容 & ".0"
				End If
			End If
			wordDoc.Tables(表格位置).Cell(3 + i, 2).Range.Text = 内容

			内容 = 获取当前行数据(测项起始列号 + 3 * 测项序号 + 1)
			If 内容 = "X" Then 内容 = ""
			'logW(内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 3).Range.Text = 内容

			内容 = 获取当前行数据(测项起始列号 + 3 * 测项序号 + 2)
			If 内容 = "X" Then 内容 = ""
			'logW(内容)
			wordDoc.Tables(表格位置).Cell(3 + i, 4).Range.Text = 内容

			内容 = 获取当前行数据(测项附加分起始列号 + 2 * 测项序号)
			If 学生测试项附加分信息(idx + i) <> 0 Then
				内容 = 获取当前行数据(测项附加分起始列号 + 2 * 测项序号 + 1)
				If 内容 <> "" And 内容 <> "0" Then
					wordDoc.Tables(表格位置).Cell(3 + i, 5).Range.Text = 内容
				Else
					wordDoc.Tables(表格位置).Cell(3 + i, 5).Range.Text = "0"
				End If
			Else
				wordDoc.Tables(表格位置).Cell(3 + i, 5).Range.Text = "/"
			End If
		Next

		'logI("结束 - 生成单项指标")

		生成单项指标 = 0
	End Function

	Function 生成各指标得分图表()
		Dim 表格位置 As UInt32 = 2
		Dim 图表工作表 As Excel.Worksheet
		Dim 测项序号 As UInt32 = 0
		Dim idx As UInt32 = 0
		Dim i As UInt32

		'logI("开始 - 生成各指标得分图表")

		Try
			图表工作表 = excelWbTmpl.Sheets(工作表名称信息(当前类别) & "图表")
			图表工作表.Activate()
			图表工作表.Cells(1, 2).Value2 = 获取当前行数据("身高体重成绩")
			图表工作表.Cells(2, 2).Value2 = 获取当前行数据("肺活量得分")

			' 动态项
			idx = 当前类别 * 6
			For i = 1 To 学生测试项信息(idx)
				测项序号 = 学生测试项信息(idx + i)
				图表工作表.Cells(2 + i, 2).Value2 = 获取当前行数据(测项起始列号 + 3 * 测项序号 + 1)
			Next

			'If useClipboard = 0 Then
			'	Dim ct As Excel.Chart = 图表工作表.ChartObjects(1).Chart
			'	ct.Export(tmpName, "GIF")
			'Else
			'	图表工作表.Shapes.SelectAll()
			'	System.Windows.Forms.Clipboard.Clear()
			'	图表工作表.Application.Selection.copy()
			'End If
		Catch e As Exception
			logE("生成各指标得分图表:" & e.Message)
			logE(e.StackTrace)
			'MsgBox("生成各指标得分图表:" & e.Message)
			GoTo out
		End Try

		wordDoc.Tables(表格位置).Select()
		wordDoc.Application.Selection.MoveDown()
		' 多塞个空行
		wordDoc.Application.Selection.TypeParagraph()

		Dim retry As Int32
		Dim suc As Int32
		suc = 0
		For retry = 1 To 20
			Try
				图表工作表.Activate()
				If useClipboard = 0 Then
					Dim ct As Excel.Chart = 图表工作表.ChartObjects(1).Chart
					ct.Export(tmpName, "GIF")
					wordDoc.Application.Selection.InlineShapes.AddPicture(tmpName, False, True)
				Else
					图表工作表.Shapes.SelectAll()
					System.Windows.Forms.Clipboard.Clear()
					图表工作表.Application.Selection.copy()
					wordDoc.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdChartPicture)
				End If
				suc = 1
			Catch ex As Exception
				If retry = 20 Then
					Throw ex
				End If
				logW("重试 " & retry)
				logW(ex.Message)
				logW(ex.StackTrace)
			End Try
			If suc = 1 Then
				Exit For
			End If
		Next

out:
		'logI("结束 - 生成各指标得分图表")

		生成各指标得分图表 = 0
	End Function

	' 10000 倍
	Function 转换百分比(ByVal 数值 As UInt32)
		If 数值 = 0 Then
			转换百分比 = "0.00"
		ElseIf 数值 < 10 Then
			转换百分比 = "0.0" & 数值
		ElseIf 数值 < 100 Then
			转换百分比 = "0." & 数值
		ElseIf 数值 = 10000 Then
			转换百分比 = "100.00"
		Else
			Dim 余数 As UInt32
			余数 = 数值 Mod 100
			If 余数 < 10 Then
				转换百分比 = Int(数值 / 100) & ".0" & 余数
			Else
				转换百分比 = Int(数值 / 100) & "." & 余数
			End If
		End If
		'logI("转换百分比: " & 数值 & " > " & 转换百分比)
	End Function

	Function 转换完整百分比(ByVal 数值 As UInt32)
		转换完整百分比 = 转换百分比(数值) & "%"
	End Function

	Function 格式化百分比(ByVal 百分比 As String)
		Dim idx As Int32

		If 百分比 = String.Empty Then
			百分比 = "0"
		End If

		idx = Strings.InStr(百分比, ".")
		If idx = 0 Then
			格式化百分比 = 百分比 & ".00"
		Else
			idx = 百分比.Length - idx
			If idx = 0 Then
				格式化百分比 = 百分比 & "00"
			ElseIf idx = 1 Then
				格式化百分比 = 百分比 & "0"
			Else
				格式化百分比 = 百分比
			End If
		End If
	End Function

	Function 生成学校整体情况()
		'logI("开始 - 生成学校整体情况")

		Dim 表格位置 As UInt32 = 3
		Dim col As UInt32
		Dim i As UInt32 = 0

		'logW("count " & wordDoc.Tables.Count & " " & 表格位置)
		If wordDoc.Tables.Count < 表格位置 Then
			logR("没有找到学校整体情况")
			GoTo out
		End If

		' last col (0) -> Name (1) -> Count(2) -> Percent(3)
		col = rptHdrTbl.Length + 3

		For i = 0 To 3
			wordDoc.Tables(表格位置).Cell(i + 2, 2).Range.Text = 格式化百分比(excelWs.Cells(i + 2, col).Value2)
			wordDoc.Tables(表格位置).Cell(i + 2, 4).Range.Text = 格式化百分比(excelWs.Cells(i + 2, col + 3).Value2)
			wordDoc.Tables(表格位置).Cell(i + 2, 6).Range.Text = 格式化百分比(excelWs.Cells(i + 2, col + 6).Value2)
		Next

out:
		'logI("结束- 生成学校整体情况")

		生成学校整体情况 = 0
	End Function

	Function 生成运动处方()
		'Dim content As String
		Dim idx As UInt32
		Dim i As UInt32
		Dim j As UInt32
		Dim k As UInt32

		'logI("开始 - 生成运动处方")

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
		等级 = 计算等级(获取当前行数据("综合评定"))
		'logW("等级 " & 等级)
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（一）" & excelWsTmpl.Range("A1").Text)
		wordDoc.Application.Selection.TypeParagraph()
		wordDoc.Application.Selection.Style = "正文1"
		wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (1 + 1 + 等级 * 3)).Text)
		wordDoc.Application.Selection.TypeParagraph()

		Select Case 获取当前行数据("身高体重等级")
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
		等级 = 计算等级(获取当前行数据("肺活量等级"))
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
			wordDoc.Application.Selection.TypeText(i + 1 & ". " & excelWsTmpl.Range("A" & (40 + 1 + i * 3)).Text)
			wordDoc.Application.Selection.TypeParagraph()
			wordDoc.Application.Selection.Style = "正文1"
			wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("E" & (40 + 1 + i * 3)).Text)
			wordDoc.Application.Selection.TypeParagraph()
		Next

		Randomize()

		'学生身体素质测试结果的建议
		wordDoc.Application.Selection.Style = "主标题1"
		wordDoc.Application.Selection.TypeText("（五）" & excelWsTmpl.Range("A62").Text)
		wordDoc.Application.Selection.TypeParagraph()
		For i = 0 To 学生测试项信息(当前类别 * 6) - 1
			' 第i个测项
			' 4个等级, 4个小等级，每个小等级3行
			idx = 62 + 1 + i * 4 * 4 * 3
			Dim 测项列号 As UInt32
			测项列号 = 测项起始列号 + 学生测试项信息(当前类别 * 6 + 1 + i) * 3 + 2
			等级 = 计算等级(获取当前行数据(测项列号))
			idx = idx + 等级 * 4 * 3 + 身高体重等级 * 3

			logR("建议" & i & " 测项列号 " & 测项列号 & " 等级 " & 获取当前行数据(测项列号) & " " & 等级)

			' 需要加粗
			wordDoc.Application.Selection.Style = "主标题2"
			wordDoc.Application.Selection.TypeText(excelWsTmpl.Range("A" & idx).Text)
			'wordDoc.Application.Selection.Paragraphs.First.
			wordDoc.Application.Selection.TypeParagraph()
			wordDoc.Application.Selection.Style = "正文1"

			Dim max As Int32 = 0

			For j = 0 To 128
				If excelWsTmpl.Cells(idx, 5 + j).Text = String.Empty Then Exit For
				max += 1
			Next

			If max > maxNumOfAdvise Then
				Dim tmp(maxNumOfAdvise - 1) As Int32
				For j = 0 To maxNumOfAdvise - 1
					tmp(j) = Int(Rnd() * (max - j))
				Next
				For j = 1 To maxNumOfAdvise - 1
					For k = 0 To j - 1
						If tmp(j) >= tmp(k) Then tmp(j) += 1
					Next
				Next
				'logR(String.Format("随机项 {0} {1} {2}", tmp(0), tmp(1), tmp(2)))
				For j = 0 To maxNumOfAdvise - 1
					wordDoc.Application.Selection.TypeText(j + 1 & ". " & excelWsTmpl.Cells(idx, 5 + tmp(j)).Text)
					wordDoc.Application.Selection.TypeParagraph()
				Next
			Else
				For j = 0 To max - 1
					wordDoc.Application.Selection.TypeText(j + 1 & ". " & excelWsTmpl.Cells(idx, 5 + j).Text)
					wordDoc.Application.Selection.TypeParagraph()
				Next
			End If
		Next

		'logI("结束 - 生成运动处方")

		生成运动处方 = 0
	End Function

	Function 生成报告()
		'logI("开始 - 生成报告")

		生成首页()

		生成学生情况()

		生成单项指标()

		生成各指标得分图表()

		生成学校整体情况()

		生成运动处方()

		'logI("结束 - 生成报告")

		生成报告 = 0
	End Function

	Public Sub New()
		Dim i As UInt32

		st = New Student()
		ReDim st.arr(最大缓存列数 - 1)

		For i = 0 To 最大缓存行数 - 1
			ReDim 数据缓存(i)(最大缓存列数 - 1)
		Next

		转pdf = 0

		列重命名0 = New Dictionary(Of String, String)
		列重命名1 = New Dictionary(Of String, String)
		列重命名2 = New Dictionary(Of String, String)
		学校转学区表 = New Dictionary(Of String, String)
		列名转列号表 = New Dictionary(Of String, UInt32)

		'Randomize(0)

		' 此调用是 Windows 窗体设计器所必需的。
		InitializeComponent()

		' 在 InitializeComponent() 调用之后添加任何初始化。

	End Sub

	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub

	' table header
	Private rptHdrTbl() As String = {
	"是否参测", "缺项数量", "ID", "所属区", "姓名", "学校", "学段", "年级", "班级", "性别",
	"综合成绩", "综合评定", "测试成绩", "测试成绩评定", "附加分",
	"身高成绩", "体重成绩", "身高体重指数", "身高体重成绩", "身高体重等级",
	"肺活量成绩", "肺活量得分", "肺活量等级",
	"50米跑成绩", "50米跑得分", "50米跑等级",
	"坐位体前屈成绩", "坐位体前屈得分", "坐位体前屈等级",
	"一分钟跳绳成绩", "一分钟跳绳得分", "一分钟跳绳等级",
	"一分钟仰卧起坐成绩", "一分钟仰卧起坐得分", "一分钟仰卧起坐等级",
	"50米×8往返跑成绩", "50米×8往返跑得分", "50米×8往返跑等级",
	"立定跳远成绩", "立定跳远得分", "立定跳远等级",
	"800米跑成绩", "800米跑得分", "800米跑等级",
	"1000米跑成绩", "1000米跑得分", "1000米跑等级",
	"引体向上成绩", "引体向上得分", "引体向上等级",
	"是否有50米跑", "50米跑附加分",
	"是否有坐位体前屈", "坐位体前屈附加分",
	"是否有一分钟跳绳", "一分钟跳绳附加分",
	"是否有一分钟仰卧起坐", "一分钟仰卧起坐附加分",
	"是否有50米×8往返跑", "50米×8往返跑附加分",
	"是否有立定跳远", "立定跳远附加分",
	"是否有800米跑", "800米跑附加分",
	"是否有1000米跑", "1000米跑附加分",
	"是否有引体向上", "引体向上附加分",
	"赋分项",
	"预留1", "预留2", "预留3", "预留4", "预留5", "预留6", "预留7", "预留8"}

	Private gradeNameTbl() As String = { _
	"一年级", "二年级", "三年级", "四年级", "五年级", "六年级", _
	"初一", "初二", "初三", _
	"高一", "高二", "高三", _
	"大一", "大二", "大三", "大四"}

	Private BMIData() As Int32 = { _
	  135, 137, 139, 142, 144, 147, 155, 157, 158, 165, 168, 173, 179, _
	  134, 136, 138, 141, 143, 146, 154, 156, 157, 164, 167, 172, 178, _
	  182, 185, 195, 202, 215, 219, 222, 226, 229, 233, 238, 239, 240, _
	  204, 205, 222, 227, 242, 246, 250, 253, 261, 264, 266, 274, 280, _
	  133, 135, 136, 137, 138, 142, 148, 153, 160, 165, 169, 171, 172, _
	  132, 134, 135, 136, 137, 141, 147, 152, 159, 164, 168, 170, 171, _
	  174, 179, 187, 195, 206, 209, 218, 223, 227, 228, 233, 234, 240, _
	  193, 203, 212, 221, 230, 237, 245, 249, 252, 253, 255, 258, 280}

	Private fhlData0() As Int32 = { _
	  100, 1700, 2000, 2300, 2600, 2900, 3200, 3640, 3940, 4240, 4540, 4740, 4940, 5040, 5140, _
	  95, 1600, 1900, 2200, 2500, 2800, 3100, 3520, 3820, 4120, 4420, 4620, 4820, 4920, 5020, _
	  90, 1500, 1800, 2100, 2400, 2700, 3000, 3400, 3700, 4000, 4300, 4500, 4700, 4800, 4900, _
	  85, 1400, 1650, 1900, 2150, 2450, 2750, 3150, 3450, 3750, 4050, 4250, 4450, 4550, 4650, _
	  80, 1300, 1500, 1700, 1900, 2200, 2500, 2900, 3200, 3500, 3800, 4000, 4200, 4300, 4400, _
	  78, 1240, 1430, 1620, 1820, 2110, 2400, 2780, 3080, 3380, 3680, 3880, 4080, 4180, 4280, _
	  76, 1180, 1360, 1540, 1740, 2020, 2300, 2660, 2960, 3260, 3560, 3760, 3960, 4060, 4160, _
	  74, 1120, 1290, 1460, 1660, 1930, 2200, 2540, 2840, 3140, 3440, 3640, 3840, 3940, 4040, _
	  72, 1060, 1220, 1380, 1580, 1840, 2100, 2420, 2720, 3020, 3320, 3520, 3720, 3820, 3920, _
	  70, 1000, 1150, 1300, 1500, 1750, 2000, 2300, 2600, 2900, 3200, 3400, 3600, 3700, 3800, _
	  68, 940, 1080, 1220, 1420, 1660, 1900, 2180, 2480, 2780, 3080, 3280, 3480, 3580, 3680, _
	  66, 880, 1010, 1140, 1340, 1570, 1800, 2060, 2360, 2660, 2960, 3160, 3360, 3460, 3560, _
	  64, 820, 940, 1060, 1260, 1480, 1700, 1940, 2240, 2540, 2840, 3040, 3240, 3340, 3440, _
	  62, 760, 870, 980, 1180, 1390, 1600, 1820, 2120, 2420, 2720, 2920, 3120, 3220, 3320, _
	  60, 700, 800, 900, 1100, 1300, 1500, 1700, 2000, 2300, 2600, 2800, 3000, 3100, 3200, _
	  50, 660, 750, 840, 1030, 1220, 1410, 1600, 1890, 2180, 2470, 2660, 2850, 2940, 3030, _
	  40, 620, 700, 780, 960, 1140, 1320, 1500, 1780, 2060, 2340, 2520, 2700, 2780, 2860, _
	  30, 580, 650, 720, 890, 1060, 1230, 1400, 1670, 1940, 2210, 2380, 2550, 2620, 2690, _
	  20, 540, 600, 660, 820, 980, 1140, 1300, 1560, 1820, 2080, 2240, 2400, 2460, 2520, _
	  10, 500, 550, 600, 750, 900, 1050, 1200, 1450, 1700, 1950, 2100, 2250, 2300, 2350}
	Private fhlData1() As Int32 = { _
	  100, 1400, 1600, 1800, 2000, 2250, 2500, 2750, 2900, 3050, 3150, 3250, 3350, 3400, 3450, _
	  95, 1300, 1500, 1700, 1900, 2150, 2400, 2650, 2850, 3000, 3100, 3200, 3300, 3350, 3400, _
	  90, 1200, 1400, 1600, 1800, 2050, 2300, 2550, 2800, 2950, 3050, 3150, 3250, 3300, 3350, _
	  85, 1100, 1300, 1500, 1700, 1950, 2200, 2450, 2650, 2800, 2900, 3000, 3100, 3150, 3200, _
	  80, 1000, 1200, 1400, 1600, 1850, 2100, 2350, 2500, 2650, 2750, 2850, 2950, 3000, 3050, _
	  78, 960, 1150, 1340, 1530, 1770, 2010, 2250, 2400, 2550, 2650, 2750, 2850, 2900, 2950, _
	  76, 920, 1100, 1280, 1460, 1690, 1920, 2150, 2300, 2450, 2550, 2650, 2750, 2800, 2850, _
	  74, 880, 1050, 1220, 1390, 1610, 1830, 2050, 2200, 2350, 2450, 2550, 2650, 2700, 2750, _
	  72, 840, 1000, 1160, 1320, 1530, 1740, 1950, 2100, 2250, 2350, 2450, 2550, 2600, 2650, _
	  70, 800, 950, 1100, 1250, 1450, 1650, 1850, 2000, 2150, 2250, 2350, 2450, 2500, 2550, _
	  68, 760, 900, 1040, 1180, 1370, 1560, 1750, 1900, 2050, 2150, 2250, 2350, 2400, 2450, _
	  66, 720, 850, 980, 1110, 1290, 1470, 1650, 1800, 1950, 2050, 2150, 2250, 2300, 2350, _
	  64, 680, 800, 920, 1040, 1210, 1380, 1550, 1700, 1850, 1950, 2050, 2150, 2200, 2250, _
	  62, 640, 750, 860, 970, 1130, 1290, 1450, 1600, 1750, 1850, 1950, 2050, 2100, 2150, _
	  60, 600, 700, 800, 900, 1050, 1200, 1350, 1500, 1650, 1750, 1850, 1950, 2000, 2050, _
	  50, 580, 680, 780, 880, 1020, 1170, 1310, 1460, 1610, 1710, 1810, 1910, 1960, 2010, _
	  40, 560, 660, 760, 860, 990, 1140, 1270, 1420, 1570, 1670, 1770, 1870, 1920, 1970, _
	  30, 540, 640, 740, 840, 960, 1110, 1230, 1380, 1530, 1630, 1730, 1830, 1880, 1930, _
	  20, 520, 620, 720, 820, 930, 1080, 1190, 1340, 1490, 1590, 1690, 1790, 1840, 1890, _
	  10, 500, 600, 700, 800, 900, 1050, 1150, 1300, 1450, 1550, 1650, 1750, 1800, 1850}

	Private M50Data0() As Int32 = { _
	  100, 102, 96, 91, 87, 84, 82, 78, 75, 73, 71, 70, 68, 67, 66, _
	  95, 103, 97, 92, 88, 85, 83, 79, 76, 74, 72, 71, 69, 68, 67, _
	  90, 104, 98, 93, 89, 86, 84, 80, 77, 75, 73, 72, 70, 69, 68, _
	  85, 105, 99, 94, 90, 87, 85, 81, 78, 76, 74, 73, 71, 70, 69, _
	  80, 106, 100, 95, 91, 88, 86, 82, 79, 77, 75, 74, 72, 71, 70, _
	  78, 108, 102, 97, 93, 90, 88, 84, 81, 79, 77, 76, 74, 73, 72, _
	  76, 110, 104, 99, 95, 92, 90, 86, 83, 81, 79, 78, 76, 75, 74, _
	  74, 112, 106, 101, 97, 94, 92, 88, 85, 83, 81, 80, 78, 77, 76, _
	  72, 114, 108, 103, 99, 96, 94, 90, 87, 85, 83, 82, 80, 79, 78, _
	  70, 116, 110, 105, 101, 98, 96, 92, 89, 87, 85, 84, 82, 81, 80, _
	  68, 118, 112, 107, 103, 100, 98, 94, 91, 89, 87, 86, 84, 83, 82, _
	  66, 120, 114, 109, 105, 102, 100, 96, 93, 91, 89, 88, 86, 85, 84, _
	  64, 122, 116, 111, 107, 104, 102, 98, 95, 93, 91, 90, 88, 87, 86, _
	  62, 124, 118, 113, 109, 106, 104, 100, 97, 95, 93, 92, 90, 89, 88, _
	  60, 126, 120, 115, 111, 108, 106, 102, 99, 97, 95, 94, 92, 91, 90, _
	  50, 128, 122, 117, 113, 110, 108, 104, 101, 99, 97, 96, 94, 93, 92, _
	  40, 130, 124, 119, 115, 112, 110, 106, 103, 101, 99, 98, 96, 95, 94, _
	  30, 132, 126, 121, 117, 114, 112, 108, 105, 103, 101, 100, 98, 97, 96, _
	  20, 134, 128, 123, 119, 116, 114, 110, 107, 105, 103, 102, 100, 99, 98, _
	  10, 136, 130, 125, 121, 118, 116, 112, 109, 107, 105, 104, 102, 101, 100}
	Private M50Data1() As Int32 = { _
	  100, 110, 100, 92, 87, 83, 82, 81, 80, 79, 78, 77, 76, 75, 74, _
	  95, 111, 101, 93, 88, 84, 83, 82, 81, 80, 79, 78, 77, 76, 75, _
	  90, 112, 102, 94, 89, 85, 84, 83, 82, 81, 80, 79, 78, 77, 76, _
	  85, 115, 105, 97, 92, 88, 87, 86, 85, 84, 83, 82, 81, 80, 79, _
	  80, 118, 108, 100, 95, 91, 90, 89, 88, 87, 86, 85, 84, 83, 82, _
	  78, 120, 110, 102, 97, 93, 92, 91, 90, 89, 88, 87, 86, 85, 84, _
	  76, 122, 112, 104, 99, 95, 94, 93, 92, 91, 90, 89, 88, 87, 86, _
	  74, 124, 114, 106, 101, 97, 96, 95, 94, 93, 92, 91, 90, 89, 88, _
	  72, 126, 116, 108, 103, 99, 98, 97, 96, 95, 94, 93, 92, 91, 90, _
	  70, 128, 118, 110, 105, 101, 100, 99, 98, 97, 96, 95, 94, 93, 92, _
	  68, 130, 120, 112, 107, 103, 102, 101, 100, 99, 98, 97, 96, 95, 94, _
	  66, 132, 122, 114, 109, 105, 104, 103, 102, 101, 100, 99, 98, 97, 96, _
	  64, 134, 124, 116, 111, 107, 106, 105, 104, 103, 102, 101, 100, 99, 98, _
	  62, 136, 126, 118, 113, 109, 108, 107, 106, 105, 104, 103, 102, 101, 100, _
	  60, 138, 128, 120, 115, 111, 110, 109, 108, 107, 106, 105, 104, 103, 102, _
	  50, 140, 130, 122, 117, 113, 112, 111, 110, 109, 108, 107, 106, 105, 104, _
	  40, 142, 132, 124, 119, 115, 114, 113, 112, 111, 110, 109, 108, 107, 106, _
	  30, 144, 134, 126, 121, 117, 116, 115, 114, 113, 112, 111, 110, 109, 108, _
	  20, 146, 136, 128, 123, 119, 118, 117, 116, 115, 114, 113, 112, 111, 110, _
	  10, 148, 138, 130, 125, 121, 120, 119, 118, 117, 116, 115, 114, 113, 112}

	Private tsData0() As Int32 = { _
	  100, 109, 117, 126, 137, 148, 157, _
	  95, 104, 112, 121, 132, 143, 152, _
	  90, 99, 107, 116, 127, 138, 147, _
	  85, 93, 101, 110, 121, 132, 141, _
	  80, 87, 95, 104, 115, 126, 135, _
	  78, 80, 88, 97, 108, 119, 128, _
	  76, 73, 81, 90, 101, 112, 121, _
	  74, 66, 74, 83, 94, 105, 114, _
	  72, 59, 67, 76, 87, 98, 107, _
	  70, 52, 60, 69, 80, 91, 100, _
	  68, 45, 53, 62, 73, 84, 93, _
	  66, 38, 46, 55, 66, 77, 86, _
	  64, 31, 39, 48, 59, 70, 79, _
	  62, 24, 32, 41, 52, 63, 72, _
	  60, 17, 25, 34, 45, 56, 65, _
	  50, 14, 22, 31, 42, 53, 62, _
	  40, 11, 19, 28, 39, 50, 59, _
	  30, 8, 16, 25, 36, 47, 56, _
	  20, 5, 13, 22, 33, 44, 53, _
	  10, 2, 10, 19, 30, 41, 50}
	Private tsData1() As Int32 = { _
	  100, 117, 127, 139, 149, 158, 166, _
	  95, 110, 120, 132, 142, 151, 159, _
	  90, 103, 113, 125, 135, 144, 152, _
	  85, 95, 105, 117, 127, 136, 144, _
	  80, 87, 97, 109, 119, 128, 136, _
	  78, 80, 90, 102, 112, 121, 129, _
	  76, 73, 83, 95, 105, 114, 122, _
	  74, 66, 76, 88, 98, 107, 115, _
	  72, 59, 69, 81, 91, 100, 108, _
	  70, 52, 62, 74, 84, 93, 101, _
	  68, 45, 55, 67, 77, 86, 94, _
	  66, 38, 48, 60, 70, 79, 87, _
	  64, 31, 41, 53, 63, 72, 80, _
	  62, 24, 34, 46, 56, 65, 73, _
	  60, 17, 27, 39, 49, 58, 66, _
	  50, 14, 24, 36, 46, 55, 63, _
	  40, 11, 21, 33, 43, 52, 60, _
	  30, 8, 18, 30, 40, 49, 57, _
	  20, 5, 15, 27, 37, 46, 54, _
	  10, 2, 12, 24, 34, 43, 51}

	Private tyData0() As Int32 = { _
	  100, 225, 240, 250, 260, 265, 270, 273, 275, _
	  95, 218, 233, 245, 255, 260, 265, 268, 270, _
	  90, 211, 226, 240, 250, 255, 260, 263, 265, _
	  85, 203, 218, 233, 243, 248, 253, 256, 258, _
	  80, 195, 210, 225, 235, 240, 245, 248, 250, _
	  78, 191, 206, 221, 231, 236, 241, 244, 246, _
	  76, 187, 202, 217, 227, 232, 237, 240, 242, _
	  74, 183, 198, 213, 223, 228, 233, 236, 238, _
	  72, 179, 194, 209, 219, 224, 229, 232, 234, _
	  70, 175, 190, 205, 215, 220, 225, 228, 230, _
	  68, 171, 186, 201, 211, 216, 221, 224, 226, _
	  66, 167, 182, 197, 207, 212, 217, 220, 222, _
	  64, 163, 178, 193, 203, 208, 213, 216, 218, _
	  62, 159, 174, 189, 199, 204, 209, 212, 214, _
	  60, 155, 170, 185, 195, 200, 205, 208, 210, _
	  50, 150, 165, 180, 190, 195, 200, 203, 205, _
	  40, 145, 160, 175, 185, 190, 195, 198, 200, _
	  30, 140, 155, 170, 180, 185, 190, 193, 195, _
	  20, 135, 150, 165, 175, 180, 185, 188, 190, _
	  10, 130, 145, 160, 170, 175, 180, 183, 185}
	Private tyData1() As Int32 = { _
	  100, 196, 200, 202, 204, 205, 206, 207, 208, _
	  95, 190, 194, 196, 198, 199, 200, 201, 202, _
	  90, 184, 188, 190, 192, 193, 194, 195, 196, _
	  85, 177, 181, 183, 185, 186, 187, 188, 189, _
	  80, 170, 174, 176, 178, 179, 180, 181, 182, _
	  78, 167, 171, 173, 175, 176, 177, 178, 179, _
	  76, 164, 168, 170, 172, 173, 174, 175, 176, _
	  74, 161, 165, 167, 169, 170, 171, 172, 173, _
	  72, 158, 162, 164, 166, 167, 168, 169, 170, _
	  70, 155, 159, 161, 163, 164, 165, 166, 167, _
	  68, 152, 156, 158, 160, 161, 162, 163, 164, _
	  66, 149, 153, 155, 157, 158, 159, 160, 161, _
	  64, 146, 150, 152, 154, 155, 156, 157, 158, _
	  62, 143, 147, 149, 151, 152, 153, 154, 155, _
	  60, 140, 144, 146, 148, 149, 150, 151, 152, _
	  50, 135, 139, 141, 143, 144, 145, 146, 147, _
	  40, 130, 134, 136, 138, 139, 140, 141, 142, _
	  30, 125, 129, 131, 133, 134, 135, 136, 137, _
	  20, 120, 124, 126, 128, 129, 130, 131, 132, _
	  10, 115, 119, 121, 123, 124, 125, 126, 127}

	Private zwtData0() As Int32 = { _
	  100, 161, 162, 163, 164, 165, 166, 176, 196, 216, 236, 243, 246, 249, 251, _
	  95, 146, 147, 149, 150, 152, 153, 159, 177, 197, 215, 224, 228, 231, 233, _
	  90, 130, 132, 134, 136, 138, 140, 142, 158, 178, 194, 205, 210, 213, 215, _
	  85, 120, 119, 118, 117, 116, 115, 123, 137, 158, 172, 183, 191, 195, 199, _
	  80, 110, 106, 102, 98, 94, 90, 104, 116, 138, 150, 161, 172, 177, 182, _
	  78, 99, 95, 91, 86, 82, 77, 91, 103, 124, 136, 147, 158, 163, 168, _
	  76, 88, 84, 80, 74, 70, 64, 78, 90, 110, 122, 133, 144, 149, 154, _
	  74, 77, 73, 69, 62, 58, 51, 65, 77, 96, 108, 119, 130, 135, 140, _
	  72, 66, 62, 58, 50, 46, 38, 52, 64, 82, 94, 105, 116, 121, 126, _
	  70, 55, 51, 47, 38, 34, 25, 39, 51, 68, 80, 91, 102, 107, 112, _
	  68, 44, 40, 36, 26, 22, 12, 26, 38, 54, 66, 77, 88, 93, 98, _
	  66, 33, 29, 25, 14, 10, -1, 13, 25, 40, 52, 63, 74, 79, 84, _
	  64, 22, 18, 14, 2, -2, -14, 0, 12, 26, 38, 49, 60, 65, 70, _
	  62, 11, 7, 3, -10, -14, -27, -13, -1, 12, 24, 35, 46, 51, 56, _
	  60, 0, -4, -8, -22, -26, -40, -26, -14, -2, 10, 21, 32, 37, 42, _
	  50, -8, -12, -16, -32, -36, -50, -38, -26, -14, 0, 11, 22, 27, 32, _
	  40, -16, -20, -24, -42, -46, -60, -50, -38, -26, -10, 1, 12, 17, 22, _
	  30, -24, -28, -32, -52, -56, -70, -62, -50, -38, -20, -9, 2, 7, 12, _
	  20, -32, -36, -40, -62, -66, -80, -74, -62, -50, -30, -19, -8, -3, 2, _
	  10, -40, -44, -48, -72, -76, -90, -86, -74, -62, -40, -29, -18, -13, -8}
	Private zwtData1() As Int32 = { _
	  100, 186, 189, 192, 195, 198, 199, 218, 227, 235, 242, 248, 253, 258, 263, _
	  95, 173, 176, 179, 181, 185, 187, 201, 210, 218, 225, 231, 236, 240, 244, _
	  90, 160, 163, 166, 169, 172, 175, 184, 193, 201, 208, 214, 219, 222, 224, _
	  85, 147, 148, 149, 150, 151, 152, 167, 176, 184, 191, 197, 202, 206, 210, _
	  80, 134, 133, 132, 131, 130, 129, 150, 159, 167, 174, 180, 185, 190, 195, _
	  78, 123, 122, 121, 120, 119, 118, 137, 146, 154, 161, 167, 172, 177, 182, _
	  76, 112, 111, 110, 109, 108, 107, 124, 133, 141, 148, 154, 159, 164, 169, _
	  74, 101, 100, 99, 98, 97, 96, 111, 120, 128, 135, 141, 146, 151, 156, _
	  72, 90, 89, 88, 87, 86, 85, 98, 107, 115, 122, 128, 133, 138, 143, _
	  70, 79, 78, 77, 76, 75, 74, 85, 94, 102, 109, 115, 120, 125, 130, _
	  68, 68, 67, 66, 65, 64, 63, 72, 81, 89, 96, 102, 107, 112, 117, _
	  66, 57, 56, 55, 54, 53, 52, 59, 68, 76, 83, 89, 94, 99, 104, _
	  64, 46, 45, 44, 43, 42, 41, 46, 55, 63, 70, 76, 81, 86, 91, _
	  62, 35, 34, 33, 32, 31, 30, 33, 42, 50, 57, 63, 68, 73, 78, _
	  60, 24, 23, 22, 21, 20, 19, 20, 29, 37, 44, 50, 55, 60, 65, _
	  50, 16, 15, 14, 13, 12, 11, 12, 21, 29, 36, 42, 47, 52, 57, _
	  40, 8, 7, 6, 5, 4, 3, 4, 13, 21, 28, 34, 39, 44, 49, _
	  30, 0, -1, -2, -3, -4, -5, -4, 5, 13, 20, 26, 31, 36, 41, _
	  20, -8, -9, -10, -11, -12, -13, -12, -3, 5, 12, 18, 23, 28, 33, _
	  10, -16, -17, -18, -19, -20, -21, -20, -11, -3, 4, 10, 15, 20, 25}

	Private ywqzData0() As Int32 = { _
	  100, 48, 49, 50, 51, 13, 14, 15, 16, 17, 18, 19, 20, _
	  95, 45, 46, 47, 48, 12, 13, 14, 15, 16, 17, 18, 19, _
	  90, 42, 43, 44, 45, 11, 12, 13, 14, 15, 16, 17, 18, _
	  85, 39, 40, 41, 42, 10, 11, 12, 13, 14, 15, 16, 17, _
	  80, 36, 37, 38, 39, 9, 10, 11, 12, 13, 14, 15, 16, _
	  78, 34, 35, 36, 37, 9, 10, 11, 12, 13, 14, 15, 16, _
	  76, 32, 33, 34, 35, 8, 9, 10, 11, 12, 13, 14, 15, _
	  74, 30, 31, 32, 33, 8, 9, 10, 11, 12, 13, 14, 15, _
	  72, 28, 29, 30, 31, 7, 8, 9, 10, 11, 12, 13, 14, _
	  70, 26, 27, 28, 29, 7, 8, 9, 10, 11, 12, 13, 14, _
	  68, 24, 25, 26, 27, 6, 7, 8, 9, 10, 11, 12, 13, _
	  66, 22, 23, 24, 25, 6, 7, 8, 9, 10, 11, 12, 13, _
	  64, 20, 21, 22, 23, 5, 6, 7, 8, 9, 10, 11, 12, _
	  62, 18, 19, 20, 21, 5, 6, 7, 8, 9, 10, 11, 12, _
	  60, 16, 17, 18, 19, 4, 5, 6, 7, 8, 9, 10, 11, _
	  50, 14, 15, 16, 17, 3, 4, 5, 6, 7, 8, 9, 10, _
	  40, 12, 13, 14, 15, 2, 3, 4, 5, 6, 7, 8, 9, _
	  30, 10, 11, 12, 13, 1, 2, 3, 4, 5, 6, 7, 8, _
	  20, 8, 9, 10, 11, 1, 1, 2, 3, 4, 5, 6, 7, _
	  10, 6, 7, 8, 9, 1, 1, 1, 2, 3, 4, 5, 6}
	Private ywqzData1() As Int32 = { _
	  100, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, _
	  95, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, _
	  90, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, _
	  85, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, _
	  80, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, _
	  78, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, _
	  76, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, _
	  74, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, _
	  72, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, _
	  70, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, _
	  68, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, _
	  66, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, _
	  64, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, _
	  62, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, _
	  60, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, _
	  50, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, _
	  40, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, _
	  30, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, _
	  20, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, _
	  10, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17}

	Private nlpData0() As Int32 = { _
	  100, 96, 90, 235, 230, 220, 210, 205, 200, 197, 195, _
	  95, 99, 93, 245, 235, 225, 215, 210, 205, 202, 200, _
	  90, 102, 96, 255, 240, 230, 220, 215, 210, 207, 205, _
	  85, 105, 99, 262, 247, 237, 227, 222, 217, 214, 212, _
	  80, 108, 102, 270, 255, 245, 235, 230, 225, 222, 220, _
	  78, 111, 105, 275, 260, 250, 240, 235, 230, 227, 225, _
	  76, 114, 108, 280, 265, 255, 245, 240, 235, 232, 230, _
	  74, 117, 111, 285, 270, 260, 250, 245, 240, 237, 235, _
	  72, 120, 114, 290, 275, 265, 255, 250, 245, 242, 240, _
	  70, 123, 117, 295, 280, 270, 260, 255, 250, 247, 245, _
	  68, 126, 120, 300, 285, 275, 265, 260, 255, 252, 250, _
	  66, 129, 123, 305, 290, 280, 270, 265, 260, 257, 255, _
	  64, 132, 126, 310, 295, 285, 275, 270, 265, 262, 260, _
	  62, 135, 129, 315, 300, 290, 280, 275, 270, 267, 265, _
	  60, 138, 132, 320, 305, 295, 285, 280, 275, 272, 270, _
	  50, 142, 136, 340, 325, 315, 305, 300, 295, 292, 290, _
	  40, 146, 140, 360, 345, 335, 325, 320, 315, 312, 310, _
	  30, 150, 144, 380, 365, 355, 345, 340, 335, 332, 330, _
	  20, 154, 148, 400, 385, 375, 365, 360, 355, 352, 350, _
	  10, 158, 152, 420, 405, 395, 385, 380, 375, 372, 370}
	Private nlpData1() As Int32 = { _
	  100, 101, 97, 215, 210, 205, 204, 202, 200, 198, 196, _
	  95, 104, 100, 222, 217, 212, 210, 208, 206, 204, 202, _
	  90, 107, 103, 229, 224, 219, 216, 214, 212, 210, 208, _
	  85, 110, 106, 237, 232, 227, 223, 221, 219, 217, 215, _
	  80, 113, 109, 245, 240, 235, 230, 228, 226, 224, 222, _
	  78, 116, 112, 250, 245, 240, 235, 233, 231, 229, 227, _
	  76, 119, 115, 255, 250, 245, 240, 238, 236, 234, 232, _
	  74, 122, 118, 260, 255, 250, 245, 243, 241, 239, 237, _
	  72, 125, 121, 265, 260, 255, 250, 248, 246, 244, 242, _
	  70, 128, 124, 270, 265, 260, 255, 253, 251, 249, 247, _
	  68, 131, 127, 275, 270, 265, 260, 258, 256, 254, 252, _
	  66, 134, 130, 280, 275, 270, 265, 263, 261, 259, 257, _
	  64, 137, 133, 285, 280, 275, 270, 268, 266, 264, 262, _
	  62, 140, 136, 290, 285, 280, 275, 273, 271, 269, 267, _
	  60, 143, 139, 295, 290, 285, 280, 278, 276, 274, 272, _
	  50, 147, 143, 305, 300, 295, 290, 288, 286, 284, 282, _
	  40, 151, 147, 315, 310, 305, 300, 298, 296, 294, 292, _
	  30, 155, 151, 325, 320, 315, 310, 308, 306, 304, 302, _
	  20, 159, 155, 335, 330, 325, 320, 318, 316, 314, 312, _
	  10, 163, 159, 345, 340, 335, 330, 328, 326, 324, 322}

	Sub 预取数据到缓存(ByRef excelWsSrc As Excel.Worksheet)
		If 已经读取的行数 < 当前行号 Then
			Dim obj(,) As Object

			obj = excelWsSrc.Range(excelWsSrc.Cells(当前行号, 1), excelWsSrc.Cells(当前行号 + 最大缓存行数 - 1, 最大缓存列数)).Value2

			Dim i As UInt32
			Dim j As UInt32
			For i = 0 To 最大缓存行数 - 1
				For j = 0 To 最大缓存列数 - 1
					If obj(i + 1, j + 1) Is Nothing Then
						数据缓存(i)(j) = String.Empty
					Else
						数据缓存(i)(j) = obj(i + 1, j + 1).ToString()
					End If
				Next
			Next

			已经读取的行数 += 最大缓存行数
		End If
	End Sub

	Function 获取当前行数据(ByRef 列名 As String) As String
		Dim 行号 As UInt32
		获取当前行数据 = String.Empty
		If 列名转列号表.ContainsKey(列名) Then
			行号 = ((当前行号 - 1) Mod 最大缓存行数)
			获取当前行数据 = 数据缓存(行号)(列名转列号表(列名) - 1)
		End If
	End Function

	Function 获取当前行数据(ByVal 列号 As Long) As String
		Dim 行号 As UInt32
		行号 = ((当前行号 - 1) Mod 最大缓存行数)
		获取当前行数据 = 数据缓存(行号)(列号 - 1)
	End Function

	Sub 移动到下一行()
		当前行号 += 1
	End Sub

	Sub 生成列信息表格(ByRef 列重命名 As Dictionary(Of String, String))
		Dim i As UInt32
		For i = 0 To 最大缓存列数 - 1
			If 数据缓存(0)(i) <> String.Empty Then
				Dim 映射后的列名 As String = 数据缓存(0)(i)
				If 列重命名.ContainsKey(数据缓存(0)(i)) Then
					映射后的列名 = 列重命名(映射后的列名)
					logI("列重命名 " & 数据缓存(0)(i) & " " & 映射后的列名)
				End If
				If Not 列名转列号表.ContainsKey(映射后的列名) Then 列名转列号表(映射后的列名) = i + 1
			End If
		Next
	End Sub

	Sub start(ByVal 共几个文件 As UInt32, ByVal 第几个文件 As UInt32, ByRef excelWsSrc As Excel.Worksheet, ByRef excelWsDst As Excel.Worksheet)
		' 重置学校整体情况
		For i = 0 To 3
			各等级计数(i) = 0
			各身体形态计数(i) = 0
			各身体机能计数(i) = 0
		Next

		列名转列号表.Clear()
		当前行号 = 1
		已经读取的行数 = 0
		预取数据到缓存(excelWs)
		生成列信息表格(列重命名0)

		logW("创建表头")

		createStudentReportHeader(excelWsDst)
		'Exit Sub

		logW("开始生成")

		Do While True
			移动到下一行()
			预取数据到缓存(excelWs)

			' end of data
			If 获取当前行数据("姓名") = String.Empty Then
				logR(String.Format("生成第{0}行。结束。", 当前行号))
				Exit Do
			End If

			' invalid grade
			If Not 获取当前行数据("年级编号") < 100 Then GoTo rowComplete
			If Not 获取当前行数据("年级编号") > 10 Then GoTo rowComplete

			logR(String.Format("生成第{0}行", 当前行号))

			initStudent(st)

			readStudent(excelWsSrc, 当前行号, st)

			calcStudentScore(st)

			createStudentReport(excelWsSrc, 当前行号, st, excelWsDst)

			sendProgress(String.Format("共{0}个文件。当前处理第{1}个文件的第{2}行", 共几个文件, 第几个文件, 当前行号))

			purgeAsync()

			If wkExiting Then Exit Do
rowComplete:
		Loop

		createSchoolOveral(excelWsDst)

		logW("生成结束")
	End Sub

	Sub initStudent(ByRef st As Student)
		st.idStr = ""
		st.schoolStr = ""
		st.classStr = ""
		st.nameStr = ""
		st.genderStr = ""
		st.gradeStr = ""
		st.heightStr = ""
		st.weightStr = ""
		st.fhlStr = ""
		st.m50Str = ""
		st.zwtStr = ""
		st.tsStr = ""
		st.tyStr = ""
		st.ywqz0Str = ""
		st.ywqz1Str = ""
		st.nlp0Str = ""
		st.nlp1Str = ""

		st.gender = 0
		st.grade = 0

		st.height = 0
		st.weight = 0
		st.bmiValid = 0
		st.bmi = 0
		st.bmiLow = 0
		st.bmiScore = 0

		st.fhlValid = 0
		st.fhl = 0
		st.fhlScore = 0

		st.m50Valid = 0
		st.m50 = 0
		st.m50Score = 0

		st.zwtValid = 0
		st.zwt = 0
		st.zwtScore = 0

		st.tsValid = 0
		st.ts = 0
		st.tsScore = 0
		st.tsJfScore = 0

		st.tyValid = 0
		st.ty = 0
		st.tyScore = 0
		st.tyJfScore = 0

		st.ywqzValid = 0
		st.ywqz0 = 0
		st.ywqz1 = 0
		st.ywqzScore = 0
		st.ywqzJfScore = 0

		st.nlpValid = 0
		st.nlp0 = 0
		st.nlp1 = 0
		st.nlpScore = 0
		st.nlpJfScore = 0

		st.totalValid = 0
		st.totalScore = 0
		st.totalJfScore = 0

		st.extraScore = 0

		'For i = 0 To UBound(rptHdrTbl)
		'st.arr(i) = ""
		'Next
	End Sub

	Function timeToSeconds(ByVal t As String)
		Dim num As Long
		Dim tmp As Long
		Dim i As Long

		timeToSeconds = 0

		num = 0
		tmp = Len(t)
		For i = 1 To tmp
			If Not IsNumeric(Mid(t, i, 1)) Then Exit For
			num = num + 1
		Next i

		If num <> 0 Then
			timeToSeconds = Int(Val(Strings.Left(t, num))) * 60 + Int(Val(Mid(t, num + 2)))
		End If
	End Function

	Function validateInput(ByVal val As Long, ByVal str As String)
		validateInput = 0
		If val <> 0 Then
			validateInput = 1
			Exit Function
		End If
		If Len(str) > 0 Then
			If str(0) = "0" Then validateInput = 1
		End If
	End Function

	Sub readStudent(ByRef excelWsSrc As Excel.Worksheet, ByVal row As Long, ByRef st As Student)
		If 获取当前行数据("年级编号") < 20 Then
			st.idStr = 获取当前行数据("学籍号")
			st.schoolStr = 获取当前行数据("学校名称")
			st.classStr = 获取当前行数据("班级名称")
			st.nameStr = 获取当前行数据("姓名")
			st.genderStr = 获取当前行数据("性别")
			st.gender = 获取当前行数据("性别")
			st.gradeStr = 获取当前行数据("年级编号")
			st.heightStr = 获取当前行数据("身高")
			st.weightStr = 获取当前行数据("体重")
			st.fhlStr = 获取当前行数据("肺活量")
			st.m50Str = 获取当前行数据("50米")
			st.zwtStr = 获取当前行数据("坐位体前屈")
			st.tsStr = 获取当前行数据("一分钟跳绳")
			st.tyStr = 获取当前行数据("一分钟跳绳")
			st.ywqz0Str = 获取当前行数据("仰卧起坐")
			st.ywqz1Str = 获取当前行数据("仰卧起坐")
			st.nlp0Str = 获取当前行数据("50米*8往返跑")
			st.nlp1Str = 获取当前行数据("50米*8往返跑")
			'st.idStr = excelWsSrc.Range("E" & row).Text
			'st.schoolStr = excelWsSrc.Range("A" & row).Text
			'st.classStr = excelWsSrc.Range("D" & row).Text
			'st.nameStr = excelWsSrc.Range("G" & row).Text
			'st.genderStr = excelWsSrc.Range("H" & row).Text
			'st.gender = excelWsSrc.Range("H" & row).Value2
			'st.gradeStr = excelWsSrc.Range("B" & row).Text
			'st.heightStr = excelWsSrc.Range("M" & row).Text
			'st.weightStr = excelWsSrc.Range("N" & row).Text
			'st.fhlStr = excelWsSrc.Range("O" & row).Text
			'st.m50Str = excelWsSrc.Range("P" & row).Text
			'st.zwtStr = excelWsSrc.Range("Q" & row).Text
			'st.tsStr = excelWsSrc.Range("R" & row).Text
			'st.tyStr = excelWsSrc.Range("R" & row).Text
			'st.ywqz0Str = excelWsSrc.Range("S" & row).Text
			'st.ywqz1Str = excelWsSrc.Range("S" & row).Text
			'st.nlp0Str = excelWsSrc.Range("T" & row).Text
			'st.nlp1Str = excelWsSrc.Range("T" & row).Text
		Else
			st.idStr = 获取当前行数据("学籍号")
			st.schoolStr = 获取当前行数据("学校名称")
			st.classStr = 获取当前行数据("班级名称")
			st.nameStr = 获取当前行数据("姓名")
			st.genderStr = 获取当前行数据("性别")
			st.gender = 获取当前行数据("性别")
			st.gradeStr = 获取当前行数据("年级编号")
			st.heightStr = 获取当前行数据("身高")
			st.weightStr = 获取当前行数据("体重")
			st.fhlStr = 获取当前行数据("肺活量")
			st.m50Str = 获取当前行数据("50米")
			st.zwtStr = 获取当前行数据("坐位体前屈")
			' ty
			st.tsStr = 获取当前行数据("立定跳远")
			st.tyStr = 获取当前行数据("立定跳远")
			If 获取当前行数据("仰卧起坐") <> String.Empty Then
				st.ywqz0Str = 获取当前行数据("仰卧起坐")
				'st.ywqz1Str = Range("u" & row)
			Else
				st.ywqz0Str = 获取当前行数据("引体向上")
				'st.ywqz1Str = Range("V" & row)
			End If
			If 获取当前行数据("800米") <> String.Empty Then
				st.nlp0Str = 获取当前行数据("800米")
				'st.nlp1Str = Range("S" & row)
			Else
				st.nlp0Str = 获取当前行数据("1000米")
				'st.nlp1Str = Range("T" & row)
			End If
			'st.idStr = excelWsSrc.Range("E" & row).Text
			'st.schoolStr = excelWsSrc.Range("A" & row).Text
			'st.classStr = excelWsSrc.Range("D" & row).Text
			'st.nameStr = excelWsSrc.Range("G" & row).Text
			'st.genderStr = excelWsSrc.Range("H" & row).Text
			'st.gender = excelWsSrc.Range("H" & row).Value2
			'st.gradeStr = excelWsSrc.Range("B" & row).Text
			'st.heightStr = excelWsSrc.Range("M" & row).Text
			'st.weightStr = excelWsSrc.Range("N" & row).Text
			'st.fhlStr = excelWsSrc.Range("O" & row).Text
			'st.m50Str = excelWsSrc.Range("P" & row).Text
			'st.zwtStr = excelWsSrc.Range("Q" & row).Text
			'' ty
			'st.tsStr = excelWsSrc.Range("R" & row).Text
			'st.tyStr = excelWsSrc.Range("R" & row).Text
			'If excelWsSrc.Range("U" & row).Text <> "" Then
			'	st.ywqz0Str = excelWsSrc.Range("U" & row).Text
			'	'st.ywqz1Str = Range("u" & row)
			'Else
			'	st.ywqz0Str = excelWsSrc.Range("v" & row).Text
			'	'st.ywqz1Str = Range("V" & row)
			'End If
			'If excelWsSrc.Range("S" & row).Text <> "" Then
			'	st.nlp0Str = excelWsSrc.Range("S" & row).Text
			'	'st.nlp1Str = Range("S" & row)
			'Else
			'	st.nlp0Str = excelWsSrc.Range("T" & row).Text
			'	'st.nlp1Str = Range("T" & row)
			'End If
		End If

		logW(String.Format("{0} {1} {2} {3} {4}", st.nameStr, st.gender, st.schoolStr, st.gradeStr, st.classStr))

		'st.gender = Int(Val(st.genderStr))
		st.grade = calcGradeIdx(Val(st.gradeStr))

		st.height = stringToInt(st.heightStr, 1)
		st.weight = stringToInt(st.weightStr, 6)
		If st.height > 0 And st.weight > 0 Then st.bmiValid = 1

		st.fhl = Int(Val(st.fhlStr))
		If st.fhl > 0 Then st.fhlValid = 1

		Dim idx = st.m50Str.IndexOf(".")
		If idx <> -1 Then
			If idx + 2 < st.m50Str.Length Then st.m50Str = st.m50Str.Substring(0, idx + 2)
		End If
		st.m50 = stringToInt(st.m50Str, 1)
		If st.m50 > 0 Then st.m50Valid = 1

		st.zwt = stringToInt(st.zwtStr, 1)
		st.zwtValid = validateInput(st.zwt, st.zwtStr)

		st.ts = Int(Val(st.tsStr))
		If st.ts >= 0 Then st.tsValid = validateInput(st.ts, st.tsStr)
		st.ty = st.ts
		st.tyValid = st.tsValid

		st.ywqz0 = Int(Val(st.ywqz0Str))
		st.ywqz1 = st.ywqz0
		If st.ywqz0 >= 0 Then st.ywqzValid = validateInput(st.ywqz0, st.ywqz0Str)

		Dim idxOfDot As Int32 = st.nlp0Str.IndexOf(".")
		' replace with '
		If idxOfDot <> -1 Then
			st.nlp0Str = st.nlp0Str.Replace(".", QQQQ)
		Else
			idxOfDot = st.nlp0Str.IndexOf(QQQQ)
		End If
		If idxOfDot = -1 Then
			st.nlp0Str = st.nlp0Str & QQQQ & "00"
		Else
			Dim padding As Int32 = st.nlp0Str.Length - (idxOfDot + 1)
			If padding = 0 Then
				st.nlp0Str = st.nlp0Str & "00"
			ElseIf padding = 1 Then
				st.nlp0Str = st.nlp0Str & "0"
			Else
				' nothing to do
			End If
		End If
		st.nlp0 = timeToSeconds(st.nlp0Str)
		st.nlp1 = st.nlp0
		If st.nlp0 >= 0 Then st.nlpValid = validateInput(st.nlp0, st.nlp0Str)

		If st.bmiValid = 1 Or st.fhlValid = 1 Or st.m50Valid = 1 Or st.tsValid = 1 Or st.ywqzValid = 1 Or st.nlpValid = 1 Then
			st.totalValid = 1
		End If
	End Sub

	Sub calcStudentScore(ByRef st As Student)
		' 共性指标
		If st.bmiValid = 1 Then calcBMIScore(st)
		If st.fhlValid = 1 Then calcFhlScore(st)
		If st.m50Valid = 1 Then calcM50Score(st)
		If st.zwtValid = 1 Then calcZwtScore(st)

		' 跳绳或者跳远
		If st.grade < 6 Then
			If st.tsValid = 1 Then calcTsScore(st)
		Else
			If st.tyValid = 1 Then calcTyScore(st)
		End If

		' 引体向上或者仰卧起坐
		If st.grade > 1 Then
			If st.ywqzValid = 1 Then calcYwqzScore(st)
		End If

		' 耐力跑: 50x8或者1000米或者800米
		If st.grade > 3 Then
			If st.nlpValid = 1 Then calcNlpScore(st)
		End If

		If st.totalValid = 1 Then
			calcTotalScore(st)
			calcExtraScore(st)
		End If
	End Sub

	Sub createStudentReportHeader(ByRef excelWsDst As Excel.Worksheet)
		'For i = 1 To UBound(rptHdrTbl)
		'	excelWsDst.Cells(1, i).Value2 = rptHdrTbl(i)
		'Next
		excelWsDst.Range(excelWsDst.Cells(1, 1), excelWsDst.Cells(1, rptHdrTbl.Count())).Value2 = rptHdrTbl
	End Sub

	Sub createSchoolOveral(ByRef excelWsDst As Excel.Worksheet)
		Dim hdr() As String = {"综合评定等级占比", "", "", "身体形态（BMI）评定等级占比", "", "", "身体机能（肺活量）评定等级占比", "", ""}

		计算百分比(各等级计数, 各等级百分比)
		计算百分比(各身体形态计数, 各身体形态百分比)
		计算百分比(各身体机能计数, 各身体机能百分比)

		' header
		excelWsDst.Range(excelWsDst.Cells(1, rptHdrTbl.Count() + 1), excelWsDst.Cells(1, rptHdrTbl.Count() + hdr.Count())).Value2 = hdr
		'data
		For i = 0 To 3
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 2).value2 = 各等级计数(i)
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 3).value2 = 转换百分比(各等级百分比(i))
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 5).value2 = 各身体形态计数(i)
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 6).value2 = 转换百分比(各身体形态百分比(i))
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 8).value2 = 各身体机能计数(i)
			excelWsDst.Cells(2 + i, rptHdrTbl.Count() + 9).value2 = 转换百分比(各身体机能百分比(i))
		Next
	End Sub

	Sub createStudentReport(ByRef excelWsSrc As Excel.Worksheet, ByVal row As Long, ByRef st As Student, ByRef excelWsDst As Excel.Worksheet)
		Dim offset As Long
		offset = row
		Dim col As Long
		col = 0
		Dim has As Long
		Dim 是否参测 As Int32 = 0
		Dim 缺项数量 As Int32 = 0

		Dim 是否有跳绳 As Int32 = 0
		Dim 是否有仰卧起坐 As Int32 = 0
		Dim 是否有50x8 As Int32 = 0
		Dim 是否有立定跳远 As Int32 = 0
		Dim 是否有800米 As Int32 = 0
		Dim 是否有1000米 As Int32 = 0
		Dim 是否有引体向上 As Int32 = 0

		Dim m50Pos As Int32 = 0

		Dim tmp As String

		' 空两个位置放"是否参测"和"缺项数量"
		col = 2

		With st
			' ID
			.arr(col) = st.idStr
			col = col + 1
			' 所属区
			If 学校转学区表.ContainsKey(st.schoolStr) Then
				.arr(col) = 学校转学区表(st.schoolStr)
			Else
				.arr(col) = ""
			End If
			col = col + 1
			' 姓名
			.arr(col) = st.nameStr
			col = col + 1
			' 学校
			.arr(col) = st.schoolStr
			col = col + 1
			' 学段
			.arr(col) = getStageName(st.grade)
			col = col + 1
			' 年级
			.arr(col) = getGradeName(st.grade)
			col = col + 1
			' 班级
			.arr(col) = st.classStr
			col = col + 1
			' 性别
			If st.gender = 1 Then
				.arr(col) = "男"
			Else
				.arr(col) = "女"
			End If
			col = col + 1

			If st.totalValid = 1 Then
				是否参测 = 1
				tmp = getTotalLevel(st.totalScore)
				' 综合成绩
				.arr(col + 0) = Math.Round(Int((st.totalScore + st.totalJfScore + 5) / 10) / 10, 1)
				' 综合评定
				.arr(col + 1) = getTotalLevel(st.totalScore + st.totalJfScore)
				' 测试成绩
				.arr(col + 2) = Math.Round(Int((st.totalScore + 5) / 10) / 10, 1)
				' 测试评定
				.arr(col + 3) = tmp
				' 附加分
				.arr(col + 4) = st.totalJfScore / 100

				各等级计数(计算等级(tmp)) += 1
			Else
				是否参测 = 0
				.arr(col + 0) = "X"
				.arr(col + 1) = "X"
				.arr(col + 2) = "X"
				.arr(col + 3) = "X"
				.arr(col + 4) = "X"
				' 不及格
				各等级计数(3) += 1
			End If
			col = col + 5

			' 身高成绩
			.arr(col) = st.heightStr
			col = col + 1
			' 体重成绩
			.arr(col) = st.weightStr
			col = col + 1
			If st.bmiValid Then
				是否参测 = 1
				tmp = getBmiLevel(st.bmiScore, st.bmiLow)
				' 身高体重指数
				.arr(col + 0) = st.bmi / 10
				' 身高体重成绩
				.arr(col + 1) = st.bmiScore
				' 身高体重等级
				.arr(col + 2) = tmp
				各身体形态计数(计算身体形态等级(tmp)) += 1
			Else
				缺项数量 += 1
				.arr(col + 0) = "X"
				.arr(col + 1) = "X"
				.arr(col + 2) = "X"
				各身体形态计数(计算身体形态等级("未知")) += 1
			End If
			col = col + 3

			' 肺活量成绩
			.arr(col) = st.fhlStr
			col = col + 1
			If st.fhlValid = 1 Then
				是否参测 = 1
				tmp = getFhlLevel(st.fhlScore)
				' 肺活量得分
				.arr(col + 0) = st.fhlScore
				' 肺活量等级
				.arr(col + 1) = tmp
				各身体机能计数(计算等级(tmp)) += 1
			Else
				缺项数量 += 1
				.arr(col + 0) = "X"
				.arr(col + 1) = "X"
				各身体机能计数(3) += 1
			End If
			col = col + 2

			' 50米跑成绩
			m50Pos = col
			.arr(col) = st.m50Str
			col = col + 1
			If st.m50Valid = 1 Then
				是否参测 = 1
				' 50米跑得分
				.arr(col + 0) = st.m50Score
				' 50米跑等级
				.arr(col + 1) = getM50Level(st.m50Score)
			Else
				缺项数量 += 1
				.arr(col + 0) = "X"
				.arr(col + 1) = "X"
			End If
			col = col + 2

			' 坐位体前屈成绩
			.arr(col) = st.zwtStr
			col = col + 1
			If st.zwtValid = 1 Then
				是否参测 = 1
				' 坐位体前屈得分
				.arr(col + 0) = st.zwtScore
				' 坐位体前屈等级
				.arr(col + 1) = getZwtLevel(st.zwtScore)
			Else
				缺项数量 += 1
				.arr(col + 0) = "X"
				.arr(col + 1) = "X"
			End If
			col = col + 2

			If st.grade < 6 Then
				' 一分钟跳绳成绩
				.arr(col + 0) = st.tsStr
				是否有跳绳 = 1
				If st.tsValid = 1 Then
					是否参测 = 1
					' 一分钟跳绳得分
					.arr(col + 1) = st.tsScore
					' 一分钟跳绳等级
					.arr(col + 2) = getTsLevel(st.tsScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 一分钟跳绳成绩
				.arr(col + 0) = ""
				' 一分钟跳绳得分
				.arr(col + 1) = ""
				' 一分钟跳绳等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			has = 0
			If st.grade >= 2 Then
				' 小学二年级以上
				If st.grade < 6 Or st.gender <> 1 Then
					has = 1
				End If
			End If

			If has = 1 Then
				' 一分钟仰卧起坐成绩
				.arr(col + 0) = st.ywqz0Str
				是否有仰卧起坐 = 1
				If st.ywqzValid = 1 Then
					是否参测 = 1
					' 一分钟仰卧起坐得分
					.arr(col + 1) = st.ywqzScore
					' 一分钟仰卧起坐等级
					.arr(col + 2) = getYwqzLevel(st.ywqzScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 一分钟仰卧起坐成绩
				.arr(col + 0) = ""
				' 一分钟仰卧起坐得分
				.arr(col + 1) = ""
				' 一分钟仰卧起坐等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			If st.grade = 4 Or st.grade = 5 Then
				' 50米×8往返跑成绩
				.arr(col + 0) = st.nlp0Str
				是否有50x8 = 1
				If st.nlpValid = 1 Then
					是否参测 = 1
					' 50米×8往返跑得分
					.arr(col + 1) = st.nlpScore
					' 50米×8往返跑等级
					.arr(col + 2) = getNlpLevel(st.nlpScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 50米×8往返跑成绩
				.arr(col + 0) = ""
				' 50米×8往返跑得分
				.arr(col + 1) = ""
				' 50米×8往返跑等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			If st.grade >= 6 Then
				' 立定跳远成绩
				.arr(col + 0) = st.tyStr
				是否有立定跳远 = 1
				If st.tyValid = 1 Then
					是否参测 = 1
					' 立定跳远得分
					.arr(col + 1) = st.tyScore
					' 立定跳远等级
					.arr(col + 2) = getTyLevel(st.tyScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 立定跳远成绩
				.arr(col + 0) = ""
				' 立定跳远得分
				.arr(col + 1) = ""
				' 立定跳远等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			If st.grade >= 6 And st.gender <> 1 Then
				' 800米跑成绩
				.arr(col + 0) = st.nlp0Str
				是否有800米 = 1
				If st.nlpValid = 1 Then
					是否参测 = 1
					' 800米跑得分
					.arr(col + 1) = st.nlpScore
					' 800米跑等级
					.arr(col + 2) = getNlpLevel(st.nlpScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 800米跑成绩
				.arr(col + 0) = ""
				' 800米跑得分
				.arr(col + 1) = ""
				' 800米跑等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			If st.grade >= 6 And st.gender = 1 Then
				' 1000米跑成绩
				.arr(col + 0) = st.nlp0Str
				是否有1000米 = 1
				If st.nlpValid = 1 Then
					是否参测 = 1
					' 1000米跑得分
					.arr(col + 1) = st.nlpScore
					' 1000米跑等级
					.arr(col + 2) = getNlpLevel(st.nlpScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 1000米跑成绩
				.arr(col + 0) = ""
				' 1000米跑得分
				.arr(col + 1) = ""
				' 1000米跑等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			If st.grade >= 6 And st.gender = 1 Then
				' 引体向上成绩
				.arr(col + 0) = st.ywqz0Str
				是否有引体向上 = 1
				If st.ywqzValid = 1 Then
					是否参测 = 1
					' 引体向上得分
					.arr(col + 1) = st.ywqzScore
					' 引体向上等级
					.arr(col + 2) = getYwqzLevel(st.ywqzScore)
				Else
					缺项数量 += 1
					.arr(col + 1) = "X"
					.arr(col + 2) = "X"
				End If
			Else
				' 引体向上成绩
				.arr(col + 0) = ""
				' 引体向上得分
				.arr(col + 1) = ""
				' 引体向上等级
				.arr(col + 2) = ""
			End If
			col = col + 3

			' 50米跑
			.arr(col + 0) = 0
			.arr(col + 1) = 0
			col += 2

			' 坐位体前屈
			.arr(col + 0) = 0
			.arr(col + 1) = 0
			col += 2

			If 是否有跳绳 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.tsJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有仰卧起坐 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.ywqzJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有50x8 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.nlpJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有立定跳远 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.tyJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有800米 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.nlpJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有1000米 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.nlpJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2
			If 是否有引体向上 = 1 Then
				.arr(col + 0) = "1"
				.arr(col + 1) = st.ywqzJfScore
			Else
				.arr(col + 0) = "0"
				.arr(col + 1) = "0"
			End If
			col += 2

			If st.grade = 3 Or st.grade = 5 Or st.grade = 7 Then
				.arr(col + 0) = Math.Round(Int((st.extraScore + 5) / 10) / 10, 1)
				col += 1
			End If

			If 是否参测 = 0 Then
				.arr(0) = "否"
			Else
				If 缺项数量 = 0 Then
					.arr(0) = "是"
				Else
					.arr(0) = "缺项"
				End If
			End If
			.arr(1) = 缺项数量
		End With

		' 保留一位小数
		excelWsDst.Columns(m50Pos + 1).NumberFormatLocal = "0.0_ "
		excelWsDst.Range(excelWsDst.Cells(row, 1), excelWsDst.Cells(row, col)).Value2 = st.arr
	End Sub

	Function getGradeName(ByVal grade As Long)
		Dim i, j
		i = LBound(gradeNameTbl)
		j = UBound(gradeNameTbl)
		If grade >= i And grade <= j Then
			getGradeName = gradeNameTbl(grade)
		Else
			getGradeName = "未知"
		End If
	End Function

	Function getStageName(ByVal grade As UInt32)
		If grade < 6 Then
			getStageName = "小学"
		ElseIf grade < 9 Then
			getStageName = "初中"
		ElseIf grade < 12 Then
			getStageName = "高中"
		Else
			getStageName = "未知"
		End If
	End Function

	Function getTotalLevel(ByVal score As Long)
		If score >= 8995 Then
			getTotalLevel = "优秀"
		ElseIf score >= 7995 Then
			getTotalLevel = "良好"
		ElseIf score >= 5995 Then
			getTotalLevel = "及格"
		Else
			getTotalLevel = "不及格"
		End If
	End Function

	Function getBmiLevel(ByVal bmi As Long, ByVal low As Long)
		If bmi <= 60 Then
			getBmiLevel = "肥胖"
		ElseIf bmi <= 80 Then
			If low = 1 Then
				getBmiLevel = "低体重"
			Else
				getBmiLevel = "超重"
			End If
		Else
			getBmiLevel = "正常"
		End If
	End Function

	Function getGeneralLevel(ByVal score As Long)
		If score >= 90 Then
			getGeneralLevel = "优秀"
		ElseIf score >= 80 Then
			getGeneralLevel = "良好"
		ElseIf score >= 60 Then
			getGeneralLevel = "及格"
		Else
			getGeneralLevel = "不及格"
		End If
	End Function

	Function getFhlLevel(ByVal fhl As Long)
		getFhlLevel = getGeneralLevel(fhl)
	End Function

	Function getM50Level(ByVal m50 As Long)
		getM50Level = getGeneralLevel(m50)
	End Function

	Function getZwtLevel(ByVal zwt As Long)
		getZwtLevel = getGeneralLevel(zwt)
	End Function

	Function getTsLevel(ByVal ts As Long)
		getTsLevel = getGeneralLevel(ts)
	End Function

	Function getTyLevel(ByVal ty As Long)
		getTyLevel = getGeneralLevel(ty)
	End Function

	Function getYwqzLevel(ByVal ywqz As Long)
		getYwqzLevel = getGeneralLevel(ywqz)
	End Function

	Function getNlpLevel(ByVal nlp As Long)
		getNlpLevel = getGeneralLevel(nlp)
	End Function

	' 0 1 2 3 4 5
	' 6 7 8
	' 9 10 11
	' 12 13 14 15
	Function calcGradeIdx(ByVal grade As Integer)
		If grade < 20 Then
			calcGradeIdx = grade - 11
		ElseIf grade < 30 Then
			calcGradeIdx = grade - 21 + 6
		ElseIf grade < 40 Then
			calcGradeIdx = grade - 31 + 9
		Else
			calcGradeIdx = grade - 41 + 12
		End If
	End Function

	Function stringToInt(ByVal s As String, ByVal rank As Integer)
		Dim v As Long
		Dim i As Integer
		Dim arr

		v = 0
		arr = Split(s, ".")
		If UBound(arr) < 0 Then
		ElseIf UBound(arr) < 1 Then
			For i = 1 To rank
				arr(0) = arr(0) + "0"
			Next i
			v = Val(arr(0))
		Else
			For i = 1 To rank
				arr(1) = arr(1) + "0"
			Next i
			v = Val(arr(0) & Mid(arr(1), 1, rank))
		End If

		stringToInt = v
	End Function

	Function doubleToInt(ByVal d As Double, ByVal rank As Integer)
		Dim s As String
		s = d
		doubleToInt = Int(stringToInt(s, rank))
	End Function

	Function calcBMI(ByRef st As Student)
		Dim h As Long
		Dim w As Long
		Dim tmp As Double
		calcBMI = 0
		h = st.height
		w = st.weight
		If h = 0 Then Exit Function
		tmp = w / (h * h)
		st.bmi = doubleToInt(tmp, 2)
		st.bmi = Int((st.bmi + 5) / 10)
	End Function

	Function calcBMIScore(ByRef st As Student)
		Dim idx As Integer

		calcBMIScore = 0
		' 计算BMI
		calcBMI(st)

		idx = st.grade
		' 大学计分规则一样
		If idx > 12 Then idx = 12
		' 女生的分数在后面
		If st.gender <> 1 Then idx = idx + 13 * 4

		st.bmiLow = 0
		If st.bmi <= BMIData(idx + 13) Then
			st.bmiScore = 80
			' 超重
			st.bmiLow = 1
		ElseIf st.bmi < BMIData(idx + 26) Then
			st.bmiScore = 100
		ElseIf st.bmi < BMIData(idx + 39) Then
			st.bmiScore = 80
		Else
			st.bmiScore = 60
		End If
	End Function

	Function calcFhlScoreImpl(ByRef st As Student, ByVal fhlData As Object)
		' 数据位置
		Dim offset As Integer
		Dim col As Integer
		Dim i As Integer

		calcFhlScoreImpl = 0
		col = st.grade
		' 大1和大2一样
		If col = 13 Then col = 12
		' 大3和大4一样
		If col > 13 Then col = 13

		' 第一列是分数
		col = col + 1
		offset = col

		st.fhlScore = 0
		For i = 0 To 19
			If st.fhl >= fhlData(offset) Then
				st.fhlScore = fhlData(offset - col)
				Exit For
			End If
			offset = offset + 15
		Next i
	End Function

	Sub calcFhlScore(ByRef st As Student)
		If st.gender = 1 Then
			calcFhlScoreImpl(st, fhlData0)
		Else
			calcFhlScoreImpl(st, fhlData1)
		End If
	End Sub

	Sub calcM50ScoreImpl(ByRef st As Student, ByVal m50Data As Object)
		' data position
		Dim offset As Long
		' column
		Dim col As Long
		Dim i As Long

		col = st.grade
		If col = 13 Then col = 12
		If col > 13 Then col = 13

		' 第一列是分数
		col = col + 1
		offset = col

		st.m50Score = 0
		For i = 0 To 19
			If st.m50 <= m50Data(offset) Then
				st.m50Score = m50Data(offset - col)
				Exit For
			End If
			offset = offset + 15
		Next i
	End Sub

	Sub calcM50Score(ByRef st As Student)
		If st.gender = 1 Then
			calcM50ScoreImpl(st, M50Data0)
		Else
			calcM50ScoreImpl(st, M50Data1)
		End If
	End Sub

	Sub calcZwtScoreImpl(ByRef st As Student, ByVal zwtData As Object)
		' data position
		Dim offset As Integer
		' column
		Dim col As Integer
		Dim i As Integer

		col = st.grade
		If col = 13 Then col = 12
		If col > 13 Then col = 13

		col = col + 1
		offset = col

		st.zwtScore = 0
		For i = 0 To 19
			If st.zwt >= zwtData(offset) Then
				st.zwtScore = zwtData(offset - col)
				Exit For
			End If
			offset = offset + 15
		Next i
	End Sub

	Sub calcZwtScore(ByRef st As Student)
		If st.gender = 1 Then
			calcZwtScoreImpl(st, zwtData0)
		Else
			calcZwtScoreImpl(st, zwtData1)
		End If
	End Sub

	Sub calcTsScoreImpl(ByRef st As Student, ByVal tsData As Object)
		' data position
		Dim offset As Integer
		' column
		Dim col As Integer
		Dim idx As Integer
		Dim i As Integer

		idx = st.grade
		If idx > 5 Then
			st.tsScore = 0
			GoTo found
		End If

		' the scores are in the first column
		col = idx + 1
		' 20 rows and 7 columns
		offset = col

		st.tsScore = 0
		For i = 0 To 19
			If st.ts >= tsData(offset) Then
				st.tsScore = tsData(offset - col)
				Exit For
			End If
			offset = offset + 7
		Next i

found:
		st.tsJfScore = 0
		If st.tsScore = 100 And idx < 6 Then
			Dim tmp As Long
			tmp = Int((st.ts - tsData(offset)) / 2)
			If tmp > 20 Then tmp = 20
			st.tsJfScore = tmp
		End If
	End Sub

	Sub calcTsScore(ByRef st As Student)
		If st.gender = 1 Then
			calcTsScoreImpl(st, tsData0)
		Else
			calcTsScoreImpl(st, tsData1)
		End If
	End Sub

	Sub calcTyScoreImpl(ByRef st As Student, ByVal tyData As Object)
		' data position
		Dim offset As Integer
		' column
		Dim col As Integer
		Dim i As Integer

		col = st.grade
		If col < 6 Then
			st.tyScore = 0
			Exit Sub
		ElseIf col < 12 Then
			col = col - 6
		ElseIf col < 14 Then
			col = 6
		Else
			col = 7
		End If

		col = col + 1
		offset = col

		st.tyScore = 0
		For i = 0 To 19
			If st.ty >= tyData(offset) Then
				st.tyScore = tyData(offset - col)
				Exit For
			End If
			offset = offset + 9
		Next i
	End Sub

	Sub calcTyScore(ByRef st As Student)
		If st.gender = 1 Then
			calcTyScoreImpl(st, tyData0)
		Else
			calcTyScoreImpl(st, tyData1)
		End If
	End Sub

	Sub calcYwqzScoreImpl(ByRef st As Student, ByVal ywqzData As Object)
		' data position
		Dim offset As Integer
		' column
		Dim col As Integer
		Dim idx As Integer
		Dim i As Integer

		idx = st.grade
		If idx < 2 Then
			st.ywqzScore = 0
			GoTo found
		ElseIf idx < 12 Then
			idx = idx - 2
		ElseIf idx < 14 Then
			idx = 10
		Else
			idx = 11
		End If

		col = idx + 1
		offset = col

		st.ywqzScore = 0
		For i = 0 To 19
			If st.ywqz0 >= ywqzData(offset) Then
				st.ywqzScore = ywqzData(offset - col)
				Exit For
			End If
			offset = offset + 13
		Next i

found:
		' 计算附加分
		st.ywqzJfScore = 0
		If st.ywqzScore = 100 Then
			Dim tmp As Long
			tmp = st.ywqz0 - ywqzData(offset)
			If st.gender <> 1 Then
				If tmp < 7 Then
					tmp = Int(tmp / 2)
				Else
					tmp = tmp - 3
				End If
			End If
			If st.grade < 6 Then tmp = 0
			If tmp > 10 Then tmp = 10
			st.ywqzJfScore = tmp
		End If
	End Sub

	Sub calcYwqzScore(ByRef st As Student)
		If st.gender = 1 Then
			calcYwqzScoreImpl(st, ywqzData0)
		Else
			calcYwqzScoreImpl(st, ywqzData1)
		End If
	End Sub

	Sub calcNlpScoreImpl(ByRef st As Student, ByVal nlpData As Object)
		' data position
		Dim offset As Integer
		' column
		Dim col As Integer
		Dim idx As Integer
		Dim i As Integer

		idx = st.grade
		If idx < 4 Then
			st.nlpScore = 0
			GoTo found
		ElseIf idx < 12 Then
			idx = idx - 4
		ElseIf idx < 14 Then
			idx = 8
		Else
			idx = 9
		End If

		col = idx + 1
		offset = col

		st.nlpScore = 0
		For i = 0 To 19
			If st.nlp0 <= nlpData(offset) Then
				st.nlpScore = nlpData(offset - col)
				' MsgBox "nlp " & i & " " & nlp & " " & idx & " " & calcNlpScoreImpl
				Exit For
			End If
			offset = offset + 11
		Next i

found:
		st.nlpJfScore = 0
		If st.nlpScore = 100 Then
			Dim x As Long
			Dim y As Long
			'x = 100 * nlpData(offset)
			'y = 100 * st.nlp0
			x = nlpData(offset)
			y = st.nlp0
			'Debug.Print x, y
			Dim tmp As Long
			'tmp = (Int(x / 100) - Int(y / 100)) * 60
			'tmp = tmp + Int(x Mod 100 - y Mod 100)
			tmp = x - y
			If st.gender = 1 Then
				If tmp < 23 Then
					tmp = Int(tmp / 4)
				Else
					tmp = 6 + Int((tmp - 23) / 3)
				End If
			Else
				tmp = Int(tmp / 5)
			End If
			If st.grade < 6 Then tmp = 0
			If tmp > 10 Then tmp = 10
			st.nlpJfScore = tmp
			'Debug.Print tmp
		End If
	End Sub

	Sub calcNlpScore(ByRef st As Student)
		If st.gender = 1 Then
			calcNlpScoreImpl(st, nlpData0)
		Else
			calcNlpScoreImpl(st, nlpData1)
		End If
	End Sub

	Sub calcTotalScore(ByRef st As Student)
		Dim idx As Long
		idx = st.grade

		' Debug.Print st.bmiScore, st.fhlScore, st.m50Score, st.zwtScore; st.tsScore, st.tyScore, st.nlpScore, st.nlpJfScore, st.ywqzScore, st.ywqzJfScore

		If idx < 2 Then
			st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 30 + st.tsScore * 20
		ElseIf idx < 4 Then
			st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 20 + st.tsScore * 20 + st.ywqzScore * 10
		ElseIf idx < 6 Then
			st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 10 + st.tsScore * 10 + st.ywqzScore * 20 + st.nlpScore * 10
		Else
			st.totalScore = st.bmiScore * 15 + st.fhlScore * 15 + st.m50Score * 20 + st.zwtScore * 10 + st.tyScore * 10 + st.ywqzScore * 10 + st.nlpScore * 20
		End If

		' 附加分
		st.totalJfScore = (st.tsJfScore + st.ywqzJfScore + st.nlpJfScore) * 100
	End Sub

	Sub calcExtraScore(ByRef st As Student)
		If st.grade = 3 Or st.grade = 5 Or st.grade = 7 Then
			Dim finalScore As Long
			finalScore = st.totalScore + st.totalJfScore
			If finalScore >= 7995 Then
				st.extraScore = 1000
			ElseIf finalScore >= 7495 Then
				st.extraScore = 850
			ElseIf finalScore >= 6995 Then
				st.extraScore = 800
			ElseIf finalScore >= 6495 Then
				st.extraScore = 750
			ElseIf finalScore >= 5995 Then
				st.extraScore = 700
			Else
				st.extraScore = 550
			End If
		End If
	End Sub

	Private Sub 计算成绩ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 计算成绩ToolStripMenuItem1.Click
		点击事件(OpType.StudentScore)
	End Sub

	Private Sub 停止处理ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 停止处理ToolStripMenuItem1.Click
		cancelReport(0)
	End Sub

	Private Sub 清空日志ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 清空日志ToolStripMenuItem1.Click
		RichTextBox1.Text = ""
	End Sub

	Private Sub 区百分比ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 区百分比ToolStripMenuItem1.Click
		点击事件(OpType.AreaPercent)
	End Sub

	Private Sub RichTextBox1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseDown
		If e.Button = Windows.Forms.MouseButtons.Right Then
			Dim p As Point
			p = PointToScreen(e.Location)
			p.Y += 28
			ContextMenuStrip1.Show(p)
		End If
	End Sub

	Private Sub 清空日志ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 清空日志ToolStripMenuItem.Click
		RichTextBox1.Text = ""
	End Sub

	Private Sub 清空状态ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 清空状态ToolStripMenuItem.Click
		ToolStripStatusLabel1.Text = "状态"
	End Sub

	Private Sub 清空状态ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 清空状态ToolStripMenuItem1.Click
		ToolStripStatusLabel1.Text = "状态"
	End Sub

	Private Sub 学生报告ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 学生报告ToolStripMenuItem.Click
		点击事件(OpType.StudentReport)
	End Sub

	Private Sub 学校报告ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 学校报告ToolStripMenuItem1.Click
		点击事件(OpType.SchoolReport)
	End Sub

	Private Sub 年级报告ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 年级报告ToolStripMenuItem1.Click
		点击事件(OpType.GradeReport)
	End Sub

	Private Sub 班级报告ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 班级报告ToolStripMenuItem2.Click
		点击事件(OpType.ClassReport)
	End Sub

	Private Sub 全部报告ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 全部报告ToolStripMenuItem.Click
		点击事件(OpType.AllReport)
	End Sub
End Class

Public Class Student
	Public idStr As String
	Public schoolStr As String
	Public classStr As String
	Public nameStr As String
	Public genderStr As String
	Public gradeStr As String
	Public heightStr As String
	Public weightStr As String
	Public fhlStr As String
	Public m50Str As String
	Public zwtStr As String
	Public tsStr As String
	Public tyStr As String
	Public ywqz0Str As String
	Public ywqz1Str As String
	Public nlp0Str As String
	Public nlp1Str As String

	Public gender As Long
	Public grade As Long
	' 10倍
	Public height As Long
	' 1000000倍
	Public weight As Long

	Public bmiValid As Long
	Public bmiScore As Long
	' 10倍
	Public bmi As Long
	Public bmiLow As Long

	Public fhlValid As Long
	Public fhl As Long
	Public fhlScore As Long

	Public m50Valid As Long
	' * 10
	Public m50 As Long
	Public m50Score As Long

	Public zwtValid As Long
	' 10倍
	Public zwt As Long
	Public zwtScore As Long

	Public tsValid As Long
	Public ts As Long
	Public tsScore As Long
	Public tsJfScore As Long

	Public tyValid As Long
	Public ty As Long
	Public tyScore As Long
	Public tyJfScore As Long

	Public ywqzValid As Long
	Public ywqz0 As Long
	Public ywqz1 As Long
	Public ywqzScore As Long
	Public ywqzJfScore As Long

	Public nlpValid As Long
	' 秒
	Public nlp0 As Long
	Public nlp1 As Long
	Public nlpScore As Long
	Public nlpJfScore As Long

	Public totalValid As Long
	' 100倍
	Public totalScore As Long
	' 100倍
	Public totalJfScore As Long
	' 100倍
	Public extraScore As Long

	' 存储报表数据
	Public arr() As Object
End Class

Public Enum MsgType
	mtNormal
	mtProgress
End Enum

Public Class MsgEntity
	Public type As MsgType
	Public data As Object
End Class