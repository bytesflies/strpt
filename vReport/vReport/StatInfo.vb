
Public Class 统计信息
	Public 区 As String

	Public 报名人数 As UInt32
	Public 参测人数 As UInt32
	Public 完测人数 As UInt32
	Public 加分人数 As UInt32

	' 1: 12个统计项
	' 2:
	'	小学 初中 高中 全学段
	'' '' '' '' ''	一年级 二年级 ... 初一 初二 初三 高一 高二 高三 ... 全年级
	'' '' '' '' ''   班级1 班级2 班级3 ... 全班级（某个年级）
	' 3: 四个等级
	Public 等级(11, 3, 3) As UInt32

	' * 10000
	Public 参测比例 As UInt32
	Public 完测比例 As UInt32
	Public 百分比(11, 3, 3) As UInt32
End Class

Public Class 学校信息详细
	Public 年级统计信息(15) As 统计信息
	'Public 班级信息 As Dictionary(Of String, UInt32)
	Public 班级信息详细(15) As Dictionary(Of String, 统计信息)

	Sub New()
		'班级信息 = New Dictionary(Of String, UInt32)
		Dim i As Int32
		For i = 0 To 15
			班级信息详细(i) = New Dictionary(Of String, 统计信息)
		Next
	End Sub
End Class

Public Class 统计项
	Public 类型 As UInt32
	Public 序号 As UInt32
	Public 名称 As String
	Public 学段(3) As Int32
	Public 关键字 As String

	Sub New(ByVal x As UInt32, ByVal y As UInt32, ByRef z As String, ByRef a() As Int32, ByRef b As String)
		类型 = x
		序号 = y
		名称 = z
		学段(0) = a(0)
		学段(1) = a(1)
		学段(2) = a(2)
		学段(3) = a(3)
		关键字 = b
	End Sub

End Class
