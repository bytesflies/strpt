
Public Class 统计信息
	Public 区 As String

	Public 报名人数 As UInt32
	Public 参测人数 As UInt32
	Public 完测人数 As UInt32
	Public 加分人数 As UInt32

	' 12个统计项
	' 小学 初中 高中 全学段
	' 四个等级
	Public 等级(11, 3, 3) As UInt32

	' * 10000
	Public 参测比例 As UInt32
	Public 完测比例 As UInt32
	Public 百分比(11, 3, 3) As UInt32
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
