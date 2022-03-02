
Public Class 统计信息
	Public 报名人数 As UInt32
	Public 参测人数 As UInt32
	Public 完测人数 As UInt32
	Public 加分人数 As UInt32

	' 12个统计项
	' 小学 初中 高中 全学段
	' 四个等级
	Public 等级(11, 3, 3) As UInt32
End Class

Public Class 统计项
	Public 类型 As UInt32
	Public 序号 As UInt32
	Public 名称 As String

	Sub New(ByVal x As UInt32, ByVal y As UInt32, ByRef z As String)
		类型 = x
		序号 = y
		名称 = z
	End Sub

End Class
