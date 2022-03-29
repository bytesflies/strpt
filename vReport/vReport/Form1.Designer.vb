<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
		Me.报告 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.学校 = New System.Windows.Forms.ToolStripMenuItem
		Me.年级ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.班级报告ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
		Me.数据处理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.计算成绩ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.区百分比ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.计算成绩ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.区百分比ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.学校报告ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.年级报告ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.班级报告ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.停止处理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.停止处理ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.清空日志ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
		Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
		Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
		Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
		Me.清空日志ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.清空状态ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
		Me.报告.SuspendLayout()
		Me.MenuStrip1.SuspendLayout()
		Me.StatusStrip1.SuspendLayout()
		Me.ContextMenuStrip1.SuspendLayout()
		Me.SuspendLayout()
		'
		'RichTextBox1
		'
		Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
					Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Control
		Me.RichTextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.RichTextBox1.Location = New System.Drawing.Point(0, 28)
		Me.RichTextBox1.Name = "RichTextBox1"
		Me.RichTextBox1.ReadOnly = True
		Me.RichTextBox1.Size = New System.Drawing.Size(909, 478)
		Me.RichTextBox1.TabIndex = 1
		Me.RichTextBox1.Text = ""
		'
		'报告
		'
		Me.报告.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.学校, Me.年级ToolStripMenuItem, Me.班级报告ToolStripMenuItem})
		Me.报告.Name = "报告"
		Me.报告.Size = New System.Drawing.Size(125, 70)
		Me.报告.Text = "报告"
		'
		'学校
		'
		Me.学校.Name = "学校"
		Me.学校.Size = New System.Drawing.Size(124, 22)
		Me.学校.Text = "学校报告"
		'
		'年级ToolStripMenuItem
		'
		Me.年级ToolStripMenuItem.Name = "年级ToolStripMenuItem"
		Me.年级ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.年级ToolStripMenuItem.Text = "年级报告"
		'
		'班级报告ToolStripMenuItem
		'
		Me.班级报告ToolStripMenuItem.Name = "班级报告ToolStripMenuItem"
		Me.班级报告ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.班级报告ToolStripMenuItem.Text = "班级报告"
		'
		'MenuStrip1
		'
		Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.数据处理ToolStripMenuItem, Me.计算成绩ToolStripMenuItem, Me.停止处理ToolStripMenuItem})
		Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
		Me.MenuStrip1.Name = "MenuStrip1"
		Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
		Me.MenuStrip1.Size = New System.Drawing.Size(909, 25)
		Me.MenuStrip1.TabIndex = 7
		Me.MenuStrip1.Text = "MenuStrip1"
		'
		'数据处理ToolStripMenuItem
		'
		Me.数据处理ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.计算成绩ToolStripMenuItem1, Me.区百分比ToolStripMenuItem1})
		Me.数据处理ToolStripMenuItem.Name = "数据处理ToolStripMenuItem"
		Me.数据处理ToolStripMenuItem.Padding = New System.Windows.Forms.Padding(0, 0, 4, 0)
		Me.数据处理ToolStripMenuItem.Size = New System.Drawing.Size(64, 21)
		Me.数据处理ToolStripMenuItem.Text = "数据处理"
		'
		'计算成绩ToolStripMenuItem1
		'
		Me.计算成绩ToolStripMenuItem1.Name = "计算成绩ToolStripMenuItem1"
		Me.计算成绩ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
		Me.计算成绩ToolStripMenuItem1.Text = "计算成绩"
		'
		'区百分比ToolStripMenuItem1
		'
		Me.区百分比ToolStripMenuItem1.Name = "区百分比ToolStripMenuItem1"
		Me.区百分比ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
		Me.区百分比ToolStripMenuItem1.Text = "区百分比"
		'
		'计算成绩ToolStripMenuItem
		'
		Me.计算成绩ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.区百分比ToolStripMenuItem, Me.学校报告ToolStripMenuItem, Me.年级报告ToolStripMenuItem, Me.班级报告ToolStripMenuItem1})
		Me.计算成绩ToolStripMenuItem.Name = "计算成绩ToolStripMenuItem"
		Me.计算成绩ToolStripMenuItem.Size = New System.Drawing.Size(68, 21)
		Me.计算成绩ToolStripMenuItem.Text = "生成报告"
		'
		'区百分比ToolStripMenuItem
		'
		Me.区百分比ToolStripMenuItem.Name = "区百分比ToolStripMenuItem"
		Me.区百分比ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.区百分比ToolStripMenuItem.Text = "学生报告"
		'
		'学校报告ToolStripMenuItem
		'
		Me.学校报告ToolStripMenuItem.Name = "学校报告ToolStripMenuItem"
		Me.学校报告ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.学校报告ToolStripMenuItem.Text = "学校报告"
		'
		'年级报告ToolStripMenuItem
		'
		Me.年级报告ToolStripMenuItem.Name = "年级报告ToolStripMenuItem"
		Me.年级报告ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.年级报告ToolStripMenuItem.Text = "年级报告"
		'
		'班级报告ToolStripMenuItem1
		'
		Me.班级报告ToolStripMenuItem1.Name = "班级报告ToolStripMenuItem1"
		Me.班级报告ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
		Me.班级报告ToolStripMenuItem1.Text = "班级报告"
		'
		'停止处理ToolStripMenuItem
		'
		Me.停止处理ToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.停止处理ToolStripMenuItem1, Me.清空日志ToolStripMenuItem1})
		Me.停止处理ToolStripMenuItem.Name = "停止处理ToolStripMenuItem"
		Me.停止处理ToolStripMenuItem.Size = New System.Drawing.Size(44, 21)
		Me.停止处理ToolStripMenuItem.Text = "其它"
		'
		'停止处理ToolStripMenuItem1
		'
		Me.停止处理ToolStripMenuItem1.Name = "停止处理ToolStripMenuItem1"
		Me.停止处理ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
		Me.停止处理ToolStripMenuItem1.Text = "停止处理"
		'
		'清空日志ToolStripMenuItem1
		'
		Me.清空日志ToolStripMenuItem1.Name = "清空日志ToolStripMenuItem1"
		Me.清空日志ToolStripMenuItem1.Size = New System.Drawing.Size(124, 22)
		Me.清空日志ToolStripMenuItem1.Text = "清空日志"
		'
		'StatusStrip1
		'
		Me.StatusStrip1.GripMargin = New System.Windows.Forms.Padding(0)
		Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
		Me.StatusStrip1.Location = New System.Drawing.Point(0, 509)
		Me.StatusStrip1.Name = "StatusStrip1"
		Me.StatusStrip1.Size = New System.Drawing.Size(909, 22)
		Me.StatusStrip1.TabIndex = 8
		Me.StatusStrip1.Text = "状态1"
		'
		'ToolStripStatusLabel1
		'
		Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
		Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(32, 17)
		Me.ToolStripStatusLabel1.Text = "状态"
		'
		'ContextMenuStrip1
		'
		Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.清空日志ToolStripMenuItem, Me.清空状态ToolStripMenuItem})
		Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
		Me.ContextMenuStrip1.Size = New System.Drawing.Size(125, 48)
		'
		'清空日志ToolStripMenuItem
		'
		Me.清空日志ToolStripMenuItem.Name = "清空日志ToolStripMenuItem"
		Me.清空日志ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.清空日志ToolStripMenuItem.Text = "清空日志"
		'
		'清空状态ToolStripMenuItem
		'
		Me.清空状态ToolStripMenuItem.Name = "清空状态ToolStripMenuItem"
		Me.清空状态ToolStripMenuItem.Size = New System.Drawing.Size(124, 22)
		Me.清空状态ToolStripMenuItem.Text = "清空状态"
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(909, 531)
		Me.Controls.Add(Me.StatusStrip1)
		Me.Controls.Add(Me.MenuStrip1)
		Me.Controls.Add(Me.RichTextBox1)
		Me.Name = "Form1"
		Me.Text = "工具"
		Me.报告.ResumeLayout(False)
		Me.MenuStrip1.ResumeLayout(False)
		Me.MenuStrip1.PerformLayout()
		Me.StatusStrip1.ResumeLayout(False)
		Me.StatusStrip1.PerformLayout()
		Me.ContextMenuStrip1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
	Friend WithEvents 报告 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents 学校 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 年级ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 班级报告ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
	Friend WithEvents 数据处理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 计算成绩ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 区百分比ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 计算成绩ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 区百分比ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 学校报告ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 年级报告ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 班级报告ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 停止处理ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 停止处理ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 清空日志ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
	Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
	Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
	Friend WithEvents 清空日志ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
	Friend WithEvents 清空状态ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
