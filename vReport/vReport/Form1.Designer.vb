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
		Me.Button1 = New System.Windows.Forms.Button
		Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
		Me.Button2 = New System.Windows.Forms.Button
		Me.Button3 = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.Button4 = New System.Windows.Forms.Button
		Me.Button5 = New System.Windows.Forms.Button
		Me.Button6 = New System.Windows.Forms.Button
		Me.SuspendLayout()
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(13, 12)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(71, 28)
		Me.Button1.TabIndex = 0
		Me.Button1.Text = "计算成绩"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'RichTextBox1
		'
		Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
					Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Control
		Me.RichTextBox1.Location = New System.Drawing.Point(13, 48)
		Me.RichTextBox1.Name = "RichTextBox1"
		Me.RichTextBox1.ReadOnly = True
		Me.RichTextBox1.Size = New System.Drawing.Size(884, 471)
		Me.RichTextBox1.TabIndex = 1
		Me.RichTextBox1.Text = ""
		'
		'Button2
		'
		Me.Button2.Location = New System.Drawing.Point(331, 12)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(71, 28)
		Me.Button2.TabIndex = 3
		Me.Button2.Text = "停止处理"
		Me.Button2.UseVisualStyleBackColor = True
		'
		'Button3
		'
		Me.Button3.Location = New System.Drawing.Point(89, 12)
		Me.Button3.Name = "Button3"
		Me.Button3.Size = New System.Drawing.Size(75, 28)
		Me.Button3.TabIndex = 1
		Me.Button3.Text = "学生报告"
		Me.Button3.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(504, 20)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(29, 12)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "    "
		'
		'Button4
		'
		Me.Button4.Location = New System.Drawing.Point(408, 12)
		Me.Button4.Name = "Button4"
		Me.Button4.Size = New System.Drawing.Size(71, 28)
		Me.Button4.TabIndex = 4
		Me.Button4.Text = "清空日志"
		Me.Button4.UseVisualStyleBackColor = True
		'
		'Button5
		'
		Me.Button5.Location = New System.Drawing.Point(169, 12)
		Me.Button5.Name = "Button5"
		Me.Button5.Size = New System.Drawing.Size(75, 28)
		Me.Button5.TabIndex = 2
		Me.Button5.Text = "学校报告"
		Me.Button5.UseVisualStyleBackColor = True
		'
		'Button6
		'
		Me.Button6.Location = New System.Drawing.Point(250, 15)
		Me.Button6.Name = "Button6"
		Me.Button6.Size = New System.Drawing.Size(75, 23)
		Me.Button6.TabIndex = 5
		Me.Button6.Text = "区百分比"
		Me.Button6.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(909, 531)
		Me.Controls.Add(Me.Button6)
		Me.Controls.Add(Me.Button5)
		Me.Controls.Add(Me.Button4)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.Button3)
		Me.Controls.Add(Me.Button2)
		Me.Controls.Add(Me.RichTextBox1)
		Me.Controls.Add(Me.Button1)
		Me.Name = "Form1"
		Me.Text = "报告"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
	Friend WithEvents Button2 As System.Windows.Forms.Button
	Friend WithEvents Button3 As System.Windows.Forms.Button
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Button4 As System.Windows.Forms.Button
	Friend WithEvents Button5 As System.Windows.Forms.Button
	Friend WithEvents Button6 As System.Windows.Forms.Button

End Class
