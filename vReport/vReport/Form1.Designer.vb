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
		Me.SuspendLayout()
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(13, 10)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(71, 28)
		Me.Button1.TabIndex = 0
		Me.Button1.Text = "生成数据"
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
		Me.RichTextBox1.Size = New System.Drawing.Size(878, 417)
		Me.RichTextBox1.TabIndex = 1
		Me.RichTextBox1.Text = ""
		'
		'Button2
		'
		Me.Button2.Location = New System.Drawing.Point(167, 10)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(71, 28)
		Me.Button2.TabIndex = 2
		Me.Button2.Text = "停止生成"
		Me.Button2.UseVisualStyleBackColor = True
		'
		'Button3
		'
		Me.Button3.Location = New System.Drawing.Point(90, 10)
		Me.Button3.Name = "Button3"
		Me.Button3.Size = New System.Drawing.Size(75, 28)
		Me.Button3.TabIndex = 3
		Me.Button3.Text = "生成报告"
		Me.Button3.UseVisualStyleBackColor = True
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(903, 477)
		Me.Controls.Add(Me.Button3)
		Me.Controls.Add(Me.Button2)
		Me.Controls.Add(Me.RichTextBox1)
		Me.Controls.Add(Me.Button1)
		Me.Name = "Form1"
		Me.Text = "报告"
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents Button1 As System.Windows.Forms.Button
	Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
	Friend WithEvents Button2 As System.Windows.Forms.Button
	Friend WithEvents Button3 As System.Windows.Forms.Button

End Class
