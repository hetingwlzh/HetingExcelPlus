<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 合并文件控件
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Timer1 = New HetingControl.Timer()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(555, 59)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "选择文件"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        Me.OpenFileDialog1.Multiselect = True
        '
        'Timer1
        '
        Me.Timer1.AutoSize = True
        Me.Timer1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Timer1.BackColor = System.Drawing.SystemColors.Control
        Me.Timer1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Timer1.Location = New System.Drawing.Point(9, 694)
        Me.Timer1.Name = "Timer1"
        Me.Timer1.Size = New System.Drawing.Size(294, 31)
        Me.Timer1.TabIndex = 1
        Me.Timer1.刷新间隔 = 100
        Me.Timer1.字体 = New System.Drawing.Font("微软雅黑", 10.0!)
        Me.Timer1.尺寸 = 10.0!
        Me.Timer1.是否启动 = True
        Me.Timer1.是否自动大小 = True
        Me.Timer1.显示样式 = HetingControl.Timer.ShowType.日期时间单行
        Me.Timer1.背景颜色 = System.Drawing.SystemColors.Control
        Me.Timer1.颜色 = System.Drawing.SystemColors.ControlText
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.LemonChiffon
        Me.ListBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 28
        Me.ListBox1.Location = New System.Drawing.Point(3, 113)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.Size = New System.Drawing.Size(555, 284)
        Me.ListBox1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Green
        Me.Label1.Location = New System.Drawing.Point(3, 400)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 35)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "选择文件"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button2.Location = New System.Drawing.Point(361, 639)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(197, 61)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "开始合并"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDown1.Location = New System.Drawing.Point(220, 63)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(120, 44)
        Me.NumericUpDown1.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(211, 36)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "列表头尾行号："
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button3.Location = New System.Drawing.Point(3, 627)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(99, 61)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "说明"
        Me.Button3.UseVisualStyleBackColor = True
        '
        '合并文件控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Timer1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "合并文件控件"
        Me.Size = New System.Drawing.Size(570, 741)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As Windows.Forms.OpenFileDialog
    Friend WithEvents Timer1 As HetingControl.Timer
    Friend WithEvents ListBox1 As Windows.Forms.ListBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Button3 As Windows.Forms.Button
End Class
