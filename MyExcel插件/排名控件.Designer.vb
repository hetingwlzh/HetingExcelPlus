<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 排名控件
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(排名控件))
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.区域选择控件1 = New MyExcel插件.区域选择控件()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(37, 46)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(92, 38)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.Text = "升序"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(187, 46)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(92, 38)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "降序"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(389, 669)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(148, 65)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "开始排名"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 352)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(525, 97)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "排序方式"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton5)
        Me.GroupBox2.Controls.Add(Me.RadioButton3)
        Me.GroupBox2.Controls.Add(Me.RadioButton4)
        Me.GroupBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 455)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(525, 101)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "相同数据"
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton5.Location = New System.Drawing.Point(337, 39)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(144, 38)
        Me.RadioButton5.TabIndex = 0
        Me.RadioButton5.Text = "最劣排名"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Checked = True
        Me.RadioButton3.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton3.Location = New System.Drawing.Point(37, 39)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(144, 38)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "最优排名"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton4.Location = New System.Drawing.Point(187, 39)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(144, 38)
        Me.RadioButton4.TabIndex = 1
        Me.RadioButton4.Text = "依次排名"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RadioButton7)
        Me.GroupBox4.Controls.Add(Me.RadioButton8)
        Me.GroupBox4.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 562)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(525, 101)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "是否忽略空值"
        '
        'RadioButton7
        '
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Checked = True
        Me.RadioButton7.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton7.Location = New System.Drawing.Point(37, 39)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(144, 38)
        Me.RadioButton7.TabIndex = 0
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.Text = "忽略空值"
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'RadioButton8
        '
        Me.RadioButton8.AutoSize = True
        Me.RadioButton8.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton8.Location = New System.Drawing.Point(187, 39)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(170, 38)
        Me.RadioButton8.TabIndex = 1
        Me.RadioButton8.Text = "不忽略空值"
        Me.RadioButton8.UseVisualStyleBackColor = True
        '
        '区域选择控件1
        '
        Me.区域选择控件1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.区域选择控件1.heting = Nothing
        Me.区域选择控件1.Location = New System.Drawing.Point(12, 3)
        Me.区域选择控件1.Name = "区域选择控件1"
        Me.区域选择控件1.Size = New System.Drawing.Size(525, 343)
        Me.区域选择控件1.TabIndex = 4
        Me.区域选择控件1.区域列表 = CType(resources.GetObject("区域选择控件1.区域列表"), System.Collections.ArrayList)
        Me.区域选择控件1.区域名集合 = CType(resources.GetObject("区域选择控件1.区域名集合"), System.Collections.ArrayList)
        Me.区域选择控件1.单列校验 = True
        Me.区域选择控件1.单行校验 = False
        Me.区域选择控件1.是否允许编辑区域名 = False
        Me.区域选择控件1.是否显示删除按钮 = False
        Me.区域选择控件1.是否显示添加按钮 = False
        Me.区域选择控件1.是否显示锁定列 = False
        Me.区域选择控件1.是否显示锁定用户区域 = False
        Me.区域选择控件1.是否显示锁定行 = False
        Me.区域选择控件1.是否显示锁定表 = False
        Me.区域选择控件1.是否显示错误信息 = True
        Me.区域选择控件1.空区域校验 = False
        Me.区域选择控件1.错误信息序列 = CType(resources.GetObject("区域选择控件1.错误信息序列"), System.Collections.ArrayList)
        '
        '排名控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.区域选择控件1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Location = New System.Drawing.Point(40, 20)
        Me.Name = "排名控件"
        Me.Size = New System.Drawing.Size(544, 743)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents RadioButton1 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As Windows.Forms.RadioButton
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton3 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton7 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton8 As Windows.Forms.RadioButton
    Friend WithEvents 区域选择控件1 As 区域选择控件
End Class
