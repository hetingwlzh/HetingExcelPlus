<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 数据拆分控件
    Inherits System.Windows.Forms.UserControl

    'UserControl 重写释放以清理组件列表。
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(数据拆分控件))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.区域选择控件1 = New MyExcel插件.区域选择控件()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 280)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(301, 35)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "列标题所在区域的行号："
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.NumericUpDown1.Location = New System.Drawing.Point(310, 278)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(120, 40)
        Me.NumericUpDown1.TabIndex = 26
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(328, 641)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(175, 56)
        Me.Button1.TabIndex = 27
        Me.Button1.Text = "开始拆分"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(6, 52)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(119, 39)
        Me.CheckBox1.TabIndex = 28
        Me.CheckBox1.Text = "工作表"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(9, 366)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(494, 108)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "拆分为"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(166, 52)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(93, 39)
        Me.CheckBox2.TabIndex = 29
        Me.CheckBox2.Text = "文件"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Green
        Me.Label2.Location = New System.Drawing.Point(4, 510)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 28)
        Me.Label2.TabIndex = 30
        Me.Label2.Text = "保存路径"
        '
        '区域选择控件1
        '
        Me.区域选择控件1.heting = Nothing
        Me.区域选择控件1.Location = New System.Drawing.Point(3, 3)
        Me.区域选择控件1.Name = "区域选择控件1"
        Me.区域选择控件1.Size = New System.Drawing.Size(500, 253)
        Me.区域选择控件1.TabIndex = 0
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
        Me.区域选择控件1.空区域校验 = True
        Me.区域选择控件1.错误信息序列 = CType(resources.GetObject("区域选择控件1.错误信息序列"), System.Collections.ArrayList)
        '
        '数据拆分控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.区域选择控件1)
        Me.Name = "数据拆分控件"
        Me.Size = New System.Drawing.Size(516, 719)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents 区域选择控件1 As 区域选择控件
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents FolderBrowserDialog1 As Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label2 As Windows.Forms.Label
End Class
