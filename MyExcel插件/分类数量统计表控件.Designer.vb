<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 分类数量统计表控件
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(分类数量统计表控件))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.区域选择控件1 = New MyExcel插件.区域选择控件()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.Button1.Location = New System.Drawing.Point(8, 491)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(287, 58)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "生成  分类数量统计表"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDown1.Location = New System.Drawing.Point(252, 394)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(120, 35)
        Me.NumericUpDown1.TabIndex = 24
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 396)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(243, 28)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "列标题所在区域的行号："
        '
        '区域选择控件1
        '
        Me.区域选择控件1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.区域选择控件1.heting = Nothing
        Me.区域选择控件1.Location = New System.Drawing.Point(3, 3)
        Me.区域选择控件1.Name = "区域选择控件1"
        Me.区域选择控件1.Size = New System.Drawing.Size(495, 375)
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
        Me.区域选择控件1.空区域校验 = False
        Me.区域选择控件1.错误信息序列 = CType(resources.GetObject("区域选择控件1.错误信息序列"), System.Collections.ArrayList)
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(8, 442)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(164, 32)
        Me.CheckBox1.TabIndex = 26
        Me.CheckBox1.Text = "是否忽略空值"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        '分类数量统计表控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.区域选择控件1)
        Me.Name = "分类数量统计表控件"
        Me.Size = New System.Drawing.Size(502, 565)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents 区域选择控件1 As 区域选择控件
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
End Class
