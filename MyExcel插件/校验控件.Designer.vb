<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 校验控件
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(校验控件))
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.区域选择控件1 = New MyExcel插件.区域选择控件()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.FlowLayoutPanel4 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel3.SuspendLayout()
        Me.FlowLayoutPanel4.SuspendLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.NumericUpDown1.Location = New System.Drawing.Point(310, 3)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(120, 40)
        Me.NumericUpDown1.TabIndex = 24
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(301, 35)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "列标题所在区域的行号："
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.Button1.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Button1.Location = New System.Drawing.Point(3, 585)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(167, 46)
        Me.Button1.TabIndex = 25
        Me.Button1.Text = "开始校验"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel2, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.区域选择控件1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Button1, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel4, 0, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 49.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 66.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 63.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(515, 648)
        Me.TableLayoutPanel1.TabIndex = 27
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel2.Controls.Add(Me.ComboBox1)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(3, 396)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(509, 54)
        Me.FlowLayoutPanel2.TabIndex = 30
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(145, 35)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "校验类型："
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"身份证校验", "手机号校验", "电子邮箱校验", "长度校验", "异同校验", "非空校验", "包含校验", "不含校验", "开头校验", "结尾校验"})
        Me.ComboBox1.Location = New System.Drawing.Point(154, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(279, 40)
        Me.ComboBox1.TabIndex = 29
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel1.Controls.Add(Me.NumericUpDown1)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 347)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(509, 43)
        Me.FlowLayoutPanel1.TabIndex = 28
        '
        '区域选择控件1
        '
        Me.区域选择控件1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.区域选择控件1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.区域选择控件1.heting = Nothing
        Me.区域选择控件1.Location = New System.Drawing.Point(3, 3)
        Me.区域选择控件1.Name = "区域选择控件1"
        Me.区域选择控件1.Size = New System.Drawing.Size(509, 338)
        Me.区域选择控件1.TabIndex = 0
        Me.区域选择控件1.区域列表 = CType(resources.GetObject("区域选择控件1.区域列表"), System.Collections.ArrayList)
        Me.区域选择控件1.区域名集合 = CType(resources.GetObject("区域选择控件1.区域名集合"), System.Collections.ArrayList)
        Me.区域选择控件1.单列校验 = True
        Me.区域选择控件1.单行校验 = False
        Me.区域选择控件1.是否允许编辑区域名 = False
        Me.区域选择控件1.是否显示删除按钮 = False
        Me.区域选择控件1.是否显示添加按钮 = False
        Me.区域选择控件1.是否显示锁定列 = False
        Me.区域选择控件1.是否显示锁定用户区域 = True
        Me.区域选择控件1.是否显示锁定行 = False
        Me.区域选择控件1.是否显示锁定表 = False
        Me.区域选择控件1.是否显示错误信息 = False
        Me.区域选择控件1.空区域校验 = True
        Me.区域选择控件1.错误信息序列 = CType(resources.GetObject("区域选择控件1.错误信息序列"), System.Collections.ArrayList)
        '
        'FlowLayoutPanel3
        '
        Me.FlowLayoutPanel3.Controls.Add(Me.Label3)
        Me.FlowLayoutPanel3.Controls.Add(Me.TextBox1)
        Me.FlowLayoutPanel3.Controls.Add(Me.Label4)
        Me.FlowLayoutPanel3.Controls.Add(Me.TextBox2)
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(3, 522)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(502, 55)
        Me.FlowLayoutPanel3.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 35)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "正确:"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.TextBox1.Location = New System.Drawing.Point(82, 3)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(142, 40)
        Me.TextBox1.TabIndex = 29
        Me.TextBox1.Text = "✔"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(230, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 35)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "错误:"
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.TextBox2.Location = New System.Drawing.Point(309, 3)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(142, 40)
        Me.TextBox2.TabIndex = 29
        Me.TextBox2.Text = "❌❌❌"
        '
        'FlowLayoutPanel4
        '
        Me.FlowLayoutPanel4.Controls.Add(Me.Label5)
        Me.FlowLayoutPanel4.Controls.Add(Me.NumericUpDown2)
        Me.FlowLayoutPanel4.Controls.Add(Me.TextBox3)
        Me.FlowLayoutPanel4.Location = New System.Drawing.Point(3, 456)
        Me.FlowLayoutPanel4.Name = "FlowLayoutPanel4"
        Me.FlowLayoutPanel4.Size = New System.Drawing.Size(502, 60)
        Me.FlowLayoutPanel4.TabIndex = 31
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(3, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(125, 35)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "校验参数:"
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.NumericUpDown2.Location = New System.Drawing.Point(134, 3)
        Me.NumericUpDown2.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(120, 40)
        Me.NumericUpDown2.TabIndex = 31
        Me.NumericUpDown2.Value = New Decimal(New Integer() {8, 0, 0, 0})
        '
        'TextBox3
        '
        Me.TextBox3.Font = New System.Drawing.Font("微软雅黑", 10.71429!)
        Me.TextBox3.Location = New System.Drawing.Point(260, 3)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(142, 40)
        Me.TextBox3.TabIndex = 32
        '
        '校验控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "校验控件"
        Me.Size = New System.Drawing.Size(515, 648)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel3.ResumeLayout(False)
        Me.FlowLayoutPanel3.PerformLayout()
        Me.FlowLayoutPanel4.ResumeLayout(False)
        Me.FlowLayoutPanel4.PerformLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents 区域选择控件1 As 区域选择控件
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents FlowLayoutPanel2 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents FlowLayoutPanel3 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents FlowLayoutPanel4 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents NumericUpDown2 As Windows.Forms.NumericUpDown
    Friend WithEvents TextBox3 As Windows.Forms.TextBox
End Class
