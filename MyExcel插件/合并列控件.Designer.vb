<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 合并列控件
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(合并列控件))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.粘贴ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.MyDataGridView1 = New HetingControl.MyDataGridView()
        Me.区域选择控件1 = New MyExcel插件.区域选择控件()
        Me.GroupBox1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.MyDataGridView1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(3, 396)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(549, 390)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "设置忽略值"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.粘贴ToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(127, 38)
        '
        '粘贴ToolStripMenuItem
        '
        Me.粘贴ToolStripMenuItem.Name = "粘贴ToolStripMenuItem"
        Me.粘贴ToolStripMenuItem.Size = New System.Drawing.Size(126, 34)
        Me.粘贴ToolStripMenuItem.Text = "粘贴"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(3, 792)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(93, 39)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "空值"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(3, 925)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(153, 55)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "开始合并"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDown1.Location = New System.Drawing.Point(180, 0)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(130, 40)
        Me.NumericUpDown1.TabIndex = 24
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(171, 35)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "列标题行号："
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.区域选择控件1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.CheckBox1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Button1, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox2, 0, 4)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 6
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 342.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 51.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 45.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 88.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 66.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(555, 988)
        Me.TableLayoutPanel1.TabIndex = 25
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.NumericUpDown1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 345)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(549, 45)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(3, 837)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(392, 76)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "合并方式"
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(180, 40)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(118, 39)
        Me.RadioButton2.TabIndex = 0
        Me.RadioButton2.Text = "全格式"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(17, 36)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(66, 39)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "值"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'MyDataGridView1
        '
        Me.MyDataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MyDataGridView1.Location = New System.Drawing.Point(3, 36)
        Me.MyDataGridView1.Margin = New System.Windows.Forms.Padding(5, 8, 5, 8)
        Me.MyDataGridView1.Name = "MyDataGridView1"
        Me.MyDataGridView1.Size = New System.Drawing.Size(543, 351)
        Me.MyDataGridView1.TabIndex = 6
        Me.MyDataGridView1.列宽模式 = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.MyDataGridView1.列排序模式 = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView1.列标题单元格样式 = DataGridViewCellStyle1
        Me.MyDataGridView1.字体 = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.MyDataGridView1.当前单元格 = Nothing
        Me.MyDataGridView1.是否显示行标题列 = True
        Me.MyDataGridView1.背景色 = System.Drawing.SystemColors.AppWorkspace
        Me.MyDataGridView1.行标题列宽度 = 72
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.MyDataGridView1.行标题单元格样式 = DataGridViewCellStyle2
        Me.MyDataGridView1.行高模式 = System.Windows.Forms.DataGridViewAutoSizeRowsMode.None
        Me.MyDataGridView1.边框颜色 = System.Drawing.SystemColors.ControlDark
        Me.MyDataGridView1.选定单元格前景色 = System.Drawing.SystemColors.HighlightText
        Me.MyDataGridView1.选定单元格背景色 = System.Drawing.SystemColors.Highlight
        Me.MyDataGridView1.默认单元格前景色 = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.MyDataGridView1.默认单元格样式 = DataGridViewCellStyle3
        Me.MyDataGridView1.默认单元格背景色 = System.Drawing.SystemColors.Window
        '
        '区域选择控件1
        '
        Me.区域选择控件1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.区域选择控件1.heting = Nothing
        Me.区域选择控件1.Location = New System.Drawing.Point(3, 3)
        Me.区域选择控件1.Name = "区域选择控件1"
        Me.区域选择控件1.Size = New System.Drawing.Size(549, 336)
        Me.区域选择控件1.TabIndex = 0
        Me.区域选择控件1.区域列表 = CType(resources.GetObject("区域选择控件1.区域列表"), System.Collections.ArrayList)
        Me.区域选择控件1.区域名集合 = CType(resources.GetObject("区域选择控件1.区域名集合"), System.Collections.ArrayList)
        Me.区域选择控件1.单列校验 = False
        Me.区域选择控件1.单行校验 = False
        Me.区域选择控件1.是否允许编辑区域名 = False
        Me.区域选择控件1.是否显示删除按钮 = False
        Me.区域选择控件1.是否显示添加按钮 = False
        Me.区域选择控件1.是否显示锁定列 = False
        Me.区域选择控件1.是否显示锁定用户区域 = False
        Me.区域选择控件1.是否显示锁定行 = False
        Me.区域选择控件1.是否显示锁定表 = False
        Me.区域选择控件1.是否显示错误信息 = False
        Me.区域选择控件1.空区域校验 = False
        Me.区域选择控件1.错误信息序列 = CType(resources.GetObject("区域选择控件1.错误信息序列"), System.Collections.ArrayList)
        '
        '合并列控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 21.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "合并列控件"
        Me.Size = New System.Drawing.Size(555, 988)
        Me.GroupBox1.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents 区域选择控件1 As 区域选择控件
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton2 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As Windows.Forms.RadioButton
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents 粘贴ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents MyDataGridView1 As HetingControl.MyDataGridView
End Class
