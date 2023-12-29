<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 列汇总控件
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.结果表 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.删除内容ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.删除行ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.全部清除ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.数据源表 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(470, 161)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(71, 37)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "刷新"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Button2.Location = New System.Drawing.Point(3, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(162, 44)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "按列汇总"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(111, 68)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(316, 40)
        Me.ComboBox2.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 114)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 28)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "数据源："
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(9, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 28)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "填充表："
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button3.Location = New System.Drawing.Point(233, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(131, 44)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "说明"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton4)
        Me.GroupBox2.Controls.Add(Me.RadioButton3)
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(14, 330)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(474, 74)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "拷贝方式"
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton4.Location = New System.Drawing.Point(363, 34)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(79, 32)
        Me.RadioButton4.TabIndex = 0
        Me.RadioButton4.Text = "格式"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton3.Location = New System.Drawing.Point(242, 34)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(79, 32)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.Text = "公式"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(25, 34)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(79, 32)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "全部"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(147, 34)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(58, 32)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "值"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.CheckedListBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(111, 117)
        Me.CheckedListBox1.MultiColumn = True
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(355, 164)
        Me.CheckedListBox1.TabIndex = 9
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button4.Location = New System.Drawing.Point(470, 114)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(71, 40)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "反选"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(106, 284)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(363, 43)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "当前数据表（共0个）"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(436, 71)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(122, 32)
        Me.CheckBox1.TabIndex = 10
        Me.CheckBox1.Text = "创建新表"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(14, 145)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(80, 32)
        Me.CheckBox2.TabIndex = 11
        Me.CheckBox2.Text = "多列"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.结果表})
        Me.DataGridView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 30
        Me.DataGridView1.RowTemplate.Height = 33
        Me.DataGridView1.Size = New System.Drawing.Size(308, 429)
        Me.DataGridView1.TabIndex = 12
        '
        '结果表
        '
        Me.结果表.HeaderText = "结果表"
        Me.结果表.MinimumWidth = 9
        Me.结果表.Name = "结果表"
        Me.结果表.Width = 175
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.删除内容ToolStripMenuItem, Me.删除行ToolStripMenuItem, Me.ToolStripSeparator1, Me.全部清除ToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(169, 112)
        '
        '删除内容ToolStripMenuItem
        '
        Me.删除内容ToolStripMenuItem.Name = "删除内容ToolStripMenuItem"
        Me.删除内容ToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.删除内容ToolStripMenuItem.Text = "删除内容"
        '
        '删除行ToolStripMenuItem
        '
        Me.删除行ToolStripMenuItem.Name = "删除行ToolStripMenuItem"
        Me.删除行ToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.删除行ToolStripMenuItem.Text = "删除行"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(165, 6)
        '
        '全部清除ToolStripMenuItem
        '
        Me.全部清除ToolStripMenuItem.Name = "全部清除ToolStripMenuItem"
        Me.全部清除ToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.全部清除ToolStripMenuItem.Text = "全部清除"
        '
        'DataGridView2
        '
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.数据源表})
        Me.DataGridView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView2.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowHeadersWidth = 30
        Me.DataGridView2.RowTemplate.Height = 33
        Me.DataGridView2.Size = New System.Drawing.Size(238, 429)
        Me.DataGridView2.TabIndex = 12
        '
        '数据源表
        '
        Me.数据源表.HeaderText = "数据源表"
        Me.数据源表.MinimumWidth = 9
        Me.数据源表.Name = "数据源表"
        Me.数据源表.Width = 135
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(-5, 407)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(197, 35)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "列标题的行号："
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDown1.Location = New System.Drawing.Point(198, 410)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(93, 40)
        Me.NumericUpDown1.TabIndex = 24
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.SplitContainer1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 460.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 62.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(562, 957)
        Me.TableLayoutPanel1.TabIndex = 25
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 463)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.DataGridView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.DataGridView2)
        Me.SplitContainer1.Size = New System.Drawing.Size(556, 429)
        Me.SplitContainer1.SplitterDistance = 308
        Me.SplitContainer1.SplitterWidth = 10
        Me.SplitContainer1.TabIndex = 0
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Location = New System.Drawing.Point(3, 898)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(556, 49)
        Me.Panel1.TabIndex = 1
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.NumericUpDown1)
        Me.Panel2.Controls.Add(Me.Button4)
        Me.Panel2.Controls.Add(Me.CheckBox2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.CheckBox1)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.CheckedListBox1)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.ComboBox2)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(556, 454)
        Me.Panel2.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton6)
        Me.GroupBox1.Controls.Add(Me.RadioButton5)
        Me.GroupBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(555, 64)
        Me.GroupBox1.TabIndex = 25
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "列匹配方式"
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Checked = True
        Me.RadioButton6.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton6.Location = New System.Drawing.Point(167, 32)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(121, 32)
        Me.RadioButton6.TabIndex = 0
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "手动匹配"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton5.Location = New System.Drawing.Point(6, 32)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(121, 32)
        Me.RadioButton5.TabIndex = 0
        Me.RadioButton5.Text = "自动匹配"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        '列汇总控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "列汇总控件"
        Me.Size = New System.Drawing.Size(562, 957)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton3 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As Windows.Forms.RadioButton
    Friend WithEvents CheckedListBox1 As Windows.Forms.CheckedListBox
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents 结果表 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridView2 As Windows.Forms.DataGridView
    Friend WithEvents 数据源表 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents Panel2 As Windows.Forms.Panel
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton6 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As Windows.Forms.RadioButton
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents 删除内容ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 删除行ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As Windows.Forms.ToolStripSeparator
    Friend WithEvents 全部清除ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class
