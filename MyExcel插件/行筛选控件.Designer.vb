<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 行筛选控件
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.列标题 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.关系 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.值 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton8 = New System.Windows.Forms.RadioButton()
        Me.RadioButton7 = New System.Windows.Forms.RadioButton()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.RadioButton11 = New System.Windows.Forms.RadioButton()
        Me.RadioButton9 = New System.Windows.Forms.RadioButton()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.RadioButton15 = New System.Windows.Forms.RadioButton()
        Me.RadioButton16 = New System.Windows.Forms.RadioButton()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.RadioButton17 = New System.Windows.Forms.RadioButton()
        Me.RadioButton18 = New System.Windows.Forms.RadioButton()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 6)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 36)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "表名："
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(108, 0)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(228, 44)
        Me.ComboBox2.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(342, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(183, 36)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "标题的行号："
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NumericUpDown1.Location = New System.Drawing.Point(529, 4)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(102, 44)
        Me.NumericUpDown1.TabIndex = 24
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.列标题, Me.关系, Me.值})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 392)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 40
        Me.DataGridView1.RowTemplate.Height = 33
        Me.DataGridView1.Size = New System.Drawing.Size(658, 299)
        Me.DataGridView1.TabIndex = 28
        '
        '列标题
        '
        Me.列标题.HeaderText = "列标题"
        Me.列标题.MinimumWidth = 9
        Me.列标题.Name = "列标题"
        Me.列标题.Width = 250
        '
        '关系
        '
        Me.关系.HeaderText = "关系"
        Me.关系.MinimumWidth = 9
        Me.关系.Name = "关系"
        Me.关系.Width = 175
        '
        '值
        '
        Me.值.HeaderText = "值"
        Me.值.MinimumWidth = 9
        Me.值.Name = "值"
        Me.值.Width = 175
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Silver
        Me.GroupBox1.Controls.Add(Me.RadioButton8)
        Me.GroupBox1.Controls.Add(Me.RadioButton7)
        Me.GroupBox1.Controls.Add(Me.RadioButton6)
        Me.GroupBox1.Controls.Add(Me.RadioButton5)
        Me.GroupBox1.Controls.Add(Me.RadioButton4)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(140, 279)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "关系"
        '
        'RadioButton8
        '
        Me.RadioButton8.AutoSize = True
        Me.RadioButton8.Font = New System.Drawing.Font("微软雅黑", 14.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton8.Location = New System.Drawing.Point(79, 201)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(66, 49)
        Me.RadioButton8.TabIndex = 0
        Me.RadioButton8.TabStop = True
        Me.RadioButton8.Text = "⊇"
        Me.RadioButton8.UseVisualStyleBackColor = True
        '
        'RadioButton7
        '
        Me.RadioButton7.AutoSize = True
        Me.RadioButton7.Font = New System.Drawing.Font("微软雅黑", 14.14286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton7.Location = New System.Drawing.Point(19, 201)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(66, 49)
        Me.RadioButton7.TabIndex = 0
        Me.RadioButton7.TabStop = True
        Me.RadioButton7.Text = "⊆"
        Me.RadioButton7.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton6.Location = New System.Drawing.Point(79, 147)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(61, 40)
        Me.RadioButton6.TabIndex = 0
        Me.RadioButton6.TabStop = True
        Me.RadioButton6.Text = "≥"
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton5.Location = New System.Drawing.Point(19, 147)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(61, 40)
        Me.RadioButton5.TabIndex = 0
        Me.RadioButton5.TabStop = True
        Me.RadioButton5.Text = "≤"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton4.Location = New System.Drawing.Point(79, 93)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(68, 40)
        Me.RadioButton4.TabIndex = 0
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "＞"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton3.Location = New System.Drawing.Point(19, 93)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(68, 40)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "＜"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(79, 39)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(61, 40)
        Me.RadioButton2.TabIndex = 0
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "≠"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(19, 39)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(61, 40)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "="
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Silver
        Me.GroupBox2.Controls.Add(Me.TextBox2)
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.RadioButton15)
        Me.GroupBox2.Controls.Add(Me.RadioButton16)
        Me.GroupBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(149, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(498, 195)
        Me.GroupBox2.TabIndex = 29
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "值"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.TextBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(270, 121)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox2.Size = New System.Drawing.Size(212, 66)
        Me.TextBox2.TabIndex = 1
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.GroupBox4.Controls.Add(Me.RadioButton11)
        Me.GroupBox4.Controls.Add(Me.RadioButton9)
        Me.GroupBox4.Controls.Add(Me.RadioButton10)
        Me.GroupBox4.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(6, 33)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(126, 154)
        Me.GroupBox4.TabIndex = 29
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "类型"
        '
        'RadioButton11
        '
        Me.RadioButton11.AutoSize = True
        Me.RadioButton11.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton11.Location = New System.Drawing.Point(6, 115)
        Me.RadioButton11.Name = "RadioButton11"
        Me.RadioButton11.Size = New System.Drawing.Size(118, 39)
        Me.RadioButton11.TabIndex = 0
        Me.RadioButton11.TabStop = True
        Me.RadioButton11.Text = "字符串"
        Me.RadioButton11.UseVisualStyleBackColor = True
        '
        'RadioButton9
        '
        Me.RadioButton9.AutoSize = True
        Me.RadioButton9.Checked = True
        Me.RadioButton9.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton9.Location = New System.Drawing.Point(6, 30)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(92, 39)
        Me.RadioButton9.TabIndex = 0
        Me.RadioButton9.TabStop = True
        Me.RadioButton9.Text = "自动"
        Me.RadioButton9.UseVisualStyleBackColor = True
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton10.Location = New System.Drawing.Point(6, 70)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(92, 39)
        Me.RadioButton10.TabIndex = 0
        Me.RadioButton10.Text = "数值"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.TextBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(270, 39)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(212, 66)
        Me.TextBox1.TabIndex = 1
        '
        'RadioButton15
        '
        Me.RadioButton15.AutoSize = True
        Me.RadioButton15.Checked = True
        Me.RadioButton15.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton15.Location = New System.Drawing.Point(138, 63)
        Me.RadioButton15.Name = "RadioButton15"
        Me.RadioButton15.Size = New System.Drawing.Size(92, 39)
        Me.RadioButton15.TabIndex = 0
        Me.RadioButton15.TabStop = True
        Me.RadioButton15.Text = "常量"
        Me.RadioButton15.UseVisualStyleBackColor = True
        '
        'RadioButton16
        '
        Me.RadioButton16.AutoSize = True
        Me.RadioButton16.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton16.Location = New System.Drawing.Point(138, 135)
        Me.RadioButton16.Name = "RadioButton16"
        Me.RadioButton16.Size = New System.Drawing.Size(118, 39)
        Me.RadioButton16.TabIndex = 0
        Me.RadioButton16.Text = "单元格"
        Me.RadioButton16.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(9, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(171, 57)
        Me.Button1.TabIndex = 30
        Me.Button1.Text = "添加条件"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(3, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(155, 36)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "筛选条件："
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Silver
        Me.GroupBox3.Controls.Add(Me.RadioButton17)
        Me.GroupBox3.Controls.Add(Me.RadioButton18)
        Me.GroupBox3.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(149, 204)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(498, 78)
        Me.GroupBox3.TabIndex = 29
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "逻辑"
        '
        'RadioButton17
        '
        Me.RadioButton17.AutoSize = True
        Me.RadioButton17.Checked = True
        Me.RadioButton17.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton17.Location = New System.Drawing.Point(12, 27)
        Me.RadioButton17.Name = "RadioButton17"
        Me.RadioButton17.Size = New System.Drawing.Size(99, 39)
        Me.RadioButton17.TabIndex = 0
        Me.RadioButton17.TabStop = True
        Me.RadioButton17.Text = "AND"
        Me.RadioButton17.UseVisualStyleBackColor = True
        '
        'RadioButton18
        '
        Me.RadioButton18.AutoSize = True
        Me.RadioButton18.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton18.Location = New System.Drawing.Point(138, 27)
        Me.RadioButton18.Name = "RadioButton18"
        Me.RadioButton18.Size = New System.Drawing.Size(78, 39)
        Me.RadioButton18.TabIndex = 0
        Me.RadioButton18.Text = "OR"
        Me.RadioButton18.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button2.Location = New System.Drawing.Point(194, 3)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(171, 57)
        Me.Button2.TabIndex = 30
        Me.Button2.Text = "删除条件"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("微软雅黑", 14.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Green
        Me.Button3.Location = New System.Drawing.Point(395, 71)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(171, 57)
        Me.Button3.TabIndex = 30
        Me.Button3.Text = "筛选内容"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(170, Byte), Integer), CType(CType(200, Byte), Integer), CType(CType(172, Byte), Integer))
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Location = New System.Drawing.Point(3, 888)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(658, 131)
        Me.RichTextBox1.TabIndex = 31
        Me.RichTextBox1.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("微软雅黑", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 847)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(155, 36)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "筛选结果："
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("微软雅黑", 14.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Green
        Me.Button4.Location = New System.Drawing.Point(9, 71)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(171, 57)
        Me.Button4.TabIndex = 30
        Me.Button4.Text = "筛选数目"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.RichTextBox1, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel3, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label4, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel4, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 7
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 57.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 289.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 43.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 305.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 153.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(664, 1022)
        Me.TableLayoutPanel1.TabIndex = 32
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.NumericUpDown1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.ComboBox2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(658, 51)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.GroupBox2)
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 60)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(658, 283)
        Me.Panel2.TabIndex = 1
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Button5)
        Me.Panel3.Controls.Add(Me.Button4)
        Me.Panel3.Controls.Add(Me.Button3)
        Me.Panel3.Controls.Add(Me.Button6)
        Me.Panel3.Controls.Add(Me.Button2)
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(3, 697)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(658, 147)
        Me.Panel3.TabIndex = 32
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("微软雅黑", 14.14286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Green
        Me.Button5.Location = New System.Drawing.Point(194, 71)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(171, 57)
        Me.Button5.TabIndex = 30
        Me.Button5.Text = "筛选区域"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button6.Location = New System.Drawing.Point(395, 0)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(171, 57)
        Me.Button6.TabIndex = 30
        Me.Button6.Text = "说 明"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.CheckBox1)
        Me.Panel4.Controls.Add(Me.Label3)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(3, 349)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(658, 37)
        Me.Panel4.TabIndex = 33
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(164, 1)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(140, 36)
        Me.CheckBox1.TabIndex = 24
        Me.CheckBox1.Text = "忽略空值"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        '行筛选控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "行筛选控件"
        Me.Size = New System.Drawing.Size(664, 1022)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton8 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton7 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton6 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents RadioButton15 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton16 As Windows.Forms.RadioButton
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents 列标题 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 关系 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 值 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton17 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton18 As Windows.Forms.RadioButton
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents RadioButton9 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton10 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton11 As Windows.Forms.RadioButton
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents RichTextBox1 As Windows.Forms.RichTextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents Panel2 As Windows.Forms.Panel
    Friend WithEvents Panel3 As Windows.Forms.Panel
    Friend WithEvents Button5 As Windows.Forms.Button
    Friend WithEvents Button6 As Windows.Forms.Button
    Friend WithEvents Panel4 As Windows.Forms.Panel
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
End Class
