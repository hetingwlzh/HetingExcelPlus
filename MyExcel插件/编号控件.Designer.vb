<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 编号控件
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
        Me.components = New System.ComponentModel.Container()
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("编号方案序列")
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CheckedListBox2 = New System.Windows.Forms.CheckedListBox()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.多列显示ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.刷新ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.序号 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.列标题 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.值 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.TreeView1 = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.上移ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.下移ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.删除规则ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.添加编号方案ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.删除编号方案ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel4 = New System.Windows.Forms.FlowLayoutPanel()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel3 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.刷新ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip2.SuspendLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.FlowLayoutPanel4.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button4.Location = New System.Drawing.Point(75, 3)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(66, 72)
        Me.Button4.TabIndex = 24
        Me.Button4.Text = "删除规则"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 72)
        Me.Button1.TabIndex = 24
        Me.Button1.Text = "添加规则"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'CheckedListBox2
        '
        Me.CheckedListBox2.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.CheckedListBox2.CheckOnClick = True
        Me.CheckedListBox2.ContextMenuStrip = Me.ContextMenuStrip2
        Me.CheckedListBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CheckedListBox2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckedListBox2.FormattingEnabled = True
        Me.CheckedListBox2.Items.AddRange(New Object() {"常量规则", "引用规则", "标志位规则", "分类编号规则", "整体编号规则"})
        Me.CheckedListBox2.Location = New System.Drawing.Point(3, 3)
        Me.CheckedListBox2.MultiColumn = True
        Me.CheckedListBox2.Name = "CheckedListBox2"
        Me.CheckedListBox2.Size = New System.Drawing.Size(501, 193)
        Me.CheckedListBox2.TabIndex = 21
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.多列显示ToolStripMenuItem, Me.刷新ToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(169, 72)
        '
        '多列显示ToolStripMenuItem
        '
        Me.多列显示ToolStripMenuItem.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.多列显示ToolStripMenuItem.Name = "多列显示ToolStripMenuItem"
        Me.多列显示ToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.多列显示ToolStripMenuItem.Text = "多列显示"
        '
        '刷新ToolStripMenuItem
        '
        Me.刷新ToolStripMenuItem.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.刷新ToolStripMenuItem.Name = "刷新ToolStripMenuItem"
        Me.刷新ToolStripMenuItem.Size = New System.Drawing.Size(168, 34)
        Me.刷新ToolStripMenuItem.Text = "刷新"
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.CheckedListBox1.CheckOnClick = True
        Me.CheckedListBox1.ContextMenuStrip = Me.ContextMenuStrip2
        Me.CheckedListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CheckedListBox1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Location = New System.Drawing.Point(3, 202)
        Me.CheckedListBox1.MultiColumn = True
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(501, 274)
        Me.CheckedListBox1.TabIndex = 21
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.序号, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4})
        Me.DataGridView3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView3.Location = New System.Drawing.Point(3, 482)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.RowHeadersWidth = 40
        Me.DataGridView3.RowTemplate.Height = 33
        Me.DataGridView3.Size = New System.Drawing.Size(501, 444)
        Me.DataGridView3.TabIndex = 27
        '
        '序号
        '
        Me.序号.HeaderText = "序号"
        Me.序号.MinimumWidth = 9
        Me.序号.Name = "序号"
        Me.序号.ReadOnly = True
        Me.序号.Width = 80
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "值"
        Me.DataGridViewTextBoxColumn3.MinimumWidth = 9
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 300
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "编号"
        Me.DataGridViewTextBoxColumn4.MinimumWidth = 9
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 200
        '
        'Label5
        '
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label5.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Label5.Location = New System.Drawing.Point(0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(271, 150)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "方案个数,当前方案规则个数"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.列标题, Me.值})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 174)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidth = 40
        Me.DataGridView1.RowTemplate.Height = 33
        Me.DataGridView1.Size = New System.Drawing.Size(501, 771)
        Me.DataGridView1.TabIndex = 27
        '
        '列标题
        '
        Me.列标题.HeaderText = "列标题"
        Me.列标题.MinimumWidth = 9
        Me.列标题.Name = "列标题"
        Me.列标题.Width = 250
        '
        '值
        '
        Me.值.HeaderText = "值"
        Me.值.MinimumWidth = 9
        Me.值.Name = "值"
        Me.值.Width = 250
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(197, 35)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "列标题的行号："
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(206, 3)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(143, 35)
        Me.NumericUpDown1.TabIndex = 22
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(102, 3)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(259, 40)
        Me.ComboBox2.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 28)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "表   名："
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button3.Location = New System.Drawing.Point(274, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(157, 55)
        Me.Button3.TabIndex = 23
        Me.Button3.Text = "删除当前方案"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button5.Location = New System.Drawing.Point(3, 3)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(109, 55)
        Me.Button5.TabIndex = 23
        Me.Button5.Text = "添加方案"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(105, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(256, 40)
        Me.ComboBox1.TabIndex = 16
        '
        'TreeView1
        '
        Me.TreeView1.BackColor = System.Drawing.SystemColors.Info
        Me.TreeView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TreeView1.HideSelection = False
        Me.TreeView1.HotTracking = True
        Me.TreeView1.Location = New System.Drawing.Point(3, 3)
        Me.TreeView1.Name = "TreeView1"
        TreeNode2.Name = "编号方案序列"
        TreeNode2.Text = "编号方案序列"
        Me.TreeView1.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode2})
        Me.TreeView1.Size = New System.Drawing.Size(271, 902)
        Me.TreeView1.TabIndex = 27
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(28, 28)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.上移ToolStripMenuItem, Me.下移ToolStripMenuItem, Me.ToolStripSeparator3, Me.刷新ToolStripMenuItem1, Me.删除规则ToolStripMenuItem, Me.ToolStripSeparator1, Me.添加编号方案ToolStripMenuItem, Me.删除编号方案ToolStripMenuItem, Me.ToolStripSeparator2})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(211, 226)
        '
        '上移ToolStripMenuItem
        '
        Me.上移ToolStripMenuItem.Font = New System.Drawing.Font("Microsoft YaHei UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.上移ToolStripMenuItem.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.上移ToolStripMenuItem.Name = "上移ToolStripMenuItem"
        Me.上移ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.上移ToolStripMenuItem.Text = "上移"
        '
        '下移ToolStripMenuItem
        '
        Me.下移ToolStripMenuItem.Font = New System.Drawing.Font("Microsoft YaHei UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.下移ToolStripMenuItem.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.下移ToolStripMenuItem.Name = "下移ToolStripMenuItem"
        Me.下移ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.下移ToolStripMenuItem.Text = "下移"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(267, 6)
        '
        '删除规则ToolStripMenuItem
        '
        Me.删除规则ToolStripMenuItem.BackColor = System.Drawing.Color.AntiqueWhite
        Me.删除规则ToolStripMenuItem.Name = "删除规则ToolStripMenuItem"
        Me.删除规则ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.删除规则ToolStripMenuItem.Text = "删除规则"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(267, 6)
        '
        '添加编号方案ToolStripMenuItem
        '
        Me.添加编号方案ToolStripMenuItem.Name = "添加编号方案ToolStripMenuItem"
        Me.添加编号方案ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.添加编号方案ToolStripMenuItem.Text = "添加编号方案"
        '
        '删除编号方案ToolStripMenuItem
        '
        Me.删除编号方案ToolStripMenuItem.BackColor = System.Drawing.Color.AntiqueWhite
        Me.删除编号方案ToolStripMenuItem.Name = "删除编号方案ToolStripMenuItem"
        Me.删除编号方案ToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.删除编号方案ToolStripMenuItem.Text = "删除编号方案"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(267, 6)
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(112, 35)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(521, 1064)
        Me.TabControl1.TabIndex = 28
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.TableLayoutPanel3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 39)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(513, 1021)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "添加方案"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.FlowLayoutPanel4, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.DataGridView1, 0, 3)
        Me.TableLayoutPanel3.Controls.Add(Me.FlowLayoutPanel2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.FlowLayoutPanel3, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.Panel1, 0, 4)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 5
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 59.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 67.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(507, 1015)
        Me.TableLayoutPanel3.TabIndex = 28
        '
        'FlowLayoutPanel4
        '
        Me.FlowLayoutPanel4.Controls.Add(Me.Label1)
        Me.FlowLayoutPanel4.Controls.Add(Me.NumericUpDown1)
        Me.FlowLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel4.Location = New System.Drawing.Point(3, 118)
        Me.FlowLayoutPanel4.Name = "FlowLayoutPanel4"
        Me.FlowLayoutPanel4.Size = New System.Drawing.Size(501, 50)
        Me.FlowLayoutPanel4.TabIndex = 29
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.Label6)
        Me.FlowLayoutPanel2.Controls.Add(Me.ComboBox1)
        Me.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(3, 3)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(501, 53)
        Me.FlowLayoutPanel2.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(3, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 28)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "方案名："
        '
        'FlowLayoutPanel3
        '
        Me.FlowLayoutPanel3.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel3.Controls.Add(Me.ComboBox2)
        Me.FlowLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel3.Location = New System.Drawing.Point(3, 62)
        Me.FlowLayoutPanel3.Name = "FlowLayoutPanel3"
        Me.FlowLayoutPanel3.Size = New System.Drawing.Size(501, 50)
        Me.FlowLayoutPanel3.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Info
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Button7)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 951)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(501, 61)
        Me.Panel1.TabIndex = 30
        '
        'Button7
        '
        Me.Button7.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Button7.Location = New System.Drawing.Point(118, 3)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(150, 55)
        Me.Button7.TabIndex = 23
        Me.Button7.Text = "修改当前方案"
        Me.Button7.UseVisualStyleBackColor = True
        Me.Button7.Visible = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TableLayoutPanel2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 39)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(513, 1021)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "方案设置"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.FlowLayoutPanel1, 0, 3)
        Me.TableLayoutPanel2.Controls.Add(Me.CheckedListBox2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.CheckedListBox1, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.DataGridView3, 0, 2)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 4
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 21.47436!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 30.1282!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 48.39743!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 84.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(507, 1015)
        Me.TableLayoutPanel2.TabIndex = 30
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button4)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button6)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 932)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(501, 80)
        Me.FlowLayoutPanel1.TabIndex = 31
        '
        'Button6
        '
        Me.Button6.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button6.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Button6.Location = New System.Drawing.Point(147, 3)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(69, 72)
        Me.Button6.TabIndex = 26
        Me.Button6.Text = "开始编号"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TableLayoutPanel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TabControl1)
        Me.SplitContainer1.Size = New System.Drawing.Size(802, 1064)
        Me.SplitContainer1.SplitterDistance = 277
        Me.SplitContainer1.TabIndex = 30
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TreeView1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 156.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(277, 1064)
        Me.TableLayoutPanel1.TabIndex = 31
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Button9)
        Me.Panel2.Controls.Add(Me.Button8)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 911)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(271, 150)
        Me.Panel2.TabIndex = 28
        '
        'Button9
        '
        Me.Button9.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button9.Location = New System.Drawing.Point(114, 104)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(105, 42)
        Me.Button9.TabIndex = 20
        Me.Button9.Text = "导入方案"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button8.Location = New System.Drawing.Point(3, 105)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(105, 42)
        Me.Button8.TabIndex = 20
        Me.Button8.Text = "保存方案"
        Me.Button8.UseVisualStyleBackColor = True
        '
        '刷新ToolStripMenuItem1
        '
        Me.刷新ToolStripMenuItem1.Name = "刷新ToolStripMenuItem1"
        Me.刷新ToolStripMenuItem1.Size = New System.Drawing.Size(270, 34)
        Me.刷新ToolStripMenuItem1.Text = "刷新"
        '
        '编号控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "编号控件"
        Me.Size = New System.Drawing.Size(802, 1064)
        Me.ContextMenuStrip2.ResumeLayout(False)
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.FlowLayoutPanel4.ResumeLayout(False)
        Me.FlowLayoutPanel4.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.PerformLayout()
        Me.FlowLayoutPanel3.ResumeLayout(False)
        Me.FlowLayoutPanel3.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents NumericUpDown1 As Windows.Forms.NumericUpDown
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CheckedListBox1 As Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Button5 As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents DataGridView3 As Windows.Forms.DataGridView
    Friend WithEvents CheckedListBox2 As Windows.Forms.CheckedListBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents Button4 As Windows.Forms.Button
    Friend WithEvents TreeView1 As Windows.Forms.TreeView
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents 添加编号方案ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 删除编号方案ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As Windows.Forms.ToolStripSeparator
    Friend WithEvents 删除规则ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents TabControl1 As Windows.Forms.TabControl
    Friend WithEvents TabPage1 As Windows.Forms.TabPage
    Friend WithEvents TabPage2 As Windows.Forms.TabPage
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Button6 As Windows.Forms.Button
    Friend WithEvents 序号 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 列标题 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 值 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TableLayoutPanel3 As Windows.Forms.TableLayoutPanel
    Friend WithEvents FlowLayoutPanel4 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents FlowLayoutPanel2 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents FlowLayoutPanel3 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents TableLayoutPanel2 As Windows.Forms.TableLayoutPanel
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents SplitContainer1 As Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents 上移ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 下移ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator3 As Windows.Forms.ToolStripSeparator
    Friend WithEvents Button7 As Windows.Forms.Button
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents Panel2 As Windows.Forms.Panel
    Friend WithEvents Button8 As Windows.Forms.Button
    Friend WithEvents Button9 As Windows.Forms.Button
    Friend WithEvents ContextMenuStrip2 As Windows.Forms.ContextMenuStrip
    Friend WithEvents 多列显示ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 刷新ToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents 刷新ToolStripMenuItem1 As Windows.Forms.ToolStripMenuItem
End Class
