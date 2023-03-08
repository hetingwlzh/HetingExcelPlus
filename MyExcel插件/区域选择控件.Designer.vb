<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 区域选择控件
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.SelectButton = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.序号 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.区域名 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.地址 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.锁行 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.锁列 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.锁表 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.锁用户区 = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lockRow = New System.Windows.Forms.CheckBox()
        Me.lockColumn = New System.Windows.Forms.CheckBox()
        Me.lockSheet = New System.Windows.Forms.CheckBox()
        Me.lockUserRange = New System.Windows.Forms.CheckBox()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.addButton = New System.Windows.Forms.Button()
        Me.delButton = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'SelectButton
        '
        Me.SelectButton.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.SelectButton.ForeColor = System.Drawing.SystemColors.Desktop
        Me.SelectButton.Location = New System.Drawing.Point(259, 3)
        Me.SelectButton.Name = "SelectButton"
        Me.SelectButton.Size = New System.Drawing.Size(122, 38)
        Me.SelectButton.TabIndex = 0
        Me.SelectButton.Text = "选择区域"
        Me.SelectButton.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.序号, Me.区域名, Me.地址, Me.锁行, Me.锁列, Me.锁表, Me.锁用户区})
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("微软雅黑", 10.0!)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.DataGridView1.RowTemplate.Height = 33
        Me.DataGridView1.Size = New System.Drawing.Size(636, 337)
        Me.DataGridView1.TabIndex = 3
        '
        '序号
        '
        Me.序号.HeaderText = "序号"
        Me.序号.MinimumWidth = 9
        Me.序号.Name = "序号"
        Me.序号.ReadOnly = True
        Me.序号.Width = 93
        '
        '区域名
        '
        Me.区域名.HeaderText = "区域名"
        Me.区域名.MinimumWidth = 9
        Me.区域名.Name = "区域名"
        Me.区域名.Width = 114
        '
        '地址
        '
        Me.地址.HeaderText = "地址"
        Me.地址.MinimumWidth = 9
        Me.地址.Name = "地址"
        Me.地址.Width = 93
        '
        '锁行
        '
        Me.锁行.HeaderText = "锁行"
        Me.锁行.MinimumWidth = 9
        Me.锁行.Name = "锁行"
        Me.锁行.ReadOnly = True
        Me.锁行.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.锁行.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.锁行.Width = 93
        '
        '锁列
        '
        Me.锁列.HeaderText = "锁列"
        Me.锁列.MinimumWidth = 9
        Me.锁列.Name = "锁列"
        Me.锁列.Width = 58
        '
        '锁表
        '
        Me.锁表.HeaderText = "锁表"
        Me.锁表.MinimumWidth = 9
        Me.锁表.Name = "锁表"
        Me.锁表.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.锁表.Width = 58
        '
        '锁用户区
        '
        Me.锁用户区.HeaderText = "锁用户区"
        Me.锁用户区.MinimumWidth = 9
        Me.锁用户区.Name = "锁用户区"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.Control
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label1, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel2, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(642, 461)
        Me.TableLayoutPanel1.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微软雅黑", 7.714286!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Label1.Location = New System.Drawing.Point(3, 437)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 24)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "校验信息"
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.AutoSize = True
        Me.FlowLayoutPanel2.Controls.Add(Me.lockRow)
        Me.FlowLayoutPanel2.Controls.Add(Me.lockColumn)
        Me.FlowLayoutPanel2.Controls.Add(Me.lockSheet)
        Me.FlowLayoutPanel2.Controls.Add(Me.lockUserRange)
        Me.FlowLayoutPanel2.Location = New System.Drawing.Point(3, 396)
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        Me.FlowLayoutPanel2.Size = New System.Drawing.Size(470, 38)
        Me.FlowLayoutPanel2.TabIndex = 5
        '
        'lockRow
        '
        Me.lockRow.AutoSize = True
        Me.lockRow.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lockRow.Location = New System.Drawing.Point(3, 3)
        Me.lockRow.Name = "lockRow"
        Me.lockRow.Size = New System.Drawing.Size(101, 32)
        Me.lockRow.TabIndex = 5
        Me.lockRow.Text = "锁定行"
        Me.lockRow.UseVisualStyleBackColor = True
        '
        'lockColumn
        '
        Me.lockColumn.AutoSize = True
        Me.lockColumn.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lockColumn.Location = New System.Drawing.Point(110, 3)
        Me.lockColumn.Name = "lockColumn"
        Me.lockColumn.Size = New System.Drawing.Size(101, 32)
        Me.lockColumn.TabIndex = 6
        Me.lockColumn.Text = "锁定列"
        Me.lockColumn.UseVisualStyleBackColor = True
        '
        'lockSheet
        '
        Me.lockSheet.AutoSize = True
        Me.lockSheet.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lockSheet.Location = New System.Drawing.Point(217, 3)
        Me.lockSheet.Name = "lockSheet"
        Me.lockSheet.Size = New System.Drawing.Size(101, 32)
        Me.lockSheet.TabIndex = 7
        Me.lockSheet.Text = "锁定表"
        Me.lockSheet.UseVisualStyleBackColor = True
        '
        'lockUserRange
        '
        Me.lockUserRange.AutoSize = True
        Me.lockUserRange.Checked = True
        Me.lockUserRange.CheckState = System.Windows.Forms.CheckState.Checked
        Me.lockUserRange.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.lockUserRange.Location = New System.Drawing.Point(324, 3)
        Me.lockUserRange.Name = "lockUserRange"
        Me.lockUserRange.Size = New System.Drawing.Size(143, 32)
        Me.lockUserRange.TabIndex = 7
        Me.lockUserRange.Text = "锁定用户区"
        Me.lockUserRange.UseVisualStyleBackColor = True
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoSize = True
        Me.FlowLayoutPanel1.Controls.Add(Me.addButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.delButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.SelectButton)
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 346)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(473, 44)
        Me.FlowLayoutPanel1.TabIndex = 5
        '
        'addButton
        '
        Me.addButton.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.addButton.Location = New System.Drawing.Point(3, 3)
        Me.addButton.Name = "addButton"
        Me.addButton.Size = New System.Drawing.Size(122, 38)
        Me.addButton.TabIndex = 5
        Me.addButton.Text = "添加区域"
        Me.addButton.UseVisualStyleBackColor = True
        '
        'delButton
        '
        Me.delButton.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.delButton.Location = New System.Drawing.Point(131, 3)
        Me.delButton.Name = "delButton"
        Me.delButton.Size = New System.Drawing.Size(122, 38)
        Me.delButton.TabIndex = 5
        Me.delButton.Text = "删除区域"
        Me.delButton.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.Desktop
        Me.Button1.Location = New System.Drawing.Point(387, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(83, 38)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "清除"
        Me.Button1.UseVisualStyleBackColor = True
        '
        '区域选择控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "区域选择控件"
        Me.Size = New System.Drawing.Size(642, 461)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.FlowLayoutPanel2.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SelectButton As Windows.Forms.Button
    Friend WithEvents DataGridView1 As Windows.Forms.DataGridView
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents addButton As Windows.Forms.Button
    Friend WithEvents FlowLayoutPanel2 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents lockRow As Windows.Forms.CheckBox
    Friend WithEvents lockColumn As Windows.Forms.CheckBox
    Friend WithEvents lockSheet As Windows.Forms.CheckBox
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents lockUserRange As Windows.Forms.CheckBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents 序号 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 区域名 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 地址 As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 锁行 As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents 锁列 As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents 锁表 As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents 锁用户区 As Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents delButton As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
