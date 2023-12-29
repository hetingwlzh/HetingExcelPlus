<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 翻译控件
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Translate1 = New HetingControl.translate()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("微软雅黑", 10.71429!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button1.Location = New System.Drawing.Point(3, 3)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(213, 48)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "翻译所选单元格"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Translate1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.FlowLayoutPanel1, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 57.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(384, 434)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Translate1
        '
        Me.Translate1.AutoSize = True
        Me.Translate1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Translate1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Translate1.Location = New System.Drawing.Point(3, 3)
        Me.Translate1.MinimumSize = New System.Drawing.Size(50, 50)
        Me.Translate1.Name = "Translate1"
        Me.Translate1.Size = New System.Drawing.Size(378, 371)
        Me.Translate1.TabIndex = 3
        Me.Translate1.字体 = New System.Drawing.Font("微软雅黑", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Translate1.拆分方向 = 0
        Me.Translate1.颜色 = System.Drawing.SystemColors.WindowText
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.Button1)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 380)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(378, 51)
        Me.FlowLayoutPanel1.TabIndex = 4
        '
        '翻译控件
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(168.0!, 168.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "翻译控件"
        Me.Size = New System.Drawing.Size(384, 434)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents Translate1 As HetingControl.translate
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
End Class
