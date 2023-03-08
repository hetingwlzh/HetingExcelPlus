Public Class MyForm


    Public Sub 自动调整窗口位置()

        Me.Left = Math.Min(Me.Left, My.Computer.Screen.WorkingArea.Width - Me.Width)
        'Me.Top = My.Computer.Screen.WorkingArea.Height * 0.2
        Me.TopMost = True

    End Sub


    Private Sub MyForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = True
        Dim Common1 As New Common '删除关闭按钮的类实例
        Common1.DisableCloseButton(Me) '删除关闭按钮
        Me.Left = My.Computer.Screen.WorkingArea.Width - Me.Width
        Me.Top = My.Computer.Screen.WorkingArea.Height * 0.1
        自动调整窗口位置()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MyForm1.TabControl1.TabPages.Clear()

        My.Settings.Save()
        Me.Hide()
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        更新大小()
        自动调整窗口位置()
    End Sub

    Public Sub 更新大小()
        If TabControl1.Controls.Count > 0 Then
            TabControl1.ItemSize = New Drawing.Size(30, 25)

            MyForm1.Height = TabControl1.SelectedTab.Controls.Item(0).Height +
                FlowLayoutPanel1.Height +
                TabControl1.ItemSize.Height + 75

            MyForm1.Width = TabControl1.SelectedTab.Controls.Item(0).Width + 40



            Me.Left = Math.Min(Me.Left, My.Computer.Screen.WorkingArea.Width - Me.Width)
        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As Windows.Forms.ControlEventArgs)

    End Sub



    Private Sub TabControl1_DoubleClick(sender As Object, e As EventArgs) Handles TabControl1.DoubleClick
        If MyForm1.TabControl1.TabPages.Count > 1 Then
            Dim n As Integer = MyForm1.TabControl1.SelectedIndex

            Dim tab As Windows.Forms.TabPage = MyForm1.TabControl1.SelectedTab
            tab.Controls.Clear()
            MyForm1.TabControl1.TabPages.Remove(tab)
            MyForm1.TabControl1.SelectedIndex = Math.Max(n - 1, MyForm1.TabControl1.TabPages.Count - 1)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        显示插件设置()
    End Sub
End Class