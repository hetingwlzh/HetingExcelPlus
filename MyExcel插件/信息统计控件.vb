Public Class 信息统计控件
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Text = 统计信息()
    End Sub
    Public Sub 设置信息(信息 As String)
        TextBox1.Text = 信息
    End Sub
    Private Sub 信息统计控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = 统计信息()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        是否自动统计当前选择信息 = CheckBox1.Checked
    End Sub
End Class
