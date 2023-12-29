Public Class errorForm
    Private Sub errorForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Text = ""
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        My.Computer.Clipboard.SetText(RichTextBox1.Text)

    End Sub



    Private Sub errorForm_DoubleClick(sender As Object, e As EventArgs) Handles MyBase.DoubleClick
        My.Computer.Clipboard.SetText(RichTextBox1.Text)
        RichTextBox1.Text = ""
        Me.Hide()
    End Sub
End Class