Public Class 翻译控件
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each cell As Excel.Range In app.Selection
            If cell.Value <> Nothing Then
                cell.AddComment()
                cell.Comment.Visible = False
                cell.Comment.Text(Translate1.翻译(cell.Value))
            End If
        Next



    End Sub


End Class
