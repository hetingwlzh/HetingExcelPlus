Public Class 随机抽取控件
    Public 待抽取区域, 分类区域, 表头区域, 排序区域 As Excel.Range


    Public Function 载入待抽取区域() As Excel.Range
        待抽取区域 = app.Selection
        If 待抽取区域 Is Nothing Then
            MsgBox("请先选择 待抽取区域！")
            Label3.Text = "0个"
            Return Nothing
        Else
            TextBox2.Text = 待抽取区域.Address(False, False)
            TextBox2.BackColor = Drawing.Color.White
            Label3.Text = 获取单元格个数(待抽取区域) & "个"
            Return 待抽取区域
        End If
    End Function

    Public Function 载入分类区域() As Excel.Range

        分类区域 = app.Selection
        If 分类区域 Is Nothing Then
            MsgBox("请先选择 分类区域！")
        Label4.Text = "0个"
        Return Nothing
        Else
        TextBox1.Text = 分类区域.Address(False, False)
        TextBox1.BackColor = Drawing.Color.White
            Label4.Text = 获取单元格个数(分类区域) & "个"
            Return 分类区域
        End If
    End Function

    Public Function 载入表头区域() As Excel.Range
        表头区域 = app.Selection
        If 表头区域 Is Nothing Then
            MsgBox("请先选择 分类区域！")
            Label2.Text = "0个"
            Return Nothing
        Else
            TextBox3.Text = 表头区域.Address(False, False)
            TextBox3.BackColor = Drawing.Color.White
            Label2.Text = 获取单元格个数(表头区域) & "个"
            Return 表头区域
        End If
    End Function

    Public Function 载入排序区域() As Excel.Range
        排序区域 = app.Selection
        If 排序区域 Is Nothing Then
            MsgBox("请先选择 分类区域！")
            Label9.Text = "0个"
            Return Nothing
        Else
            TextBox4.Text = 排序区域.Address(False, False)
            TextBox4.BackColor = Drawing.Color.White
            Label9.Text = 获取单元格个数(排序区域) & "个"
            Return 排序区域
        End If
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        载入待抽取区域()
    End Sub

    Private Sub TextBox1_KeyUp(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox1.KeyUp
        Dim range As Excel.Range
        Try
            range = app.Range(TextBox1.Text)
            range = app.Intersect(range, app.ActiveSheet.UsedRange)
            TextBox1.BackColor = Drawing.Color.White
            分类区域 = range
        Catch ex As Exception
            TextBox1.BackColor = 警告色
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        载入分类区域()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        TextBox1.Enabled = CheckBox1.Checked
        Button1.Enabled = CheckBox1.Checked
        Label4.Enabled = CheckBox1.Checked
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        载入表头区域()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged, CheckBox5.CheckedChanged
        TextBox3.Enabled = CheckBox2.Checked
        Button2.Enabled = CheckBox2.Checked
        Label2.Enabled = CheckBox2.Checked
    End Sub

    Private Sub 随机抽取控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        ListBox1.Items.Clear()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        载入排序区域()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        Try
            Dim n As Integer = ListBox1.Items.Count
            If NumericUpDown1.Value > n Then
                For i As Integer = n - 1 To NumericUpDown1.Value
                    ListBox1.Items.Insert(i, i)
                Next
            ElseIf NumericUpDown1.Value < n Then
                For i As Integer = n - 1 To NumericUpDown1.Value Step -1
                    ListBox1.Items.RemoveAt(i)
                Next
            End If
        Catch ex As Exception
            MsgBox("ERROR")
        End Try



    End Sub

    Private Sub TextBox4_KeyUp(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox4.KeyUp
        Dim range As Excel.Range
        Try
            range = app.Range(TextBox4.Text)
            range = app.Intersect(range, app.ActiveSheet.UsedRange)
            TextBox4.BackColor = Drawing.Color.White
            分类区域 = range
        Catch ex As Exception
            TextBox4.BackColor = 警告色
        End Try
    End Sub

    Private Sub TextBox3_KeyUp(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox3.KeyUp
        Dim range As Excel.Range
        Try
            range = app.Range(TextBox3.Text)
            range = app.Intersect(range, app.ActiveSheet.UsedRange)
            TextBox3.BackColor = Drawing.Color.White
            分类区域 = range
        Catch ex As Exception
            TextBox3.BackColor = 警告色
        End Try
    End Sub

    Private Sub TextBox2_KeyUp(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox2.KeyUp
        Dim range As Excel.Range
        Try
            range = app.Range(TextBox2.Text)
            range = app.Intersect(range, app.ActiveSheet.UsedRange)
            TextBox2.BackColor = Drawing.Color.White
            待抽取区域 = range
        Catch ex As Exception
            TextBox2.BackColor = 警告色
        End Try
    End Sub
End Class
