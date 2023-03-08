Public Class SelectRangeForm
    Private Sub SelectRangeForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Me.Visible = True
        Dim Common1 As New Common '删除关闭按钮的类实例
        Common1.DisableCloseButton(Me) '删除关闭按钮
        加载工作表()
    End Sub

    Public Function 获取地址(Optional range As Excel.Range = Nothing) As String
        If range Is Nothing Then
            range = app.Selection
        End If
        Try
            If range IsNot Nothing Then
                Dim 地址格式 As Integer
                If RadioButton1.Checked = True And RadioButton2.Checked = False Then
                    地址格式 = Excel.XlReferenceStyle.xlA1
                ElseIf RadioButton1.Checked = False And RadioButton2.Checked = True Then
                    地址格式 = Excel.XlReferenceStyle.xlR1C1
                End If




                Dim adress As String = range.Address(Not CheckBox1.Checked,
                                                     Not CheckBox2.Checked,
                                                     地址格式, CheckBox3.Checked, 0)

                If CheckBox3.Checked = True Then
                    adress = adress.Replace(",", "," & range.Worksheet.Name & "!")
                End If

                Return adress
            End If

        Catch ex As Exception

        End Try

    End Function

    Public Function 取并集(Adress As String) As String
        Dim range As Excel.Range = 地址转区域(Adress)
        range = app.Union(range, range)
        Return 获取地址(range)

    End Function


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged, CheckBox3.CheckedChanged, RadioButton1.CheckedChanged
        TextBox1.Text = 获取地址()
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = 取并集(TextBox1.Text)
        Dim range As Excel.Range = 地址转区域(TextBox1.Text)
        TextBox1.Text &= 列表转字符串(区域转序列(range, True))
    End Sub
    Public Function 地址转区域(Address As String) As Excel.Range
        'Dim sheet As Excel.Worksheet = app.ActiveSheet
        Return app.Range(Address)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            地址转区域(TextBox1.Text).Select()
        Catch ex As Exception
            MsgBox("地址字符串有误！")
        End Try

    End Sub

    Private Sub SelectRangeForm_DoubleClick(sender As Object, e As EventArgs) Handles MyBase.DoubleClick
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            当前选区 = 地址转区域(TextBox1.Text)
            地址单元格.Value = TextBox1.Text
            自动列宽(地址单元格.Worksheet)
            自动行高(地址单元格.Worksheet)
            地址单元格.Worksheet.Activate()
            地址单元格 = Nothing
            TextBox1.Text = ""
            Me.Hide()
        Catch ex As Exception
            MsgBox("地址字符串有误！请重新输入。")
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Label1.Text = "单元格数：" & 获取单元格个数(地址转区域(TextBox1.Text))

        Catch ex As Exception
            Label1.Text = "地址有误！"
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try

            app.Worksheets(ComboBox1.Text).activate
            MySelectForm.TextBox1.Text = MySelectForm.获取地址()
        Catch ex As Exception
            MsgBox("找不到指定的表")
        End Try
    End Sub
    Public Function 加载工作表() As Integer
        Try
            ComboBox1.Items.Clear()
            For Each sheet As Excel.Worksheet In app.Worksheets
                ComboBox1.Items.Add(sheet.Name)
            Next
            ComboBox1.Text = app.ActiveSheet.Name
            Return app.Worksheets.Count
        Catch ex As Exception
            'MsgBox("加载工作表时错误！")
        End Try


    End Function

    Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        ComboBox1.Items.Clear()
        For Each sheet As Excel.Worksheet In app.Worksheets
            ComboBox1.Items.Add(sheet.Name)
        Next
        ComboBox1.Text = app.ActiveSheet.Name
    End Sub



    Private Sub SelectRangeForm_MouseEnter(sender As Object, e As EventArgs) Handles MyBase.MouseEnter, TextBox1.MouseEnter
        MySelectForm.TextBox1.Text = MySelectForm.获取地址()
    End Sub
End Class