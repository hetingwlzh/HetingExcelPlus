Public Class 枚举控件
    Private Sub 枚举控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        RadioButton1.Checked = My.Settings.是否先行后列
        RadioButton2.Checked = Not My.Settings.是否先行后列

        CheckBox1.Checked = My.Settings.是否剔除重复值
        CheckBox2.Checked = My.Settings.是否忽略空值
        CheckBox3.Checked = My.Settings.是否忽略首尾空白字符


        NumericUpDown1.Enabled = RadioButton7.Checked
        NumericUpDown2.Enabled = Not RadioButton7.Checked
    End Sub

    Public Sub 更改设置()
        My.Settings.是否先行后列 = RadioButton1.Checked
        My.Settings.是否剔除重复值 = CheckBox1.Checked
        My.Settings.是否忽略空值 = CheckBox2.Checked
        My.Settings.是否忽略首尾空白字符 = CheckBox3.Checked
        My.Settings.Save()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        当前选区 = 区域交集(app.Selection, app.Selection.Worksheet.UsedRange)
        My.Settings.是否先行后列 = RadioButton1.Checked
        My.Settings.是否剔除重复值 = CheckBox1.Checked
        My.Settings.是否忽略空值 = CheckBox2.Checked
        My.Settings.是否忽略首尾空白字符 = CheckBox3.Checked









        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码



        Dim NumOfRow, NumOfColumn, Num As Integer



        NumOfRow = NumericUpDown1.Value
        NumOfColumn = NumericUpDown2.Value








        Dim sheet As Excel.Worksheet
        Dim TempList As System.Collections.ArrayList '= 枚举选区(当前选区内容序列)
        sheet = 获取缓冲表() '新建工作表("枚举结果")
        Dim Count As Integer = 1
        Dim CurrentRow, CurrentColumn, StartRow, StartColumn As Integer





        StartRow = 2
        StartColumn = 2
        CurrentRow = StartRow
        CurrentColumn = StartColumn
        TempList = 枚举选区(当前选区, My.Settings.是否先行后列, My.Settings.是否剔除重复值, My.Settings.是否忽略空值, My.Settings.是否忽略首尾空白字符)
        sheet.Cells(1, 1).value = "共" & TempList.Count & "个"
        If RadioButton7.Checked = True Then '按照行数枚举
            Num = Math.Ceiling(TempList.Count / NumOfRow)
            For Each t In TempList
                If CurrentColumn > Num + StartColumn - 1 Then
                    CurrentRow += 1
                    CurrentColumn = StartColumn
                End If
                sheet.Cells(CurrentRow, CurrentColumn).value = t
                CurrentColumn += 1
            Next


            For i As Integer = 2 To NumOfRow + 1
                sheet.Cells(i, 1).value = i - 1
            Next
            For i As Integer = 2 To Num + 1
                sheet.Cells(1, i).value = i - 1
            Next
        ElseIf RadioButton8.Checked = True Then '按照列数枚举
            Num = Math.Ceiling(TempList.Count / NumOfColumn)
            For Each t In TempList
                If CurrentRow > Num + StartRow - 1 Then
                    CurrentColumn += 1
                    CurrentRow = StartRow
                End If
                sheet.Cells(CurrentRow, CurrentColumn).value = t
                CurrentRow += 1
            Next



            For i As Integer = 2 To NumOfColumn + 1
                sheet.Cells(1, i).value = i - 1
            Next
            For i As Integer = 2 To Num + 1
                sheet.Cells(i, 1).value = i - 1
            Next

        Else
            错误信息序列.Add("行列数设置错误！")
        End If


        'sheet.Cells(1, 1).value = "序号"
        'sheet.Cells(1, 2).value = "枚举结果(" & Count - 1 & "个)"


        居中(sheet.Cells)
        设置为表头样式(sheet, 1, 1)

        '设置为表头样式(sheet, 1, 1)
        'MsgBox("共枚举 " & num & " 对象" & vbCrLf & 列表转字符串(枚举区选区(), " "))

        创建结果页并显示结果("枚举结果", sheet)


        'sheet.Copy(app.Sheets(1))
        'Dim resultSheet As Excel.Worksheet = app.Sheets(1)
        ''MsgBox(resultSheet.Name)
        'If 是否存在工作表("枚举结果") Then
        '    app.DisplayAlerts = False '删除工作表不再提示
        '    app.Sheets("枚举结果").Delete() '删除工作表

        'End If
        'resultSheet.Name = "枚举结果"
        'resultSheet.Visible = True

        'resultSheet.Activate()


















    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        NumericUpDown1.Enabled = RadioButton7.Checked
        NumericUpDown2.Enabled = Not RadioButton7.Checked
    End Sub











    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        更改设置()
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        更改设置()
    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        更改设置()
    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        更改设置()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        更改设置()
    End Sub


End Class
