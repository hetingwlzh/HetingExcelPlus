Public Class 列汇总控件
    Private Sub 列填充_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表(CheckedListBox1)
        加载表(ComboBox2)
    End Sub

    Public Function 加载表(ComboBox As Windows.Forms.ComboBox) As Integer




        ComboBox.Items.Clear()

        For Each sheet As Excel.Worksheet In app.Worksheets

            ComboBox.Items.Add(sheet.Name)

        Next
        Return app.Worksheets.Count
    End Function


    Public Function 加载表(ComboBox As Windows.Forms.CheckedListBox) As Integer




        ComboBox.Items.Clear()





        For Each sheet As Excel.Worksheet In app.Worksheets

            ComboBox.Items.Add(sheet.Name)

        Next
        刷新数据()
        Return app.Worksheets.Count

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        加载表(ComboBox2)
        加载表(CheckedListBox1)

    End Sub



    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        加载表(ComboBox2)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("数据表第一行必须为单行的列表头，从数据源表中搜索列表头，匹配到填充表中，且保持行不变。")
    End Sub

    Private Function 创建新的列汇总表() As Excel.Worksheet
        Dim 汇总表 As Excel.Worksheet = Module1.新建工作表("列汇总结果", True)


        Dim 填充表区域当前最大行, 数据源区域最大行 As Integer
        Dim ColumenNum As Integer = 0
        Dim 数据源表, 填充表 As Excel.Worksheet
        Dim 数据源表头, 汇总表表头, 填充表列头, 操作区域 As Excel.Range
        Dim 是否存在 As Boolean = False
        Dim num As Integer = 1
        For Each Name As Object In CheckedListBox1.CheckedItems
            数据源表 = app.Sheets(Name.ToString)
            数据源表头 = 数据源表.UsedRange.Rows(1)



            For Each sCell As Excel.Range In 数据源表头.Columns
                If sCell.Value = Nothing Then
                    Continue For
                End If
                是否存在 = False
                For Each cell As Excel.Range In 汇总表.UsedRange.Rows(1).columns
                    If cell.Value = Nothing Then
                        Continue For
                    End If
                    If cell.Value.ToString = sCell.Value.ToString Then
                        是否存在 = True
                        Exit For
                    End If
                Next
                If 是否存在 = False Then
                    汇总表.Cells(1, num).Value = sCell.Value.ToString
                    num += 1
                End If

            Next
        Next
        Return 汇总表
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If CheckedListBox1.CheckedItems.Count = 0 Then
            MsgBox("请选择 数据源表 和 填充表，不然没法干！")
            Exit Sub
        End If



        Dim 拷贝方式 As Integer
        Dim 填充表区域当前最大行, 数据源区域最大行 As Integer
        Dim ColumenNum As Integer = 0
        Dim 数据源表, 填充表 As Excel.Worksheet
        Dim 数据源区域, 填充表区域, 填充表列头, 操作区域 As Excel.Range


        If CheckBox1.Checked = True Then
            填充表 = 创建新的列汇总表()
        Else

            If ComboBox2.Text = "" Then
                MsgBox("请选择 数据源表 和 填充表，不然没法干！")
                Exit Sub
            End If
            If 是否存在工作表(ComboBox2.Text) = False Then
                MsgBox("坑我吗？这表就不存在！")
                Exit Sub
            End If
            填充表 = app.Sheets(ComboBox2.Text)
        End If





        填充表区域 = 填充表.UsedRange
        填充表列头 = 填充表区域.Rows(1)

        If RadioButton1.Checked = True Then
            拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll
        ElseIf RadioButton2.Checked = True Then
            拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues
        ElseIf RadioButton3.Checked = True Then
            拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas

        ElseIf RadioButton4.Checked = True Then
            拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats
        End If


        For Each Name As Object In CheckedListBox1.CheckedItems
            数据源表 = app.Sheets(Name.ToString)
            数据源区域 = 数据源表.UsedRange
            填充表区域 = 填充表.UsedRange
            填充表区域当前最大行 = 填充表区域.Rows.Count
            数据源区域最大行 = 数据源区域.Rows.Count

            If 数据源区域最大行 = 1 Then '只有表头没有内容
                Continue For
            End If

            For Each cell As Excel.Range In 填充表列头.Rows(1).columns
                If cell.Value Is Nothing Then
                    Continue For
                End If
                For Each sCell As Excel.Range In 数据源区域.Rows(1).columns
                    If sCell.Value Is Nothing Then
                        Continue For
                    End If
                    If cell.Value.ToString = sCell.Value.ToString Then
                        操作区域 = MyRange(数据源表, sCell.Row + 1, sCell.Column, sCell.Row + 数据源区域最大行 - 1, sCell.Column)
                        操作区域.Copy()
                        MyRange(填充表, cell.Row + 填充表区域当前最大行, cell.Column).PasteSpecial(拷贝方式)
                        ColumenNum += 1
                        Exit For
                    End If
                Next
            Next
        Next

        MsgBox("数据被汇总到:  " & 填充表.Name & vbCrLf & "共拷贝了 " & ColumenNum & " 列")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For i = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, Not CheckedListBox1.GetItemChecked(i))
        Next
        刷新数据()
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        刷新数据()
    End Sub
    Private Sub 刷新数据()
        Label3.Text = "当前数据表（共" & CheckedListBox1.Items.Count & "个,选中" & CheckedListBox1.CheckedItems.Count & "个）"

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox2.Enabled = Not CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckedListBox1.MultiColumn = CheckBox2.Checked
    End Sub
End Class
