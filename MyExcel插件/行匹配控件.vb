Public Class 行匹配
    Public 填充表, 数据源表 As Excel.Worksheet
    Public 匹配列号序列 As New System.Collections.ArrayList
    Public 结果列号序列 As New System.Collections.ArrayList


    Private Sub 行汇总_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表(ComboBox1)
        匹配列号序列.Clear()
    End Sub



    Public Function 加载表(ComboBox As Windows.Forms.ComboBox) As Integer
        ComboBox.Items.Clear()

        For Each sheet As Excel.Worksheet In app.Worksheets

            ComboBox.Items.Add(sheet.Name)

        Next
        Return app.Worksheets.Count
    End Function


    Public Sub 加载列标题(sheet As Excel.Worksheet, 标题行行号 As Integer, ComboBox As Windows.Forms.CheckedListBox)
        If sheet IsNot Nothing Then
            ComboBox.Items.Clear()
            Dim str As String = ""

            If 标题行行号 > sheet.UsedRange.Rows.Count Then
                MsgBox("标题行行号超出使用区范围，请重新设置标题行所在区域的行号。")
                Exit Sub
            End If

            For Each cell As Excel.Range In sheet.UsedRange.Rows.Item(标题行行号).Cells
                ComboBox.Items.Add(ToStr(cell))
            Next
        End If




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        加载表(ComboBox1)
        加载表(ComboBox2)

        加载列标题(填充表, NumericUpDown1.Value, CheckedListBox1)
        加载列标题(数据源表, NumericUpDown2.Value, CheckedListBox2)


    End Sub



    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        加载表(sender)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("这里还在施工
请绕行"， MsgBoxStyle.ApplicationModal, "功能说明")
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




    Public Function 按行匹配(填充表 As Excel.Worksheet, 填充表列标题所在区域行号 As Integer,
                         数据源表 As Excel.Worksheet, 数据源表列标题所在区域行号 As Integer,
                         columnsList As System.Collections.ArrayList,
                         ResultColumnList As System.Collections.ArrayList,
                         Optional 未匹配的显示值 As String = "")

        If columnsList.Count = 0 Then
            MsgBox("按行匹配  你还没有选择 每行中要匹配的条件所在的列。")
            Exit Function
        End If
        If ResultColumnList.Count = 0 Then
            MsgBox("按行匹配  你还没有选择 每行中要匹配的结果所在的列。")
            Exit Function
        End If

        Dim 数据源表开始单元格 As Excel.Range = 获取用户区开始单元格(数据源表)
        Dim 数据源表结束单元格 As Excel.Range = 获取结束单元格(数据源表)

        Dim 填充表开始单元格 As Excel.Range = 获取用户区开始单元格(填充表)
        Dim 填充表结束单元格 As Excel.Range = 获取结束单元格(填充表)

        Dim tempStr As String = ""

        Dim 辅助列 As Excel.Range = 插入列(数据源表, 数据源表结束单元格.Column + 1)
        For Each t As 数对 In columnsList
            tempStr &= MyRange(数据源表, 数据源表开始单元格.Row + 数据源表列标题所在区域行号, 数据源表开始单元格.Column + t.y - 1).Address(False, True) & "&"
        Next

        If tempStr.EndsWith("&") Then
            tempStr = tempStr.Substring(0, tempStr.Length - 1)
        End If

        MyRange(数据源表, 数据源表开始单元格.Row + 数据源表列标题所在区域行号 - 1, 数据源表结束单元格.Column + 1).Value = "辅助列"
        插入公式("=" & tempStr, MyRange(数据源表, 数据源表开始单元格.Row + 数据源表列标题所在区域行号, 数据源表结束单元格.Column + 1))
        拖拽填充(MyRange(数据源表, 数据源表开始单元格.Row + 数据源表列标题所在区域行号, 数据源表结束单元格.Column + 1, 数据源表结束单元格.Row, 数据源表结束单元格.Column + 1))









        tempStr = ""

        For Each t As 数对 In columnsList
            tempStr &= MyRange(填充表, 填充表开始单元格.Row + 填充表列标题所在区域行号, 填充表开始单元格.Column + t.x - 1).Address(False, True) & "&"
        Next

        If tempStr.EndsWith("&") Then
            tempStr = tempStr.Substring(0, tempStr.Length - 1)
        End If



        For Each column As Integer In ResultColumnList
            插入列(填充表, 填充表结束单元格.Column + 1, False)


            MyRange(填充表, 填充表开始单元格.Row + 填充表列标题所在区域行号 - 1, 填充表结束单元格.Column + 1).Value = "#" & 数据源表.UsedRange(数据源表列标题所在区域行号, column).value

            插入公式("=IFERROR(INDEX('" & 数据源表.Name & "'!" & 数据源表.UsedRange.Address &
             ",MATCH(" & tempStr & ",'" & 数据源表.Name & "'!" & app.Intersect(辅助列, 数据源表.UsedRange).Address & ",0)," & column & ")," & """" & 未匹配的显示值 & """" & ")",
             MyRange(填充表, 填充表开始单元格.Row + 填充表列标题所在区域行号, 填充表结束单元格.Column + 1))

            拖拽填充(MyRange(填充表, 填充表开始单元格.Row + 填充表列标题所在区域行号, 填充表结束单元格.Column + 1, 填充表结束单元格.Row, 填充表结束单元格.Column + 1))

            填充表开始单元格 = 获取用户区开始单元格(填充表)
            填充表结束单元格 = 获取结束单元格(填充表)
        Next





    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        按行匹配(填充表, NumericUpDown1.Value, 数据源表, NumericUpDown2.Value, 匹配列号序列, 结果列号序列, TextBox1.Text)



    End Sub



    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged, CheckedListBox2.SelectedIndexChanged, CheckedListBox3.SelectedIndexChanged, CheckedListBox4.SelectedIndexChanged
        Dim n As Integer = sender.SelectedIndex
        If n <> -1 Then
            For i = 0 To sender.Items.Count - 1
                sender.SetItemChecked(i, False)
            Next
            sender.SetItemChecked(n, True)
        End If

    End Sub
    Private Sub 刷新数据()
        Label3.Text = "当前数据表（共" & CheckedListBox1.Items.Count & "个,选中" & CheckedListBox1.CheckedItems.Count & "个）"

    End Sub



    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckedListBox1.MultiColumn = CheckBox2.Checked
        CheckedListBox2.MultiColumn = CheckBox2.Checked
    End Sub

    Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        加载表(ComboBox2)
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim 填充表列号, 数据源表列号 As Integer
        填充表列号 = CheckedListBox1.SelectedIndex + 1
        数据源表列号 = CheckedListBox2.SelectedIndex + 1
        If 填充表列号 = 0 Or 数据源表列号 = 0 Then
            MsgBox("请选择要匹配的列标题！")
        Else
            CheckedListBox3.Items.Add(填充表列号 & ":" & CheckedListBox1.SelectedItem & "---" & 数据源表列号 & ":" & CheckedListBox2.SelectedItem)
            Dim xy As New 数对(填充表列号, 数据源表列号)
            匹配列号序列.Add(xy)
            Label2.Text = 匹配号记录()

            加载列标题(填充表, NumericUpDown1.Value, CheckedListBox1)
            加载列标题(数据源表, NumericUpDown2.Value, CheckedListBox2)
        End If





    End Sub

    Public Function 匹配号记录() As String
        Dim str As String = ""
        For Each xy As 数对 In 匹配列号序列
            str &= xy.ToString & ";"
        Next
        Return str
    End Function

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        填充表 = app.Sheets(ComboBox1.Text)
        清除匹配列号序列()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        数据源表 = app.Sheets(ComboBox2.Text)
        清除匹配列号序列()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        加载列标题(填充表, NumericUpDown1.Value, CheckedListBox1)

    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged

        加载列标题(数据源表, NumericUpDown2.Value, CheckedListBox2)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        清除匹配列号序列()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        结果列号序列.Add(CheckedListBox2.SelectedIndex + 1)
        CheckedListBox4.Items.Add(CheckedListBox2.SelectedIndex + 1 & ":" & CheckedListBox2.SelectedItem)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        结果列号序列.Clear()
        CheckedListBox4.Items.Clear()

    End Sub



    Public Sub 清除匹配列号序列()
        匹配列号序列.Clear()
        CheckedListBox3.Items.Clear()
        Label2.Text = "按行匹配时，你需要选择匹配列号，请选择。"
        加载列标题(填充表, NumericUpDown1.Value, CheckedListBox1)
        加载列标题(数据源表, NumericUpDown2.Value, CheckedListBox2)
    End Sub

End Class
