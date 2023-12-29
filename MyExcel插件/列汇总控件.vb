'Imports Microsoft.Office.Interop.Excel

Imports System.Diagnostics.Eventing.Reader
Imports System.Windows.Forms
Imports HetingControl

Public Class 列汇总控件


    Public 合并表序列 As New Collections.ArrayList '其中存放 表中列序列
    Public 表中列序列 As New List(Of String) '存放 要合并的每个表的各个列信息，0号元素为表名，之后为要合并的列的序号格式"操作序列号,合并目标序列号",如"1,3"
    Public 当前填充表索引, 当前操作表索引 As Integer

    Private Sub 列填充_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表(CheckedListBox1)
        加载表(ComboBox2)
        当前填充表索引 = -1
        当前操作表索引 = -1
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
        If RadioButton5.Checked = True Then
            合并自动设置的表格列()

        Else
            合并手动设置的表格列()

        End If




    End Sub


    Public Sub 合并自动设置的表格列()
        If CheckedListBox1.CheckedItems.Count = 0 Then
            MsgBox("请选择 数据源表 和 填充表，不然怎么干！")
            Exit Sub
        End If






        Dim 拷贝方式 As Integer
        Dim 填充表区域当前最大行, 数据源区域最大行 As Integer
        Dim ColumenNum As Integer = 0
        Dim 数据源表, 填充表 As Excel.Worksheet
        Dim 数据源区域, 填充表区域, 填充表列头, 操作区域 As Excel.Range





        For Each temp In CheckedListBox1.CheckedItems
            数据源表 = app.Sheets(temp.ToString)
            If 冗余行列检查(数据源表, True) = True Then
                Exit Sub
            End If
        Next

        If CheckBox1.Checked = True Then
            填充表 = 创建新的列汇总表()
        Else

            If ComboBox2.Text = "" Then
                MsgBox("请选择 数据源表 和 填充表，不然没法干！")
                Exit Sub
            End If
            If 是否存在工作表(ComboBox2.Text) = False Then
                MsgBox(ComboBox2.Text & " 这表就不存在！")
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









    Public Sub 合并手动设置的表格列()
        Dim 已汇总的表个数 As Integer = 0
        Dim 填充表 As Excel.Worksheet
        Dim 数据源表 As Excel.Worksheet
        Dim 数据源区域, 填充表区域, 填充表列头, 操作区域 As Excel.Range
        Dim 拷贝方式 As Integer
        Dim 填充表区域当前最大行, 数据源区域最大行 As Integer
        Dim ColumenNum As Integer = 0




        For Each 表 As List(Of String) In 合并表序列
            If 已汇总的表个数 = 0 Then
                填充表 = app.Worksheets(表(0))
                填充表区域 = 获取工作区域(填充表)
                填充表列头 = 填充表区域.Rows(NumericUpDown1.Value)
                填充表区域当前最大行 = 填充表区域.Rows.Count
            Else
                填充表区域 = 获取工作区域(填充表)
                填充表列头 = 填充表区域.Rows(NumericUpDown1.Value)
                填充表区域当前最大行 = 填充表区域.Rows.Count

                数据源表 = app.Worksheets(表(0))
                数据源区域 = 获取工作区域(数据源表)
                数据源区域最大行 = 数据源区域.Rows.Count

                If RadioButton1.Checked = True Then
                    拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll
                ElseIf RadioButton2.Checked = True Then
                    拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues
                ElseIf RadioButton3.Checked = True Then
                    拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas

                ElseIf RadioButton4.Checked = True Then
                    拷贝方式 = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats
                End If


                If 数据源区域最大行 = 1 Then '只有表头没有内容
                    Continue For
                End If

                Dim 数据源列, 填充表列 As Integer
                Dim 表头单元格 As Excel.Range
                For Each record As String In 表
                    数据源列 = 获取操作列号(record)
                    填充表列 = 获取合并目标列号(record)
                    If 数据源列 < 0 Or 填充表列 < 0 Then
                        Continue For
                    Else
                        表头单元格 = MyRange(填充表, NumericUpDown1.Value, 填充表列)
                        If 表头单元格.Value = Nothing Then
                            表头单元格.Value = MyRange(数据源表, NumericUpDown1.Value, 数据源列)
                        End If
                        操作区域 = MyRange(数据源表, NumericUpDown1.Value + 1, 数据源列, 数据源区域最大行, 数据源列)
                        操作区域.Copy()
                        MyRange(填充表, 填充表区域当前最大行 + 1, 填充表列).PasteSpecial(拷贝方式)
                        ColumenNum += 1
                    End If

                Next


            End If
            已汇总的表个数 += 1
        Next
        If CheckBox1.Checked = True Then
            已汇总的表个数 -= 1
        End If
        自动列宽(填充表)

        MsgBox("数据被汇总到:  " & 填充表.Name & vbCrLf & "共汇总 " & 已汇总的表个数 & " 个表" & vbCrLf & "共汇总了 " & ColumenNum & " 列")
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
        If CheckedListBox1.SelectedItem <> Nothing Then
            加载列标题到数据表(app.Worksheets(CheckedListBox1.SelectedItem.ToString), NumericUpDown1.Value, DataGridView2)
            DataGridView2.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable

            当前操作表索引 = CheckedListBox1.SelectedIndex
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        ComboBox2.Enabled = Not CheckBox1.Checked
        全部清理()

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckedListBox1.MultiColumn = CheckBox2.Checked
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedItem <> Nothing Then
            加载列标题到数据表(app.Worksheets(ComboBox2.SelectedItem.ToString), NumericUpDown1.Value, DataGridView1, True)
            DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            Dim 操作序列 As New List(Of String)
            操作序列.Add(ComboBox2.Text)

            Dim sheet As Excel.Worksheet = app.Worksheets(ComboBox2.SelectedItem.ToString)
            Dim 表头区域 As Excel.Range = 区域交集(sheet.Rows(NumericUpDown1.Value), sheet.UsedRange)
            For Each cell As Excel.Range In 表头区域
                If cell.Value = Nothing Then
                    操作序列.Add(Nothing)
                Else
                    操作序列.Add(cell.Value.ToString)
                End If

            Next

            合并表序列.Clear()
            合并表序列.Add(操作序列)
            当前填充表索引 = ComboBox2.SelectedIndex

        End If
    End Sub

    Private Sub DataGridView2_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseDoubleClick
        If 当前操作表索引 < 0 Then
            Exit Sub
        End If
        If 合并表序列.Count = 0 Then
            If CheckBox1.Checked = True Then '创建新表
                Dim tempSheet As Excel.Worksheet = app.Worksheets(当前操作表索引 + 1)

                'Dim lastSheetIndex As Integer = app.Worksheets.Count - 1
                ' 获取当前工作簿中最后一个工作表的索引

                Dim newsheet As Excel.Worksheet = 新建工作表("列汇总表", True, False, After:=app.Worksheets(app.Worksheets.Count))
                加载列标题到数据表(newsheet, NumericUpDown1.Value, DataGridView1, False)
                DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                Dim 新序列 As New List(Of String)
                新序列.Add(newsheet.Name)
                合并表序列.Add(新序列)

                加载表(CheckedListBox1)
                加载表(ComboBox2)
                当前操作表索引 = tempSheet.Index - 1
            Else
                MsgBox("请选择要合并到的结果表格。")
                Exit Sub
            End If
        End If


        If DataGridView1.CurrentCell Is Nothing Then
            MsgBox("请在左边选择要与此列合并的列，或者是选择新列。")
        End If

        If DataGridView1.CurrentCell.RowIndex = DataGridView1.Rows.Count - 1 Then
            DataGridView1.Rows.Add()
            DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex - 1).Cells(DataGridView1.CurrentCell.ColumnIndex)

        End If

        Dim 操作表 As Excel.Worksheet = app.Worksheets(当前操作表索引 + 1)
        Dim 操作序列 As List(Of String)
        'app.Worksheets(CheckedListBox1.SelectedItem.ToString)
        For Each lis In 合并表序列
            If lis(0) = 操作表.Name Then
                操作序列 = lis
            End If
        Next
        If 操作序列 Is Nothing Then
            操作序列 = New List(Of String)
            操作序列.Add(操作表.Name)
            合并表序列.Add(操作序列)
            Dim index As Integer = DataGridView1.Columns.Add(操作表.Name, 操作表.Name)

            CheckedListBox1.SetItemChecked(当前操作表索引, True)
            Label3.Text = "当前数据表（共" & CheckedListBox1.Items.Count & "个,选中" & CheckedListBox1.CheckedItems.Count & "个）"
            DataGridView1.Columns(index).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView1.Columns(index).ReadOnly = True
        End If

        Dim 操作序列的操作元素序号 As Integer = Nothing
        For n = 1 To 操作序列.Count - 1
            If 获取合并目标列号(操作序列(n)) = DataGridView1.CurrentCell.RowIndex + 1 Then
                操作序列的操作元素序号 = n
                'MsgBox(DataGridView2.CurrentCell.RowIndex + 1 & "," & DataGridView1.CurrentCell.RowIndex + 1)
                操作序列(操作序列的操作元素序号) = DataGridView2.CurrentCell.RowIndex + 1 & "," & DataGridView1.CurrentCell.RowIndex + 1
            End If
        Next

        If 操作序列的操作元素序号 = Nothing Then
            操作序列.Add(DataGridView2.CurrentCell.RowIndex + 1 & "," & DataGridView1.CurrentCell.RowIndex + 1)
        End If


        Dim columnIndex As DataGridViewColumn = DataGridView1.Columns.Item(操作表.Name)




        DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(columnIndex.Index).Value = DataGridView2.CurrentCell.Value

        If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = Nothing Then
            DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = DataGridView2.CurrentCell.Value
            If DataGridView1.CurrentCell.RowIndex + 1 > 合并表序列(0).Count - 1 Then
                合并表序列(0).add(DataGridView2.CurrentCell.Value)
            Else

                合并表序列(0)(DataGridView1.CurrentCell.RowIndex + 1) = DataGridView2.CurrentCell.Value
            End If

        End If

        'If DataGridView1.CurrentCell.RowIndex = DataGridView1.Rows.Count - 2 Then
        '    DataGridView1.Rows.Add()
        'End If

        If DataGridView1.CurrentCell.RowIndex + 1 <= DataGridView1.Rows.Count - 1 Then
            DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex + 1).Cells(DataGridView1.CurrentCell.ColumnIndex)
        End If





        'MsgBox(DataGridView1.CurrentCell.RowIndex & "  " & DataGridView1.Rows.Count)
    End Sub





    Public Function 获取操作列号(ByVal inputString As String) As Integer
        ' 使用String.Split方法分离字符串
        Dim parts As String() = inputString.Split(","c)

        ' 检查是否至少有一个部分
        If parts.Length > 0 Then
            ' 尝试将第一个部分转换为整数并返回
            If Integer.TryParse(parts(0), 获取操作列号) Then
                Return 获取操作列号
            Else
                Return -1
            End If
        Else
            Return -1
        End If
    End Function

    Public Function 获取合并目标列号(ByVal inputString As String) As Integer
        ' 使用String.Split方法分离字符串
        Dim parts As String() = inputString.Split(","c)

        ' 检查是否有第二个部分
        If parts.Length > 1 Then
            ' 尝试将第二个部分转换为整数并返回
            If Integer.TryParse(parts(1), 获取合并目标列号) Then
                Return 获取合并目标列号
            Else
                Return -1
            End If
        Else
            Return -1
        End If
    End Function

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub

    Private Sub 全部清除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 全部清除ToolStripMenuItem.Click
        合并表序列.Clear()
        DataGridView1.Columns.Clear()
        DataGridView2.Columns.Clear()
        ComboBox2.SelectedItem = Nothing

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then

            DataGridView1.Visible = False
            DataGridView2.Visible = False
            CheckedListBox1.CheckOnClick = True
        Else

            DataGridView1.Visible = True
            DataGridView2.Visible = True
            CheckedListBox1.CheckOnClick = False
        End If
        全部清理()
    End Sub

    Public Sub 全部清理()
        合并表序列.Clear()
        DataGridView1.Columns.Clear()
        DataGridView2.Columns.Clear()
        ComboBox2.SelectedItem = Nothing
        当前填充表索引 = -1
        当前操作表索引 = -1
        加载表(ComboBox2)
        加载表(CheckedListBox1)
    End Sub


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    '''


End Class
