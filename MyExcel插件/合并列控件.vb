Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class 合并列控件
    Public 无效值序列 As New List(Of String)
    Private Sub 合并列控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"合并列1", "合并列2", "合并列3", "合并列4"})
        区域选择控件1.是否允许编辑区域名 = False
        区域选择控件1.设置锁定用户区(True)
        区域选择控件1.是否显示添加按钮 = True

        无效值序列.Add(Nothing)



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If 区域选择控件1.区域校验(0, True, False, True) And
            区域选择控件1.区域校验(1, True, False, True) And
            区域选择控件1.区域校验(2, False, False, True) And
            区域选择控件1.区域校验(3, False, False, True) Then


            设置无效字符序列()

            If 冗余行列检查(区域选择控件1.获取区域(0).Worksheet, True) = True Then
                Exit Sub
            End If


            合并列(区域选择控件1.获取区域(0),
                  区域选择控件1.获取区域(1),
                  区域选择控件1.获取区域(2),
                  区域选择控件1.获取区域(3))

            流水信息.显示信息("警告")
        Else

            区域选择控件1.显示错误()
        End If

    End Sub


    Public Sub 合并列(合并列1 As Excel.Range,
                      合并列2 As Excel.Range,
                     Optional 合并列3 As Excel.Range = Nothing,
                     Optional 合并列4 As Excel.Range = Nothing)
        Dim 待合并的列序列 As New List(Of Excel.Range)
        Dim sheet As Excel.Worksheet = 合并列1.Worksheet
        Dim maxRow As Integer = 获取结束行号(sheet.UsedRange)
        合并列1 = sheet.Cells(NumericUpDown1.Value + 1, 合并列1.Column).Resize(maxRow - NumericUpDown1.Value, 1)

        合并列2 = sheet.Cells(NumericUpDown1.Value + 1, 合并列2.Column).Resize(maxRow - NumericUpDown1.Value, 1)

        待合并的列序列.Add(合并列1)
        待合并的列序列.Add(合并列2)
        If 合并列3 IsNot Nothing Then
            合并列3 = sheet.Cells(NumericUpDown1.Value + 1, 合并列3.Column).Resize(maxRow - NumericUpDown1.Value, 1)
            待合并的列序列.Add(合并列3)
        End If
        If 合并列4 IsNot Nothing Then
            合并列4 = sheet.Cells(NumericUpDown1.Value + 1, 合并列4.Column).Resize(maxRow - NumericUpDown1.Value, 1)
            待合并的列序列.Add(合并列4)
        End If

        Dim 数据行数 As Integer = 合并列1.Rows.Count
        Dim 数据序列 As New List(Of Excel.Range)
        Dim 新列 As Excel.Range = 插入列(sheet, 待合并的列序列(待合并的列序列.Count - 1).Column + 1)
        新列.Cells(NumericUpDown1.Value) = "#合并列"
        新列 = 新列.Cells(NumericUpDown1.Value + 1).Resize(maxRow - NumericUpDown1.Value, 1)

        If RadioButton1.Checked = True Then
            设置单元格格式(新列, "文本")
            Dim cell As Excel.Range
            For i = 1 To maxRow - NumericUpDown1.Value
                For Each 列 As Excel.Range In 待合并的列序列
                    数据序列.Add(列.Cells(i))
                Next
                cell = 获取合并值的单元格(数据序列)
                If cell IsNot Nothing Then
                    新列.Cells(i) = cell.Value
                End If
                数据序列.Clear()
            Next
        ElseIf RadioButton2.Checked = True Then
            Dim cell As Excel.Range
            For i = 1 To maxRow - NumericUpDown1.Value
                For Each 列 As Excel.Range In 待合并的列序列
                    数据序列.Add(列.Cells(i))
                Next
                cell = 获取合并值的单元格(数据序列)
                If cell IsNot Nothing Then


                    ' 复制源单元格的格式到目标单元格
                    新列.Cells(i).NumberFormatLocal = cell.NumberFormatLocal

                    新列.Cells(i).Font.Name = cell.Font.Name
                    '新列.Cells(i).Font.Size = cell.Font.Size
                    '新列.Cells(i).Font.Bold = cell.Font.Bold
                    '新列.Cells(i).Font.Italic = cell.Font.Italic
                    新列.Cells(i).Font.Color = cell.Font.Color
                    '新列.Cells(i).Font.SetSourceArray(cell.Font.GetFormatArray())


                    新列.Cells(i).Interior.Color = cell.Interior.Color

                    新列.Cells(i) = cell.Formula




                End If
                数据序列.Clear()
            Next
        End If







    End Sub

    Public Function 获取合并值的单元格(valueList As List(Of Excel.Range)) As Excel.Range
        Dim 有效值个数 As Integer = 0
        Dim 值单元格 As Excel.Range = Nothing
        For Each cell As Object In valueList
            If 是否为有效值(cell) = True Then
                If 有效值个数 = 0 Then
                    值单元格 = cell
                End If
                有效值个数 += 1

            End If
        Next
        If 有效值个数 > 1 Then
            Dim str As String = ""
            For Each cell As Excel.Range In valueList
                If cell.Value <> Nothing Then
                    str &= cell.Value.ToString & vbTab
                End If

            Next
            流水信息.记录信息("[警告！多个有效值，已引用首个]   工作表：" & valueList(0).Worksheet.Name & "  第 " & valueList(0）.Row & "行      数据：" & str)

        ElseIf 有效值个数 = 0 Then
            Dim str As String = ""
            For Each cell As Excel.Range In valueList
                If cell.Value <> Nothing Then
                    str &= cell.Value.ToString & vbTab
                End If
            Next
            流水信息.记录信息("[警告！无有效值]   工作表：" & valueList(0).Worksheet.Name & "  第 " & valueList(0）.Row & "行      数据：" & str)
        End If
        If 有效值个数 > 0 Then
            Return 值单元格
        Else
            Return Nothing
        End If
    End Function

    Public Function 是否为有效值(cell As Excel.Range) As Boolean
        Dim result As Boolean = True
        Dim cellValue As String
        For Each value As String In 无效值序列



            If cell.Value <> Nothing Then
                cellValue = cell.Value.ToString
            Else
                cellValue = ""
            End If

            If cellValue = value Then
                result = False
            End If
        Next
        Return result
    End Function



    Public Sub 设置无效字符序列()
        无效值序列.Clear()
        Dim n As Integer = MyDataGridView1.Rows.Count - 1
        If CheckBox1.Checked = True Then
            无效值序列.Add(Nothing)
        End If
        For i = 0 To n - 1
            For Each cell In MyDataGridView1.Rows.Item(i).Cells
                无效值序列.Add(cell.Value)
            Next
        Next
    End Sub


    Private Sub CheckBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles CheckBox1.MouseClick

    End Sub

    'Private Sub 粘贴ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 粘贴ToolStripMenuItem.Click
    '    Dim 行序列 As String = Clipboard.GetText()
    '    If 行序列.EndsWith(vbCrLf) Then
    '        行序列 = 行序列.Substring(0, 行序列.Length - 2)
    '    End If
    '    Dim 列数 As Integer = 行序列.Split(vbCr & vbLf)(0).Split(vbTab).Length
    '    Dim 行数 As Integer = 行序列.Split(vbCr & vbLf).Length
    '    Dim CurrentRow, CurrentColumn, 总行数, 总列数 As Integer
    '    CurrentRow = DataGridView1.CurrentCell.RowIndex
    '    CurrentColumn = DataGridView1.CurrentCell.ColumnIndex
    '    总行数 = DataGridView1.RowCount
    '    总列数 = DataGridView1.ColumnCount
    '    If 列数 > 总列数 - CurrentColumn Then
    '        MsgBox("控件不允许粘贴！")
    '        Exit Sub
    '    End If
    '    'For i As Integer = 1 To 列数
    '    '    DataGridView1.Columns.Add("数据" & i， "数据" & i)
    '    'Next


    '    For i As Integer = 1 To 行数 - (总行数 - CurrentRow)
    '        DataGridView1.Rows.Add()
    '    Next



    '    Dim j As Integer = CurrentColumn

    '    Dim n As Integer = CurrentRow
    '    For Each 行 As String In 行序列.Split(vbCr & vbLf)
    '        If 行.StartsWith(vbLf) Then
    '            行 = 行.Substring(1, 行.Length - 1)
    '        End If
    '        For Each 单元格 As String In 行.Split(vbTab)
    '            DataGridView1.Rows(n).Cells(j).Value = 单元格
    '            j += 1
    '        Next
    '        n += 1
    '        j = CurrentColumn
    '    Next
    'End Sub

    Private Sub 区域选择控件1_Load(sender As Object, e As EventArgs) Handles 区域选择控件1.Load

        MyDataGridView1.Columns.Clear()
        MyDataGridView1.addColumn("无效值", "无效值")

    End Sub
End Class
