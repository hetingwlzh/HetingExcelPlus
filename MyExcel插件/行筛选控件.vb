Imports System.CodeDom.Compiler
Imports System.Drawing
Imports System.Windows.Forms
Imports MyExcel插件.行筛选控件

Public Class 行筛选控件

    Public 筛选器 As New 多列条件类
    Public 待筛选表 As Excel.Worksheet
    Public 筛选出的区域 As Excel.Range
    Public 是否忽略空值 As Boolean = True
    Public Enum 筛选结果类型

        通过 = 1
        未通过 = 0
        出错 = 3
    End Enum
    Public Enum 比对类型
        自动 = 0
        数值 = 1
        字符串 = 2
    End Enum
    Public Enum 关系运算

        等于 = 0
        不等 = 1
        大于 = 2
        大于等于 = 3
        小于 = 4
        小于等于 = 5
        包含于 = 6
        包含 = 7
    End Enum
    Public Enum 逻辑运算
        且 = 0
        或 = 1

    End Enum
    Public Class 条件类
        Public 关系 As 关系运算
        Public 值类型 As 比对类型
        Public 数值 As Double
        Public 字符串值 As String
        Sub New(关系_ As 关系运算, 值_ As String, 值类型 As 比对类型)
            关系 = 关系_
            If IsNumeric(值_) Then
                数值 = CType(值_, Double)
            End If
            字符串值 = 值_
        End Sub
        Public Function 条件检测(testCell As Excel.Range, Optional 是否忽略空值 As Boolean = True) As 筛选结果类型
            Dim isNumber1 As Boolean = IsNumeric(testCell.Value)
            Dim isNumber2 As Boolean = IsNumeric(字符串值)
            If 关系 = 关系运算.等于 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If 值类型 = 比对类型.自动 Then
                        If isNumber1 <> isNumber2 Then
                            Return 筛选结果类型.未通过
                        Else
                            If testCell.Value = 字符串值 Then
                                Return 筛选结果类型.通过
                            Else
                                Return 筛选结果类型.未通过
                            End If
                        End If
                    ElseIf 值类型 = 比对类型.数值 Then
                        If testCell.Value = 数值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    ElseIf 值类型 = 比对类型.字符串 Then
                        If testCell.Value.ToString = 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    Else
                        Return 筛选结果类型.出错
                    End If
                End If



            ElseIf 关系 = 关系运算.不等 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If 值类型 = 比对类型.自动 Then
                        If isNumber1 <> isNumber2 Then
                            Return 筛选结果类型.通过
                        Else
                            If testCell.Value <> 字符串值 Then
                                Return 筛选结果类型.通过
                            Else
                                Return 筛选结果类型.未通过
                            End If
                        End If



                    ElseIf 值类型 = 比对类型.数值 Then


                        If testCell.Value <> 数值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If





                    ElseIf 值类型 = 比对类型.字符串 Then

                        If testCell.Value.ToString <> 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If



                    Else
                        Return 筛选结果类型.出错
                    End If
                End If


            ElseIf 关系 = 关系运算.小于 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If isNumber1 = True And isNumber2 = True Then
                        If testCell.Value < 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If


                    Else
                        'MsgBox("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        Return 筛选结果类型.出错
                    End If
                End If





            ElseIf 关系 = 关系运算.小于等于 Then

                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If isNumber1 = True And isNumber2 = True Then


                        If testCell.Value <= 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    Else
                        'MsgBox("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        Return 筛选结果类型.出错
                    End If
                End If



            ElseIf 关系 = 关系运算.大于 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If isNumber1 = True And isNumber2 = True Then
                        If testCell.Value > 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    Else
                        'MsgBox("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")

                        Return 筛选结果类型.出错
                    End If
                End If



            ElseIf 关系 = 关系运算.大于等于 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If isNumber1 = True And isNumber2 = True Then
                        If testCell.Value >= 字符串值 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    Else
                        'MsgBox("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & """" & testCell.Value & """" & " 是文本类型不能比较大小")
                        Return 筛选结果类型.出错
                    End If
                End If


            ElseIf 关系 = 关系运算.包含 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If testCell.Value Is Nothing Then
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & " 为空")
                        Return 筛选结果类型.未通过
                    Else
                        If testCell.Value.ToString.IndexOf(字符串值) >= 0 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If

                    End If
                End If



            ElseIf 关系 = 关系运算.包含于 Then
                If testCell.Value Is Nothing And 是否忽略空值 = True Then
                    Return 筛选结果类型.未通过
                Else
                    If testCell.Value Is Nothing Then
                        流水信息.记录信息("单元格(" & testCell.Row & "," & testCell.Column & ")" & " 为空")
                        Return 筛选结果类型.未通过
                    Else

                        If 字符串值.IndexOf(testCell.Value.ToString) >= 0 Then
                            Return 筛选结果类型.通过
                        Else
                            Return 筛选结果类型.未通过
                        End If
                    End If

                End If




            Else
                流水信息.记录信息("未预料错误")

                Return 筛选结果类型.出错
            End If

        End Function
        Public Overrides Function ToString() As String

            Dim 条件关系 As String
            If 关系 = 关系运算.等于 Then
                条件关系 = "="
            ElseIf 关系 = 关系运算.不等 Then
                条件关系 = "<>"
            ElseIf 关系 = 关系运算.小于 Then
                条件关系 = "<"
            ElseIf 关系 = 关系运算.大于 Then
                条件关系 = ">"
            ElseIf 关系 = 关系运算.小于等于 Then
                条件关系 = "<="
            ElseIf 关系 = 关系运算.大于等于 Then
                条件关系 = ">="
            ElseIf 关系 = 关系运算.包含于 Then
                条件关系 = "⊆"
            ElseIf 关系 = 关系运算.包含 Then
                条件关系 = "⊇"
            Else
                条件关系 = ""
            End If

            If IsNumeric(字符串值) Then
                Return 条件关系 & 字符串值
            Else
                Return 条件关系 & """" & 字符串值 & """"

            End If

        End Function
    End Class
    Public Class 单列条件类
        Public 条件序列 As List(Of 条件类)
        Public 逻辑 As 逻辑运算
        Public 列号 As Integer
        Sub New()
            条件序列 = New List(Of 条件类)
            逻辑 = 逻辑运算.且
            列号 = 1
        End Sub
        Sub New(待检测列号 As Integer, 条件1 As 条件类, Optional 检测逻辑 As 逻辑运算 = 逻辑运算.且)
            列号 = 待检测列号
            条件序列 = New List(Of 条件类)
            条件序列.Add(条件1)
            逻辑 = 检测逻辑
        End Sub
        Sub New(待检测列号 As Integer, 条件1 As 条件类, 条件2 As 条件类, Optional 检测逻辑 As 逻辑运算 = 逻辑运算.且)
            列号 = 待检测列号
            条件序列 = New List(Of 条件类)
            条件序列.Add(条件1)
            条件序列.Add(条件2)
            逻辑 = 检测逻辑
        End Sub
        Sub New(待检测列号 As Integer, 条件1 As 条件类, 条件2 As 条件类, 条件3 As 条件类, Optional 检测逻辑 As 逻辑运算 = 逻辑运算.且)
            列号 = 待检测列号
            条件序列 = New List(Of 条件类)
            条件序列.Add(条件1)
            条件序列.Add(条件2)
            条件序列.Add(条件3)
            逻辑 = 检测逻辑
        End Sub
        Sub New(待检测列号 As Integer, 条件1 As 条件类, 条件2 As 条件类, 条件3 As 条件类, 条件4 As 条件类, Optional 检测逻辑 As 逻辑运算 = 逻辑运算.且)
            列号 = 待检测列号
            条件序列 = New List(Of 条件类)
            条件序列.Add(条件1)
            条件序列.Add(条件2)
            条件序列.Add(条件3)
            条件序列.Add(条件4)
            逻辑 = 逻辑运算.且
        End Sub
        Public Sub Add(新条件 As 条件类)
            条件序列.Add(新条件)
        End Sub

        Public Function 条件检测(testCell As Excel.Range, Optional 是否忽略空值 As Boolean = True) As 筛选结果类型
            Dim Temp As 筛选结果类型
            If 逻辑 = 逻辑运算.且 Then

                For Each 条件 As 条件类 In 条件序列
                    Temp = 条件.条件检测(testCell, 是否忽略空值)
                    If Temp = 筛选结果类型.未通过 Then
                        Return 筛选结果类型.未通过
                    ElseIf Temp = 筛选结果类型.出错 Then
                        Return 筛选结果类型.出错
                    End If
                Next
                Return 筛选结果类型.通过
            ElseIf 逻辑 = 逻辑运算.或 Then
                For Each 条件 As 条件类 In 条件序列
                    Temp = 条件.条件检测(testCell, 是否忽略空值)
                    If Temp = 筛选结果类型.通过 Then
                        Return 筛选结果类型.通过
                    ElseIf Temp = 筛选结果类型.出错 Then
                        Return 筛选结果类型.出错

                    End If
                Next
                Return 筛选结果类型.未通过
            Else
                流水信息.记录信息("未预料错误")
                Return 筛选结果类型.出错
            End If
        End Function


    End Class

    Public Class 多列条件类
        Public 单列条件序列 As New List(Of 单列条件类）
        Public 通过行数 As Long = 0
        Public 未通过行数 As Long = 0
        Public 出错行数 As Long = 0
        Public Sub 计数清零()
            通过行数 = 0
            未通过行数 = 0
            出错行数 = 0
        End Sub
        Sub New()
            单列条件序列.Clear()
            通过行数 = 0
            未通过行数 = 0
            出错行数 = 0
        End Sub
        Sub New(单列条件1 As 单列条件类)
            单列条件序列.Add(单列条件1)
        End Sub
        Sub New(单列条件1 As 单列条件类, 单列条件2 As 单列条件类)
            单列条件序列.Add(单列条件1)
            单列条件序列.Add(单列条件2)
        End Sub
        Sub New(单列条件1 As 单列条件类, 单列条件2 As 单列条件类, 单列条件3 As 单列条件类)
            单列条件序列.Add(单列条件1)
            单列条件序列.Add(单列条件2)
            单列条件序列.Add(单列条件3)
        End Sub
        Sub New(单列条件1 As 单列条件类, 单列条件2 As 单列条件类, 单列条件3 As 单列条件类, 单列条件4 As 单列条件类)
            单列条件序列.Add(单列条件1)
            单列条件序列.Add(单列条件2)
            单列条件序列.Add(单列条件3)
            单列条件序列.Add(单列条件4)
        End Sub

        Public Function Add(单列条件 As 单列条件类) As Integer
            单列条件序列.Add(单列条件)
            Return 单列条件序列.Count
        End Function
        Public Function 条件检测(TestRow As Excel.Range, Optional 是否忽略空值 As Boolean = True) As 筛选结果类型
            Try
                If TestRow.Rows.Count = 1 Then
                    Dim Temp As 筛选结果类型
                    For Each 单列条件 As 单列条件类 In 单列条件序列
                        Temp = 单列条件.条件检测(TestRow.Cells(1, 单列条件.列号), 是否忽略空值)
                        If Temp = 筛选结果类型.未通过 Then

                            未通过行数 += 1
                            Return 筛选结果类型.未通过

                        ElseIf Temp = 筛选结果类型.出错 Then
                            出错行数 += 1
                            Return 筛选结果类型.出错
                        End If
                    Next
                    通过行数 += 1
                    Return 筛选结果类型.通过
                Else
                    MsgBox("待测区域不是单行，请检查！")
                    Return False
                End If
            Catch ex As Exception
                MsgBox("多列条件检测时出现未知错误，请检查！")
                Return False
            End Try
        End Function

        Public Sub 删除条件(所在单列条件列号 As Integer, 条件序号 As Integer)
            Try
                Dim 待删除单列条件 As 单列条件类 = Nothing
                If 条件序号 > 0 Then
                    For Each 单列条件 As 单列条件类 In 单列条件序列
                        If 单列条件.列号 = 所在单列条件列号 Then
                            单列条件.条件序列.RemoveAt(条件序号 - 1)
                            If 单列条件.条件序列.Count = 0 Then
                                待删除单列条件 = 单列条件
                            End If
                        End If
                    Next
                    If 待删除单列条件 IsNot Nothing Then
                        单列条件序列.Remove(待删除单列条件)
                    End If
                End If
            Catch ex As Exception
                MsgBox("删除时出错！")
            End Try


        End Sub



    End Class
    Public Sub 加载列标题(sheet As Excel.Worksheet, 标题行行号 As Integer, MyDataGridView As Windows.Forms.DataGridView)


        If sheet IsNot Nothing Then

            MyDataGridView.Columns.Clear()
            'CheckedListBox1.Items.Clear()
            MyDataGridView.Columns.Add("列标题", "列标题")
            MyDataGridView.Columns(0).ReadOnly = True
            MyDataGridView.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            MyDataGridView.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            MyDataGridView.Columns(0).HeaderCell.Style.BackColor = Color.BlueViolet
            MyDataGridView.Columns(0).HeaderCell.Style.ForeColor = Color.White
            MyDataGridView.Columns(0).HeaderCell.Style.Font = New Font("Microsoft YaHei", 12, FontStyle.Bold)




            'MyDataGridView.Columns.Add("关系", "关系")
            'MyDataGridView.Columns(0).ReadOnly = True
            'MyDataGridView.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            'MyDataGridView.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            'MyDataGridView.Columns(0).HeaderCell.Style.BackColor = Color.BlueViolet
            'MyDataGridView.Columns(0).HeaderCell.Style.ForeColor = Color.White
            'MyDataGridView.Columns(0).HeaderCell.Style.Font = New Font("Microsoft YaHei", 12, FontStyle.Bold)



            'MyDataGridView.Columns.Add("值", "值")
            'MyDataGridView.Columns(0).ReadOnly = True
            'MyDataGridView.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            'MyDataGridView.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            'MyDataGridView.Columns(0).HeaderCell.Style.BackColor = Color.BlueViolet
            'MyDataGridView.Columns(0).HeaderCell.Style.ForeColor = Color.White
            'MyDataGridView.Columns(0).HeaderCell.Style.Font = New Font("Microsoft YaHei", 12, FontStyle.Bold)






            Dim str As String = ""

            If 标题行行号 > sheet.UsedRange.Rows.Count Then
                MsgBox("标题行行号超出使用区范围，请重新设置标题行所在区域的行号。")
                Exit Sub
            End If
            Dim n As Integer = 0
            For Each cell As Excel.Range In sheet.UsedRange.Rows.Item(标题行行号).Cells
                n = MyDataGridView.Rows.Add()
                MyDataGridView.Rows(n).Cells(0).Value = ToStr(cell)
                'CheckedListBox1.Items.Add(ToStr(cell))
                'MyDataGridView.Rows(n).Cells(1).Value = ""
            Next
        End If

    End Sub




    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        待筛选表 = app.Sheets(ComboBox2.Text)
        加载列标题(待筛选表, NumericUpDown1.Value, DataGridView1)

    End Sub

    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        加载表(ComboBox2)
    End Sub

    Private Sub 行筛选控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表(ComboBox2)
        ComboBox2.SelectedIndex = 0


    End Sub

    Public Function 生成条件(关系 As 关系运算,
                         值 As String,
                         值类型 As 比对类型) As 条件类
        Dim result As New 条件类(关系, 值, 值类型)

    End Function

    Public Function 添加条件(列号 As Integer,
                         条件 As 条件类,
                         逻辑 As 逻辑运算) As 单列条件类
        Dim 是否存在此列 As Boolean = False
        For Each 单列条件 As 单列条件类 In 筛选器.单列条件序列
            If 单列条件.列号 = 列号 Then
                是否存在此列 = True
                If 逻辑 = 单列条件.逻辑 Then
                    单列条件.Add(条件)
                    Return 单列条件
                Else
                    MsgBox("与之前的条件逻辑不吻合！添加失败！")
                    Return Nothing
                End If

            End If

        Next
        If 是否存在此列 = False Then
            Dim 新单列条件 As New 单列条件类(列号, 条件, 逻辑)
            筛选器.Add(新单列条件)

        End If

        'Dim result As New 条件类(关系, 值, 值类型)

    End Function

    ''' <summary>
    ''' 显示多列条件对象中包含的所有筛选条件
    ''' </summary>
    ''' <param name="多列条件">要显示的多列条件对象</param>
    Public Sub 显示多列条件(多列条件 As 多列条件类)
        Dim 所选行索引, 所选列索引 As Integer
        所选行索引 = DataGridView1.CurrentCell.OwningRow.Index
        所选列索引 = DataGridView1.CurrentCell.OwningColumn.Index
        加载列标题(app.Sheets(ComboBox2.Text), NumericUpDown1.Value, DataGridView1)


        For Each 单列条件 As 单列条件类 In 多列条件.单列条件序列
            Dim 条件序号 As Integer = 1
            For Each 条件 As 条件类 In 单列条件.条件序列


                Dim 条件逻辑 As String
                If 单列条件.逻辑 = 逻辑运算.且 Then
                    条件逻辑 = "(AND)"
                Else
                    条件逻辑 = "(OR)"
                End If
                If DataGridView1.Columns.Count < 条件序号 + 1 Then
                    Dim index As Integer = DataGridView1.Columns.Add("条件" & 条件序号, "条件" & 条件序号)
                    DataGridView1.Columns(index).ReadOnly = True
                    DataGridView1.Columns(index).SortMode = DataGridViewColumnSortMode.NotSortable
                    DataGridView1.Columns(index).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                    DataGridView1.Columns(index).HeaderCell.Style.BackColor = Color.BlueViolet
                    DataGridView1.Columns(index).HeaderCell.Style.ForeColor = Color.White
                    DataGridView1.Columns(index).HeaderCell.Style.Font = New Font("Microsoft YaHei", 12, FontStyle.Bold)
                End If
                'MsgBox(DataGridView1.CurrentCell.OwningRow.Index & "   " & DataGridView1.CurrentCell.OwningColumn.Index)
                DataGridView1.Rows(单列条件.列号 - 1).Cells(条件序号).Value = 条件.ToString & 条件逻辑
                条件序号 += 1
            Next



        Next
        If DataGridView1.Columns.Count - 1 < 所选列索引 Then
            所选列索引 = DataGridView1.Columns.Count - 1
        End If
        DataGridView1.CurrentCell = DataGridView1.Rows(所选行索引).Cells(所选列索引)





    End Sub
    ''' <summary>
    ''' 根据用户在界面上设置的条件,创建筛选条件对象并添加到筛选器中
    ''' </summary>
    Public Sub 创建新条件()
        Dim 当前编辑列号 As Integer = DataGridView1.CurrentCell.OwningRow.Index + 1
        Dim 关系 As 关系运算
        Dim 值 As String
        Dim 值类型 As 比对类型
        Dim 逻辑 As 逻辑运算
        If RadioButton1.Checked = True Then
            关系 = 关系运算.等于
        ElseIf RadioButton2.Checked = True Then
            关系 = 关系运算.不等
        ElseIf RadioButton3.Checked = True Then
            关系 = 关系运算.小于
        ElseIf RadioButton4.Checked = True Then
            关系 = 关系运算.大于
        ElseIf RadioButton5.Checked = True Then
            关系 = 关系运算.小于等于
        ElseIf RadioButton6.Checked = True Then
            关系 = 关系运算.大于等于
        ElseIf RadioButton7.Checked = True Then
            关系 = 关系运算.包含于
        ElseIf RadioButton8.Checked = True Then
            关系 = 关系运算.包含
        End If



        If RadioButton9.Checked = True Then
            值类型 = 比对类型.自动
        ElseIf RadioButton10.Checked = True Then
            值类型 = 比对类型.数值
        ElseIf RadioButton11.Checked = True Then
            值类型 = 比对类型.字符串
        End If


        If RadioButton17.Checked = True Then
            逻辑 = 逻辑运算.且
        ElseIf RadioButton18.Checked = True Then
            逻辑 = 逻辑运算.或
        End If

        If RadioButton15.Checked Then
            值 = TextBox1.Text
        ElseIf RadioButton16.Checked = True Then
            值 = TextBox2.Text
        End If

        Dim 当前条件 As New 条件类(关系, 值, 值类型)

        添加条件(当前编辑列号, 当前条件, 逻辑)

        显示多列条件(筛选器)
    End Sub

    Public Sub 删除当前所选条件()
        Dim 所选行索引, 所选列索引 As Integer
        所选行索引 = DataGridView1.CurrentCell.OwningRow.Index
        所选列索引 = DataGridView1.CurrentCell.OwningColumn.Index
        筛选器.删除条件(所选行索引 + 1, 所选列索引)
        'If DataGridView1.Columns.Count - 1 < 所选列索引 Then
        '    所选列索引 = DataGridView1.Columns.Count - 1
        'End If
        'DataGridView1.CurrentCell = DataGridView1.Rows(所选行索引).Cells(所选列索引)
        显示多列条件(筛选器)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        创建新条件()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        删除当前所选条件()
    End Sub
    ''' <summary>
    ''' 根据指定的筛选条件,对Excel工作表进行数据筛选,将筛选结果写入新工作表中
    ''' </summary>
    ''' <param name="sheet">要筛选的工作表</param>
    ''' <param name="列表头尾行号">列表头占用的行数</param>
    ''' <param name="筛选条件">筛选条件规则</param>
    ''' <returns>筛选后的总行数</returns>

    Public Function 工作表筛选内容(sheet As Excel.Worksheet,
                            列表头尾行号 As Integer,
                            筛选条件 As 多列条件类,
                            Optional 是否忽略空值 As Boolean = True,
                            Optional 是否输出到新表 As Boolean = True) As Integer
        Dim 总筛选行数 As Integer = 0
        Dim 工作区 As Excel.Range = 获取工作区域(sheet)

        Dim 总行数 As Integer = 工作区.Rows.Count
        Dim 总列数 As Integer = 工作区.Columns.Count
        Dim 数据行数 As Integer = 总行数 - 列表头尾行号
        Dim 列表头区 As Excel.Range = 工作区(1, 1).Resize(列表头尾行号, 总列数)
        Dim 行数据区 As Excel.Range = 工作区(列表头尾行号 + 1, 1).Resize(数据行数, 总列数)

        Dim 筛选结果写入行号 As Integer = 1
        'app.ScreenUpdating = False
        '列表头区.Copy()
        If 是否输出到新表 = True Then
            筛选出的区域 = 列表头区
        End If

        'resultSheet.Cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
        筛选结果写入行号 = 列表头尾行号 + 1
        筛选器.计数清零()
        Dim temp As 筛选结果类型
        For Each row As Excel.Range In 行数据区.Rows
            temp = 筛选器.条件检测(row， 是否忽略空值)
            If temp = 筛选结果类型.通过 Then

                If 筛选出的区域 IsNot Nothing Then
                    筛选出的区域 = app.Union(筛选出的区域, row)
                Else
                    筛选出的区域 = row
                End If


                筛选结果写入行号 += 1
                'For Each cell As Excel.Range In row.Cells
                '    RichTextBox1.AppendText(cell.Value.ToString & "  ")
                'Next
                'RichTextBox1.AppendText(vbCrLf)

            End If
            总筛选行数 += 1
        Next


        If 是否输出到新表 = True Then
            Dim resultSheet As Excel.Worksheet = 新建工作表("筛选_" & sheet.Name, True, False, sheet)
            筛选出的区域.Copy()
            resultSheet.Cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

            自动列宽(resultSheet)

        Else
            筛选出的区域.Select()
        End If



        Dim inf As String = "总筛选行数：" & 总筛选行数 & vbCrLf &
                                "筛选通过数：" & 筛选器.通过行数 & vbCrLf &
                                "筛选未通过数：" & 筛选器.未通过行数 & vbCrLf &
                                "存在问题行数：" & 筛选器.出错行数

        流水信息.记录信息(vbCrLf & inf)


        RichTextBox1.Clear()
        RichTextBox1.AppendText("---------------------------------------------" & vbCrLf)
        RichTextBox1.AppendText(inf)
        'app.ScreenUpdating = True
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'Dim 通过数 As Integer = 0
        'Dim 未通过数 As Integer = 0
        'For Each row As Excel.Range In 获取工作区域(sheet).Rows
        '    If 筛选器.条件检测(row) Then
        '        通过数 += 1
        '        For Each cell As Excel.Range In row.Cells
        '            RichTextBox1.AppendText(cell.Value.ToString & "  ")
        '        Next
        '        RichTextBox1.AppendText(vbCrLf)
        '    Else
        '        未通过数 += 1
        '    End If
        'Next
        'RichTextBox1.AppendText(vbCrLf & vbCrLf)
        'RichTextBox1.AppendText("---------------------------------------------")
        'RichTextBox1.AppendText("通过数" & 通过数 & "   " & "未通过数" & 未通过数)
        工作表筛选内容(待筛选表, NumericUpDown1.Value, 筛选器, CheckBox1.Checked, True)
        流水信息.显示信息()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        工作表筛选数目(待筛选表, NumericUpDown1.Value, 筛选器)
    End Sub


    Public Function 工作表筛选数目(sheet As Excel.Worksheet, 列表头尾行号 As Integer, 筛选条件 As 多列条件类) As Integer
        Dim 公式文本 As String = ""
        Dim 当前活动单元格 As Excel.Range = app.ActiveCell

        Dim 工作区 As Excel.Range = 获取工作区域(sheet)

        Dim 总行数 As Integer = 工作区.Rows.Count
        Dim 总列数 As Integer = 工作区.Columns.Count
        Dim 数据行数 As Integer = 总行数 - 列表头尾行号
        Dim 列表头区 As Excel.Range = 工作区(1, 1).Resize(列表头尾行号, 总列数)
        Dim 行数据区 As Excel.Range = 工作区(列表头尾行号 + 1, 1).Resize(数据行数, 总列数)




        For Each 单列条件 As 单列条件类 In 筛选器.单列条件序列
            If 单列条件.逻辑 = 逻辑运算.且 Then
                For Each 条件 As 条件类 In 单列条件.条件序列
                    '公式文本 &= "*(" & 获取地址(行数据区.Columns(单列条件.列号), True, True, True) & 条件.ToString & ")"
                    公式文本 &= "*" & 单元格按照条件转换成数组的公式(获取地址(行数据区.Columns(单列条件.列号), True, True, True), 条件)
                Next
            Else
                公式文本 &= "*("
                For Each 条件 As 条件类 In 单列条件.条件序列
                    '公式文本 &= "(" & 获取地址(行数据区.Columns(单列条件.列号), True, True, True) & 条件.ToString & ")+"
                    公式文本 &= "(" & 单元格按照条件转换成数组的公式(获取地址(行数据区.Columns(单列条件.列号), True, True, True), 条件) & ")+"
                Next
                If 公式文本.EndsWith("+") Then
                    公式文本 = 公式文本.TrimEnd("+"c)
                End If
                公式文本 &= ")"
            End If
        Next
        '包含条件时可用这个公式=IF(ISNUMBER(SEARCH("3", E2:L2)), TRUE, FALSE)

        If 公式文本.StartsWith("*") Then
            公式文本 = 公式文本.TrimStart("*"c)
        End If
        公式文本 = "=SUMPRODUCT(IF(" & 公式文本 & "," & "1,0)"
        公式文本 &= ")"
        插入公式(公式文本, 当前活动单元格, True)


    End Function


    Public Function 单元格按照条件转换成数组的公式(区域地址 As String, 条件 As 条件类) As String
        Dim 公式文本 As String = ""
        If 条件.关系 <> 关系运算.包含于 And 条件.关系 <> 关系运算.包含 Then
            公式文本 = "(" & 区域地址 & 条件.ToString & ")"
        ElseIf 条件.关系 = 关系运算.包含 Then
            公式文本 = "ISNUMBER(SEARCH(""" & 条件.字符串值.Replace(vbCrLf, "") & """," & 区域地址 & "))"
        Else
            公式文本 = "ISNUMBER(SEARCH(" & 区域地址 & ",""" & 条件.字符串值.Replace(vbCrLf, "") & """))"
        End If

        Return 公式文本


    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        工作表筛选内容(待筛选表, NumericUpDown1.Value, 筛选器, CheckBox1.Checked, False)
        流水信息.显示信息()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MsgBox("这里还在施工，请绕行。")
    End Sub

    Private Sub NumericUpDown1_Click(sender As Object, e As EventArgs) Handles NumericUpDown1.Click
        加载列标题(待筛选表, NumericUpDown1.Value, DataGridView1)
    End Sub
End Class
