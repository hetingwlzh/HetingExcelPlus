
Imports System.Drawing
Imports System.Windows.Forms
Imports MyExcel插件.编号控件

Public Class 编号控件

    Public 当前编号方案 As 编号方案

    Public 新标志位规则 As 标志位规则
    Public 新常量规则 As 常量规则
    Public 新引用规则 As 引用规则
    Public 新分类编号规则 As 分类编号规则
    Public 新整体编号规则 As 整体编号规则

    Public 当前表 As Excel.Worksheet
    Public 编号方案保存表名 As String = "编号方案存储表_heting"


    'Public Class 错误管理
    '    Public 错误信息序列 As System.Collections.ArrayList
    '    Public Sub New()
    '        错误信息序列 = New System.Collections.ArrayList
    '    End Sub
    '    Public Sub 记录错误信息(信息 As String)
    '        错误信息序列.Add(信息)
    '    End Sub
    '    Public Sub 显示错误信息(Optional 标题 As String = "错误信息")

    '        If 错误信息序列.Count > 0 Then
    '            Dim n As Integer = 1
    '            Dim 总文本 As String = "方案：" & 标题 & vbCrLf
    '            For Each str As String In 错误信息序列
    '                总文本 &= n & "、" & str & vbCrLf
    '                n += 1
    '            Next
    '            MsgBox(总文本)
    '        Else
    '            MsgBox("方案：" & 标题 & vbCrLf & "没有错误！！",, "方案：" & 标题)
    '        End If
    '        清除错误()
    '    End Sub
    '    Public Sub 清除错误()
    '        错误信息序列.Clear()
    '    End Sub
    'End Class

    Public Class 元组
        Public 健 As String
        Public 值 As String
        Public Sub New(key As String, value As String)

            健 = key
            值 = value
        End Sub
        Public Sub New()

            健 = ""
            值 = ""
        End Sub

        Public Overrides Function ToString() As String
            Return "(" & 健 & "≡" & 值 & ")"
        End Function
        Public Function CreateFromString(str As String)
            If str IsNot Nothing Then
                str = str.Trim("(").Trim(")")
                Dim result() As String = str.Split("≡")
                If result.Length > 1 Then
                    Return New 元组(result(0), result(1))
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If


        End Function
    End Class
    Public Class 编号方案
        Public 编号方案名称 As String '方案编号的名字
        Public 整体序号 As Integer = 0 '方案整体编号时的当前整体编号值
        Public 工作表 As Excel.Worksheet '编号方案要实施的表对象
        Public 工作表的列标题所在行号 As Integer '工作表的列标题所在行号
        'Public 错误管理者 As New 错误管理
        Public 行对象筛选序列 As New System.Collections.ArrayList '用来存放筛需要编号的行的筛选条件数据。
        Public 编号所在的列索引 As Integer = 1 '生成的编号要显示在的列号
        Public 编号规则序列 As System.Collections.ArrayList '编号方案的各个编号规则序列


        Public Sub New()
            编号方案名称 = "未命名"
            编号规则序列 = New System.Collections.ArrayList
        End Sub
        Public Sub New(编号方案名称_ As String)
            编号方案名称 = 编号方案名称_
            编号规则序列 = New System.Collections.ArrayList
        End Sub
        Public Sub Clear()
            编号规则序列.Clear()
        End Sub
        Public Function Add(_标志位规则 As 标志位规则) As Integer
            编号规则序列.Add(_标志位规则)
            Return 编号规则序列.Count - 1
        End Function
        Public Function Add(_常量规则 As 常量规则) As Integer
            编号规则序列.Add(_常量规则)
            Return 编号规则序列.Count - 1
        End Function
        Public Function Add(_引用规则 As 引用规则) As Integer
            编号规则序列.Add(_引用规则)
            Return 编号规则序列.Count - 1
        End Function
        Public Function Add(_分类编号规则 As 分类编号规则) As Integer
            编号规则序列.Add(_分类编号规则)
            Return 编号规则序列.Count - 1
        End Function
        Public Function Add(_整体编号规则 As 整体编号规则) As Integer
            编号规则序列.Add(_整体编号规则)
            Return 编号规则序列.Count - 1
        End Function

        ''' <summary>
        ''' 根据指定的数据行,按照编号规则生成唯一编号
        ''' </summary>
        ''' <param name="数据行">需要生成编号的Excel数据行</param>
        ''' <returns>生成的唯一编号字符串</returns>
        Public Function 申请编号(数据行 As Excel.Range) As String
            Try

                Dim 最终编号 As String = ""
                整体序号 += 1
                For Each 规则 As Object In 编号规则序列
                    If 规则.GetType.Name = "标志位规则" Then
                        Dim temp As 标志位规则 = CType(规则, 标志位规则)
                        If 数据行.Cells(1, temp.标志位所在列号).value = Nothing Then
                            流水信息.记录信息("标志位   数据表：" & 工作表.Name & " 的第 " & 数据行.Cells(1, temp.标志位所在列号).row & " 行，第 " & temp.标志位所在列号 & " 列 为空值！")
                        Else
                            最终编号 &= temp.获取编号(数据行.Cells(1, temp.标志位所在列号).value)
                        End If


                    ElseIf 规则.GetType.Name = "常量规则" Then
                        Dim temp As 常量规则 = CType(规则, 常量规则)
                        最终编号 &= temp.获取编号()

                    ElseIf 规则.GetType.Name = "引用规则" Then
                        Dim temp As 引用规则 = CType(规则, 引用规则)
                        If 数据行.Cells(1, temp.引用列所在列号).value = Nothing Then
                            流水信息.记录信息("引用   数据表：" & 工作表.Name & " 的第 " & 数据行.Cells(1, temp.引用列所在列号).row & " 行，第 " & temp.引用列所在列号 & " 列 为空值！")
                        Else
                            最终编号 &= temp.获取编号(数据行.Cells(1, temp.引用列所在列号).value)
                        End If


                    ElseIf 规则.GetType.Name = "分类编号规则" Then
                        Dim temp As 分类编号规则 = CType(规则, 分类编号规则)
                        Dim 行特征字符串 As String = ""
                        For Each n As Integer In temp.分类列号序列
                            If 数据行.Cells(1, n).value = Nothing Then
                                流水信息.记录信息("分类编号   数据表：" & 工作表.Name & " 的第 " & 数据行.Cells(1, n).row & " 行，第 " & n & " 列 为空值！")
                            Else
                                行特征字符串 &= 数据行.Cells(1, n).value & 特征码连结字符
                            End If
                        Next
                        最终编号 &= temp.获取编号(行特征字符串, 数据行.Row)


                    ElseIf 规则.GetType.Name = "整体编号规则" Then
                        Dim temp As 整体编号规则 = CType(规则, 整体编号规则)
                        最终编号 &= temp.获取编号(整体序号 + temp.起始编号 - 1)


                    End If
                Next

                Return 最终编号
            Catch ex As Exception

                MsgBox("编号方案：" & 编号方案名称 & " 编号时出错！！！请请检查后重试！")
                整体序号 -= 1
                Return Nothing

            End Try

        End Function

        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<编号方案名称>" & 编号方案名称 & "</编号方案名称>"
            result &= "<整体序号>" & 0 & "</整体序号>"
            result &= "<工作表>" & 工作表.Name & "</工作表>"
            result &= "<工作表的列标题所在行号>" & 工作表的列标题所在行号 & "</工作表的列标题所在行号>"
            result &= "<行对象筛选序列>" & ListToString(行对象筛选序列, 一级列表元素分隔符) & "</行对象筛选序列>"
            result &= "<编号所在的列索引>" & 编号所在的列索引 & "</编号所在的列索引>"

            result &= "<编号规则序列>" & ListToString(编号规则序列， 二级列表元素分隔符) & "</编号规则序列>"






            Return result
        End Function
        Public Function CreatFromString(str As String) As 编号方案
            Dim resulr As New 编号方案

            resulr.编号方案名称 = GetValueFromString(str, "编号方案名称")
            resulr.整体序号 = GetValueFromString(str, "整体序号")
            resulr.工作表 = app.Worksheets(GetValueFromString(str, "工作表"))
            resulr.工作表的列标题所在行号 = GetValueFromString(str, "工作表的列标题所在行号")





            resulr.行对象筛选序列.Clear()
            For Each text As String In StringToList(GetValueFromString(str, "行对象筛选序列")， 一级列表元素分隔符)
                resulr.行对象筛选序列.Add(New 元组().CreateFromString(text))
            Next

            resulr.编号所在的列索引 = GetValueFromString(str, "编号所在的列索引")




            resulr.编号规则序列.Clear()
            Dim 规则字符序列 As String() = StringToList(GetValueFromString(str, "编号规则序列"), 二级列表元素分隔符)
            For Each text As String In 规则字符序列
                Select Case GetValueFromString(text, "规则类型")
                    Case "标志位规则"
                        resulr.编号规则序列.Add(New 标志位规则().CreatFromString(text))
                    Case "常量规则"
                        resulr.编号规则序列.Add(New 常量规则().CreatFromString(text))
                    Case "引用规则"
                        resulr.编号规则序列.Add(New 引用规则().CreatFromString(text))
                    Case "分类编号规则"
                        resulr.编号规则序列.Add(New 分类编号规则().CreatFromString(text))
                    Case "整体编号规则"
                        resulr.编号规则序列.Add(New 整体编号规则().CreatFromString(text))
                    Case Else
                        流水信息.记录信息("导入代码格式可能有误，没能成功加载！")
                        Return Nothing
                End Select
            Next
            If resulr.编号方案名称 Is Nothing Or
                   (Not IsNumeric(resulr.整体序号)) Or
                   resulr.工作表 Is Nothing Or
                   resulr.工作表的列标题所在行号 = Nothing Or
                   resulr.编号规则序列.Count = 0 Then
                流水信息.记录信息("导入代码格式可能有误，没能成功加载！")
                Return Nothing
            End If

            Return resulr
        End Function





    End Class
    Public Class 标志位规则

        Public 标志位所在列号 As Integer
        Public 标志位标题 As String
        Public 标志位编号序列 As New System.Collections.ArrayList

        Public Sub New()
            标志位所在列号 = 1
            标志位标题 = ""
            标志位编号序列.Clear()
        End Sub
        Public Sub New(标志位所在列号_ As Integer, 标志位标题_ As String)
            标志位所在列号 = 标志位所在列号_
            标志位标题 = 标志位标题_

        End Sub
        Public Sub New(标志位所在列号_ As Integer, 标志位标题_ As String,
                           标志位值 As String, 编号 As String)
            标志位所在列号 = 标志位所在列号_
            标志位标题 = 标志位标题_
            add(标志位值, 编号)
        End Sub
        Public Sub New(标志位所在列号_ As Integer, 标志位标题_ As String,
                           标志位值1 As String, 编号1 As String，
                           标志位值2 As String, 编号2 As String)
            标志位所在列号 = 标志位所在列号_
            标志位标题 = 标志位标题_
            add(标志位值1, 编号1)
            add(标志位值2, 编号2)
        End Sub

        Public Sub New(标志位所在列号_ As Integer, 标志位标题_ As String,
                           标志位值1 As String, 编号1 As String，
                           标志位值2 As String, 编号2 As String，
                           标志位值3 As String, 编号3 As String)
            标志位所在列号 = 标志位所在列号_
            标志位标题 = 标志位标题_
            add(标志位值1, 编号1)
            add(标志位值2, 编号2)
            add(标志位值3, 编号3)
        End Sub

        Public Sub New(标志位所在列号_ As Integer, 标志位标题_ As String,
                           标志位值1 As String, 编号1 As String，
                           标志位值2 As String, 编号2 As String，
                           标志位值3 As String, 编号3 As String，
                           标志位值4 As String, 编号4 As String)
            标志位所在列号 = 标志位所在列号_
            标志位标题 = 标志位标题_
            add(标志位值1, 编号1)
            add(标志位值2, 编号2)
            add(标志位值3, 编号3)
            add(标志位值4, 编号4)
        End Sub

        Public Sub add(标志位值 As String, 编号 As String)
            Dim xy As New 元组(标志位值, 编号)
            标志位编号序列.Add(xy)

        End Sub
        Public Function 获取编号(标志位值 As String) As String
            For Each t As 元组 In 标志位编号序列
                If t.健 = 标志位值 Then
                    Return t.值
                End If
            Next
            Return Nothing

        End Function

        Public Function 获取标志位值(编号 As String) As String
            For Each t As 元组 In 标志位编号序列
                If t.值 = 编号 Then
                    Return t.健
                End If
            Next
            Return Nothing

        End Function
        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<规则>"
            result &= "<规则类型>" & "标志位规则" & "</规则类型>"
            result &= "<标志位所在列号>" & 标志位所在列号 & "</标志位所在列号>"
            result &= "<标志位标题>" & 标志位标题 & "</标志位标题>"
            result &= "<标志位编号序列>"
            result &= ListToString(标志位编号序列, 一级列表元素分隔符)
            result &= "</标志位编号序列>"
            result &= "</规则>"
            Return result
        End Function
        Public Function CreatFromString(str As String) As 标志位规则
            Dim resulr As New 标志位规则

            resulr.标志位所在列号 = GetValueFromString(str, "标志位所在列号")
            resulr.标志位标题 = GetValueFromString(str, "标志位标题")
            resulr.标志位编号序列.Clear()
            For Each text As String In StringToList(GetValueFromString(str, "标志位编号序列"), 一级列表元素分隔符)
                resulr.标志位编号序列.Add(New 元组().CreateFromString(text))
            Next
            Return resulr

        End Function



    End Class


    Public Class 常量规则
        Public 常量值 As String
        Public Sub New()
            常量值 = ""
        End Sub
        Public Sub New(常量的值 As String)
            常量值 = 常量的值
        End Sub
        Public Function 获取常量值()
            Return 常量值
        End Function
        Public Function 获取编号()
            Return 常量值
        End Function


        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<规则>"
            result &= "<规则类型>" & "常量规则" & "</规则类型>"
            result &= "<常量值>" & 常量值 & "</常量值>"
            result &= "</规则>"
            Return result
        End Function
        Public Function CreatFromString(str As String) As 常量规则
            Dim resulr As New 常量规则
            resulr.常量值 = GetValueFromString(str, "常量值")
            Return resulr
        End Function


    End Class

    Public Class 引用规则
        Public 引用列所在列号 As Integer
        Public 引用列标题 As String
        Public 引用列限制位数 As Integer
        Public 补位字符 As String
        Public Sub New()
        End Sub
        Public Sub New(_引用列所在列号 As Integer, _引用列标题 As String, _引用列限制位数 As Integer, _补位字符 As String)
            引用列所在列号 = _引用列所在列号
            引用列标题 = _引用列标题
            引用列限制位数 = _引用列限制位数
            补位字符 = _补位字符
        End Sub
        Public Function 获取编号（值 As String)
            Return 值.PadLeft(引用列限制位数, 补位字符)
        End Function
        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<规则>"
            result &= "<规则类型>" & "引用规则" & "</规则类型>"
            result &= "<引用列所在列号>" & 引用列所在列号 & "</引用列所在列号>"
            result &= "<引用列标题>" & 引用列标题 & "</引用列标题>"
            result &= "<引用列限制位数>" & 引用列限制位数 & "</引用列限制位数>"
            result &= "<补位字符>" & 补位字符 & "</补位字符>"
            result &= "</规则>"
            Return result
        End Function
        Public Function CreatFromString(str As String) As 引用规则
            Dim resulr As New 引用规则
            resulr.引用列所在列号 = GetValueFromString(str, "引用列所在列号")
            resulr.引用列标题 = GetValueFromString(str, "引用列标题")
            resulr.引用列限制位数 = GetValueFromString(str, "引用列限制位数")
            resulr.补位字符 = GetValueFromString(str, "补位字符")
            Return resulr
        End Function

    End Class


    Public Class 分类编号规则
        Public 分类列号序列 As System.Collections.ArrayList
        Public 限制位数 As Integer
        Public 补位字符 As String
        Public 起始编号 As Integer
        Public 分类序号记录表 As New System.Collections.ArrayList '其内元素格式：{行特征代码，当前尾编号，此特征代码的首行索引，此特征代码的尾行索引}


        Public Sub New()
            分类列号序列 = New System.Collections.ArrayList
            限制位数 = 0
            补位字符 = 0
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String)
            分类列号序列 = New System.Collections.ArrayList
            限制位数 = _限制位数
            补位字符 = _补位字符
            起始编号 = 1
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String, _起始编号 As Integer)
            限制位数 = _限制位数
            补位字符 = _补位字符
            起始编号 = _起始编号
            分类列号序列 = New System.Collections.ArrayList
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String, _分类列号序列 As System.Collections.ArrayList)
            分类列号序列 = _分类列号序列
            限制位数 = _限制位数
            补位字符 = _补位字符
            起始编号 = 1
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String, _起始编号 As Integer, _分类列号序列 As System.Collections.ArrayList)
            限制位数 = _限制位数
            补位字符 = _补位字符
            起始编号 = _起始编号
            分类列号序列 = _分类列号序列
        End Sub

        Public Function 获取编号(行特征代码 As String, 行索引 As Integer)
            For Each cell As System.Collections.ArrayList In 分类序号记录表
                '分类序号记录表  中每个记录均为 System.Collections.ArrayList 类型，
                '格式：{行特征代码，当前尾编号，此特征代码的首行索引，此特征代码的尾行索引}
                If cell(0) = 行特征代码 Then
                    cell(1) += 1 '当前尾编号加一
                    cell(3) = 行索引 '尾行号
                    Return cell(1).ToString.PadLeft(限制位数, 补位字符)
                End If
            Next
            Dim temp As New System.Collections.ArrayList  'New 值号(行特征代码, 起始编号)
            temp.Add(行特征代码)
            temp.Add(起始编号) '当前尾编号
            temp.Add(行索引) '首行号
            temp.Add(行索引) '尾行号
            分类序号记录表.Add(temp)
            Return 起始编号.ToString.PadLeft(限制位数, 补位字符)
        End Function


        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<规则>"
            result &= "<规则类型>" & "分类编号规则" & "</规则类型>"
            result &= "<分类列号序列>" & ListToString(分类列号序列, 一级列表元素分隔符) & "</分类列号序列>"
            result &= "<限制位数>" & 限制位数 & "</限制位数>"
            result &= "<补位字符>" & 补位字符 & "</补位字符>"
            result &= "<起始编号>" & 起始编号 & "</起始编号>"
            result &= "</规则>"
            Return result
        End Function
        Public Function CreatFromString(str As String) As 分类编号规则
            Dim resulr As New 分类编号规则
            For Each num As String In StringToList(GetValueFromString(str, "分类列号序列"), 一级列表元素分隔符)
                resulr.分类列号序列.Add(Int(num))
            Next
            resulr.限制位数 = GetValueFromString(str, "限制位数")
            resulr.补位字符 = GetValueFromString(str, "补位字符")
            resulr.起始编号 = GetValueFromString(str, "起始编号")
            resulr.分类序号记录表.Clear()
            Return resulr
        End Function





    End Class
    Public Class 整体编号规则
        Public 限制位数 As Integer
        Public 补位字符 As String
        Public 起始编号 As Integer
        Public Sub New()
            限制位数 = 0
            补位字符 = 0
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String)
            限制位数 = _限制位数
            补位字符 = _补位字符
        End Sub
        Public Sub New(_限制位数 As Integer, _补位字符 As String, _起始编号 As Integer)
            限制位数 = _限制位数
            补位字符 = _补位字符
            起始编号 = _起始编号
        End Sub

        Public Function 获取编号（序号 As String)
            Return 序号.PadLeft(限制位数, 补位字符)
        End Function

        Public Overrides Function ToString() As String
            Dim result As String = ""
            result &= "<规则>"
            result &= "<规则类型>" & "整体编号规则" & "</规则类型>"
            result &= "<限制位数>" & 限制位数 & "</限制位数>"
            result &= "<补位字符>" & 补位字符 & "</补位字符>"
            result &= "<起始编号>" & 起始编号 & "</起始编号>"
            result &= "</规则>"
            Return result
        End Function
        Public Function CreatFromString(str As String) As 整体编号规则
            Dim resulr As New 整体编号规则
            resulr.限制位数 = GetValueFromString(str, "限制位数")
            resulr.补位字符 = GetValueFromString(str, "补位字符")
            resulr.起始编号 = GetValueFromString(str, "起始编号")
            Return resulr
        End Function
    End Class

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '加载列标题(当前表, NumericUpDown1.Value, DataGridView1)
        Try

            For Each 方案 As 编号方案 In 编号方案列表
                If 方案.编号方案名称 = ComboBox1.Text Then
                    当前编号方案 = 方案
                    选中指定命令(ComboBox2, 当前编号方案.工作表.Name)
                    For Each temp As 元组 In 当前编号方案.行对象筛选序列
                        DataGridView1.Rows(temp.健).Cells(1).Value = temp.值
                    Next



                    刷新方案信息()
                End If
            Next


        Catch ex As Exception

            MsgBox("似乎表格已经删除！")


        End Try




    End Sub

    Public Function 选中指定命令(ComboBox_ As Windows.Forms.ComboBox, 命令文本 As String) As Integer
        For i As Integer = 0 To ComboBox_.Items.Count - 1
            If ComboBox_.Items(i).ToString = 命令文本 Then
                ComboBox_.SelectedIndex = i
                Return i
            End If
        Next

    End Function
    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If ComboBox2.Text <> "" Then
            If 是否存在工作表(ComboBox2.Text) Then
                加载列标题(app.Sheets(ComboBox2.Text), NumericUpDown1.Value, DataGridView1)
            End If

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim 新编号方案 As New 编号方案(ComboBox1.Text)

        If 是否存在工作表(ComboBox2.Text) Then

            新编号方案.工作表 = app.Sheets(ComboBox2.Text)
            新编号方案.工作表的列标题所在行号 = NumericUpDown1.Value
            Dim temp As 元组
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If DataGridView1.Rows(i).Cells(1).Value <> "" Then
                    temp = New 元组(i, DataGridView1.Rows(i).Cells(1).Value)
                    新编号方案.行对象筛选序列.Add(temp)
                End If

            Next




        Else
            MsgBox("请选择编号方案所对应的工作表 先！")
            Exit Sub
        End If
        当前编号方案 = 新编号方案

        编号方案列表.Add(新编号方案)

        刷新方案树()


        TabControl1.SelectedIndex = 1

        Button7.Visible = True


        My.Settings.Save()
    End Sub
    Public Sub 刷新方案信息()
        Dim 当前方案名, 规则个数 As String
        If 当前编号方案 Is Nothing Then
            当前方案名 = "Nothing"
            规则个数 = 0
        Else
            当前方案名 = 当前编号方案.编号方案名称
            规则个数 = 当前编号方案.编号规则序列.Count
        End If
        Label5.Text = "方案个数：" & 编号方案列表.Count & vbCrLf & " 当前方案：" & 当前方案名 & vbCrLf & " 规则个数：" & 规则个数
    End Sub

    Public Sub 刷新方案树(Optional 要选中的项目索引 As Integer = -1)
        ComboBox1.Items.Clear()
        清理方案树()
        For Each 方案 As 编号方案 In 编号方案列表
            If 方案 IsNot Nothing Then
                ComboBox1.Items.Add(方案.编号方案名称)
                添加方案节点(方案)
            End If
        Next

        If 要选中的项目索引 >= 0 Then

            For Each node As Windows.Forms.TreeNode In TreeView1.Nodes(0).Nodes
                If node.Text = 当前编号方案.编号方案名称 Then
                    TreeView1.SelectedNode = node.Nodes(要选中的项目索引)
                    node.Nodes(要选中的项目索引).Text &= "◀"
                    'node.Nodes(要选中的项目索引).Checked = True
                End If
            Next
        End If
        刷新方案信息()

    End Sub

    Public Sub 清理方案树()
        TreeView1.Nodes(0).Nodes.Clear()
    End Sub
    Public Function 添加方案节点(方案 As 编号方案) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = 方案.编号方案名称
        TreeView1.Nodes(0).Nodes.Add(node)
        For Each t As Object In 方案.编号规则序列
            添加规则节点(node, t)
        Next
        TreeView1.Nodes(0).Expand()
        Return node
    End Function

    Public Function 添加规则节点(节点 As Windows.Forms.TreeNode， 规则 As 常量规则) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = "[常量]" & 规则.常量值
        节点.Nodes.Add(node)
        节点.Expand()
        Return node
    End Function


    Public Function 添加规则节点(节点 As Windows.Forms.TreeNode， 规则 As 引用规则) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = "[引用]" & 规则.引用列标题
        节点.Nodes.Add(node)
        节点.Expand()
        Return node
    End Function

    Public Function 添加规则节点(节点 As Windows.Forms.TreeNode， 规则 As 标志位规则) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = "[标志]" & 规则.标志位标题
        节点.Nodes.Add(node)
        节点.Expand()
        Return node
    End Function

    Public Function 添加规则节点(节点 As Windows.Forms.TreeNode， 规则 As 分类编号规则) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = "[类号]" & 规则.起始编号.ToString.PadLeft(规则.限制位数, 规则.补位字符)
        节点.Nodes.Add(node)
        节点.Expand()
        Return node
    End Function

    Public Function 添加规则节点(节点 As Windows.Forms.TreeNode， 规则 As 整体编号规则) As Windows.Forms.TreeNode
        Dim node = New Windows.Forms.TreeNode
        node.Text = "[序号]" & 规则.起始编号.ToString.PadLeft(规则.限制位数, 规则.补位字符)
        节点.Nodes.Add(node)
        节点.Expand()
        Return node
    End Function



    Private Sub 编号控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表(ComboBox2)
        'DataGridView2.Columns(0).ReadOnly = True
        刷新方案树()
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
        End If
        检测是否有可导入的编号规则()


    End Sub
    Public Function 检测是否有可导入的编号规则() As Boolean
        If 是否存在工作表(编号方案保存表名) Then
            Label5.ForeColor = Color.Red
            Label5.Text = "当前文件储存有编号规则，可尝试导入！"
            Button9.ForeColor = Color.Red
            Button9.Text = "可导入"
        Else
            Label5.ForeColor = Color.DarkGreen
            Button9.ForeColor = Color.Black
            Button9.Text = "导入规则"
        End If

    End Function

    Public Sub 加载列标题(sheet As Excel.Worksheet, 标题行行号 As Integer, MyDataGridView As Windows.Forms.DataGridView)


        If sheet IsNot Nothing Then

            MyDataGridView.Columns.Clear()
            CheckedListBox1.Items.Clear()
            MyDataGridView.Columns.Add("列标题", "列标题")
            MyDataGridView.Columns.Add("匹配值", "匹配值")
            MyDataGridView.Columns(0).ReadOnly = True
            Dim str As String = ""

            If 标题行行号 > sheet.UsedRange.Rows.Count Then
                MsgBox("标题行行号超出使用区范围，请重新设置标题行所在区域的行号。")
                Exit Sub
            End If
            Dim n As Integer = 0
            For Each cell As Excel.Range In sheet.UsedRange.Rows.Item(标题行行号).Cells
                n = MyDataGridView.Rows.Add()
                MyDataGridView.Rows(n).Cells(0).Value = ToStr(cell)
                CheckedListBox1.Items.Add(ToStr(cell))
                MyDataGridView.Rows(n).Cells(1).Value = ""





            Next
        End If




    End Sub




    'Public Sub 刷新规则序列()
    '    DataGridView2.Rows.Clear()
    '    For Each 规则 As Object In 当前编号方案.编号规则序列
    '        If 规则.GetType.Name = "标志位规则" Then
    '            Dim temp As 标志位规则 = CType(规则, 标志位规则)


    '            Dim n As Integer = DataGridView2.Rows.Add()
    '            DataGridView2.Rows(n).Cells(0).Value = temp.标志位所在列号
    '            DataGridView2.Rows(n).Cells(1).Value = temp.标志位标题
    '            DataGridView2.Rows(n).Cells(2).Value = 规则.GetType.Name

    '        ElseIf 规则.GetType.Name = "常量规则" Then
    '            Dim temp As 常量规则 = CType(规则, 常量规则)
    '            Dim n As Integer = DataGridView2.Rows.Add()
    '            'DataGridView2.Rows(n).Cells(0).Value = CheckedListBox1.SelectedIndex + 1
    '            DataGridView2.Rows(n).Cells(1).Value = temp.获取常量值
    '            DataGridView2.Rows(n).Cells(2).Value = 规则.GetType.Name

    '        ElseIf 规则.GetType.Name = "引用规则" Then
    '            Dim temp As 引用规则 = CType(规则, 引用规则)
    '            Dim n As Integer = DataGridView2.Rows.Add()
    '            DataGridView2.Rows(n).Cells(0).Value = temp.引用列所在列号
    '            DataGridView2.Rows(n).Cells(1).Value = temp.引用列标题
    '            DataGridView2.Rows(n).Cells(2).Value = 规则.GetType.Name

    '        ElseIf 规则.GetType.Name = "分类编号规则" Then
    '            Dim temp As 分类编号规则 = CType(规则, 分类编号规则)
    '            Dim n As Integer = DataGridView2.Rows.Add()
    '            'DataGridView2.Rows(n).Cells(0).Value = CheckedListBox1.SelectedIndex + 1
    '            DataGridView2.Rows(n).Cells(1).Value = "1".PadLeft(temp.限制位数, temp.补位字符)
    '            DataGridView2.Rows(n).Cells(2).Value = 规则.GetType.Name

    '        ElseIf 规则.GetType.Name = "整体编号规则" Then
    '            Dim temp As 整体编号规则 = CType(规则, 整体编号规则)
    '            Dim n As Integer = DataGridView2.Rows.Add()
    '            'DataGridView2.Rows(n).Cells(0).Value = CheckedListBox1.SelectedIndex + 1
    '            DataGridView2.Rows(n).Cells(1).Value = "1".PadLeft(temp.限制位数, temp.补位字符)
    '            DataGridView2.Rows(n).Cells(2).Value = 规则.GetType.Name


    '        End If
    '    Next
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim 新添加的规则索引 As Integer
        If CheckedListBox2.SelectedItem IsNot Nothing Then

            If CheckedListBox2.SelectedItem.ToString = "常量规则" Then
                保存当前规则设置(新常量规则)
                新添加的规则索引 = 当前编号方案.Add(新常量规则)

            ElseIf CheckedListBox2.SelectedItem.ToString = "引用规则" Then

                保存当前规则设置(新引用规则)
                新添加的规则索引 = 当前编号方案.Add(新引用规则)


            ElseIf CheckedListBox2.SelectedItem.ToString = "标志位规则" Then
                保存当前规则设置(新标志位规则)
                新添加的规则索引 = 当前编号方案.Add(新标志位规则)

            ElseIf CheckedListBox2.SelectedItem.ToString = "分类编号规则" Then
                保存当前规则设置(新分类编号规则)
                新添加的规则索引 = 当前编号方案.Add(新分类编号规则)
            ElseIf CheckedListBox2.SelectedItem.ToString = "整体编号规则" Then
                保存当前规则设置(新整体编号规则)
                新添加的规则索引 = 当前编号方案.Add(新整体编号规则)


            End If

            创建默认新规则()
            '保存当前规则设置(当前编号方案.编号规则序列(新添加的规则索引))
            '刷新规则序列()
            刷新方案树(新添加的规则索引)
            TreeView1.SelectedNode = Nothing
        Else

            MsgBox("请选择要添加的规则类型，并设置后正在添加。")


        End If


    End Sub

    Private Sub CheckedListBox1_Click(sender As Object, e As EventArgs) Handles CheckedListBox1.Click
        '此过程的目的在于当选择 分类标号规则  时下方的选框必须能够多选。其他情况不能多选。

        If CheckedListBox2.SelectedItem Is Nothing Then
            Dim n As Integer = CheckedListBox1.SelectedIndex
            If n <> -1 Then
                For i = 0 To CheckedListBox1.Items.Count - 1
                    sender.SetItemChecked(i, False)
                Next
                'CheckedListBox1.SetItemChecked(n, True)
                Dim 列 As Excel.Range = 当前编号方案.工作表.UsedRange.Columns(n + 1).Cells
                Dim 列标题 As String = 列.Cells(当前编号方案.工作表的列标题所在行号, 1).value
                列 = app.Range(列.Cells(NumericUpDown1.Value + 1, 1), 列.Cells(列.Rows.Count, 1))

                刷新列元素(列, 列标题)
            End If


        Else
            If CheckedListBox2.SelectedItem.ToString <> "分类编号规则" Then
                Dim n As Integer = CheckedListBox1.SelectedIndex
                If n <> -1 Then
                    For i = 0 To CheckedListBox1.Items.Count - 1
                        sender.SetItemChecked(i, False)
                    Next
                    'CheckedListBox1.SetItemChecked(n, True)                          '

                    'If CheckedListBox2.SelectedItem.ToString = "引用规则" Or CheckedListBox2.SelectedItem.ToString = "标志位规则" Then
                    '    Dim 列 As Excel.Range = 当前编号方案.工作表.UsedRange.Columns(n + 1).Cells
                    '    Dim 列标题 As String = 列.Cells(当前编号方案.工作表的列标题所在行号, 1).value
                    '    列 = app.Range(列.Cells(NumericUpDown1.Value + 1, 1), 列.Cells(列.Rows.Count, 1))
                    '    刷新列元素(列, 列标题)
                    'End If

                End If
            End If




        End If


    End Sub



    Public Sub 刷新列元素(列区域 As Excel.Range, Optional 列标题 As String = "列标题")
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()

        DataGridView3.DefaultCellStyle.Font = New Drawing.Font("微软雅黑", 12)
        DataGridView3.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        DataGridView3.DefaultCellStyle.BackColor = System.Drawing.Color.Beige
        DataGridView3.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Red
        DataGridView3.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Yellow



        Dim 互异元素 As New System.Collections.ArrayList
        互异元素 = 枚举列中互异元素(列区域)
        Dim t As Integer = 1


        DataGridView3.Columns.Add("序号", "序号")
        DataGridView3.Columns.Add(列标题, 列标题)
        DataGridView3.Columns.Add("编号", "编号")
        DataGridView3.Columns.Item(0).Width = 100
        DataGridView3.Columns.Item(1).Width = 300
        DataGridView3.Columns.Item(2).Width = 250



        For Each value As String In 互异元素
            Dim n As Integer = DataGridView3.Rows.Add()
            DataGridView3.Rows(n).Cells(0).Value = n + 1
            DataGridView3.Rows(n).Cells(1).Value = value
            DataGridView3.Rows(n).Cells(2).Value = t
            t += 1
        Next

        'DataGridView3.Columns.Item(0).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        'DataGridView3.Columns.Item(1).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        'DataGridView3.Columns.Item(2).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells

        DataGridView3.Columns.Item(0).Width = DataGridView3.Columns.Item(0).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)
        DataGridView3.Columns.Item(1).Width = DataGridView3.Columns.Item(1).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)
        DataGridView3.Columns.Item(2).Width = DataGridView3.Columns.Item(2).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)




    End Sub

    Private Sub ComboBox2_DropDown(sender As Object, e As EventArgs) Handles ComboBox2.DropDown
        加载表(ComboBox2)
    End Sub



    Private Sub CheckedListBox2_Click(sender As Object, e As EventArgs) Handles CheckedListBox2.Click

        TreeView1.SelectedNode = Nothing

        Dim n As Integer = CheckedListBox2.SelectedIndex
        If n <> -1 Then
            For i = 0 To CheckedListBox2.Items.Count - 1
                sender.SetItemChecked(i, False)
            Next
            'CheckedListBox2.SetItemChecked(n, True)

            If CheckedListBox2.SelectedItem.ToString = "分类编号规则" Then
                For Each temp As Object In 当前编号方案.编号规则序列
                    If temp.GetType.Name = "标志位规则" Then
                        Dim 规则 As 标志位规则 = CType(temp, 标志位规则)
                        'indexList.Add(规则.标志位所在列号)
                        CheckedListBox1.SetItemChecked(规则.标志位所在列号 - 1, True)
                    End If
                Next
            End If



            创建默认新规则()



        End If
    End Sub
    ''' <summary>
    ''' 根据当前选择条件,创建默认的新编号规则对象
    ''' </summary>
    ''' <returns>新的编号规则对象</returns>
    Public Function 创建默认新规则() As Object
        If CheckedListBox2.SelectedItem.ToString = "标志位规则" Then
            If CheckedListBox1.SelectedIndex = -1 Then
                CheckedListBox1.SelectedIndex = 0
            End If
            新标志位规则 = New 标志位规则(CheckedListBox1.SelectedIndex + 1, CheckedListBox1.SelectedItem.ToString)

            Dim 互异元素 As New System.Collections.ArrayList
            Dim 列 As Excel.Range = 当前编号方案.工作表.UsedRange.Columns(新标志位规则.标志位所在列号).Cells
            列 = app.Range(列.Cells(当前编号方案.工作表的列标题所在行号 + 1, 1), 列.Cells(列.Rows.Count, 1))
            互异元素 = 枚举列中互异元素(列)
            Dim t As Integer = 1
            For Each value As String In 互异元素
                新标志位规则.add(value, t)
                t += 1
            Next
            显示当前规则设置(新标志位规则)
            Return 新标志位规则



        ElseIf CheckedListBox2.SelectedItem.ToString = "常量规则" Then
            新常量规则 = New 常量规则("2023")
            显示当前规则设置(新常量规则)
            Return 新常量规则

        ElseIf CheckedListBox2.SelectedItem.ToString = "引用规则" Then
            If CheckedListBox1.SelectedIndex = -1 Then
                CheckedListBox1.SelectedIndex = 0
            End If
            新引用规则 = New 引用规则(CheckedListBox1.SelectedIndex + 1, CheckedListBox1.SelectedItem.ToString, 4, "0")
            显示当前规则设置(新引用规则)
            Return 新引用规则

        ElseIf CheckedListBox2.SelectedItem.ToString = "分类编号规则" Then
            If CheckedListBox1.SelectedIndex = -1 Then
                CheckedListBox1.SelectedIndex = 0
            End If
            Dim indexList As New System.Collections.ArrayList


            For i = 0 To CheckedListBox1.Items.Count - 1
                If CheckedListBox1.GetItemChecked(i) = True Then
                    indexList.Add(i + 1)
                End If

            Next





            新分类编号规则 = New 分类编号规则(4, "0", 1, indexList)
            显示当前规则设置(新分类编号规则)
            Return 新分类编号规则

        ElseIf CheckedListBox2.SelectedItem.ToString = "整体编号规则" Then
            新整体编号规则 = New 整体编号规则(4, "0")
            显示当前规则设置(新整体编号规则)
            Return 新整体编号规则
        End If




    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        删除当前编号方案()

    End Sub

    Public Sub 删除当前编号方案()
        删除方案(当前编号方案)
        'ComboBox1.Items.RemoveAt(选中指定命令(ComboBox1, 当前编号方案.编号方案名称))
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
        Else
            ComboBox1.Text = ""
        End If
    End Sub
    Public Function 删除方案(方案 As 编号方案)
        For Each 方案0 As 编号方案 In 编号方案列表
            If 方案0.Equals(方案) Then
                编号方案列表.Remove(方案0)
                刷新方案树()
                Exit Function
            End If
        Next

    End Function
    Public Function 删除方案(方案名 As String)
        For Each 方案0 As 编号方案 In 编号方案列表
            If 方案0.编号方案名称 = 方案名 Then
                编号方案列表.Remove(方案0)
                刷新方案树()
                Exit Function
            End If
        Next

    End Function

    Private Sub 添加编号方案ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 添加编号方案ToolStripMenuItem.Click
        TabControl1.SelectedIndex = 0
    End Sub


    Public Sub ClickTreeView()
        CheckedListBox2.SelectedItem = Nothing
        If TreeView1.SelectedNode IsNot Nothing Then
            If TreeView1.SelectedNode.Level = 0 Then '根节点


                DataGridView3.Rows.Clear()
                For Each 方案 As 编号方案 In 编号方案列表
                    Dim n As Integer = DataGridView3.Rows.Add()
                    DataGridView3.Rows(n).Cells(1).Value = "方案 " & n
                    DataGridView3.Rows(n).Cells(2).Value = 方案.编号方案名称
                Next
                同步选择信息()
            ElseIf TreeView1.SelectedNode.Level = 1 Then '方案节点


                当前编号方案 = 获取方案(TreeView1.SelectedNode.Text)
                DataGridView3.Rows.Clear()
                For Each 规则 As Object In 当前编号方案.编号规则序列
                    Dim n As Integer = DataGridView3.Rows.Add()
                    写入行(n,, 规则.GetType.Name)
                Next
                同步选择信息()
            ElseIf TreeView1.SelectedNode.Level = 2 Then '规则节点
                当前编号方案 = 获取方案(TreeView1.SelectedNode.Parent.Text)
                同步选择信息()
                显示当前规则设置(当前编号方案.编号规则序列(TreeView1.SelectedNode.Index))

            End If
        End If



    End Sub
    Private Sub TreeView1_AfterSelect(sender As Object, e As Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        ClickTreeView()
    End Sub
    ''' <summary>
    ''' 向DataGridView控件指定行写入数据
    ''' </summary>
    ''' <param name="行索引">要写入的行索引</param>
    ''' <param name="单元格1">第一列单元格的值</param>
    ''' <param name="单元格2">第二列单元格的值</param>  
    ''' <param name="单元格3">第三列单元格的值</param>
    ''' <param name="单元格4">第四列单元格的值</param>
    Public Sub 写入行(行索引 As Integer,
                       Optional 单元格1 As String = Nothing,
                       Optional 单元格2 As String = Nothing,
                       Optional 单元格3 As String = Nothing,
                       Optional 单元格4 As String = Nothing)
        If 行索引 > DataGridView3.Rows.Count - 1 Then
            行索引 = DataGridView3.Rows.Add
        End If



        Dim 最大列数 As Integer = 0

        If 单元格1 IsNot Nothing Then
            最大列数 = 1
        End If
        If 单元格2 IsNot Nothing Then
            最大列数 = 2
        End If
        If 单元格3 IsNot Nothing Then
            最大列数 = 3
        End If
        If 单元格4 IsNot Nothing Then
            最大列数 = 4
        End If


        For i As Integer = DataGridView3.ColumnCount + 1 To 最大列数
            DataGridView3.Columns.Add("列" & i, "列" & i)
        Next




        If 单元格1 Is Nothing Then
            DataGridView3.Rows(行索引).Cells(0).Value = 行索引 + 1
        Else
            DataGridView3.Rows(行索引).Cells(0).Value = 单元格1
        End If

        If 单元格2 IsNot Nothing Then
            DataGridView3.Rows(行索引).Cells(1).Value = 单元格2
        End If


        If 单元格3 IsNot Nothing Then
            DataGridView3.Rows(行索引).Cells(2).Value = 单元格3
        End If

        If 单元格4 IsNot Nothing Then
            DataGridView3.Rows(行索引).Cells(3).Value = 单元格4
        End If

    End Sub
    ''' <summary>
    ''' 根据树视图控件的选择,同步和显示相关的编号方案信息
    ''' </summary>
    Public Sub 同步选择信息()

        Dim 方案名 As String = ""
        If TreeView1.SelectedNode Is Nothing Then
            方案名 = 当前编号方案.编号方案名称

        Else
            If TreeView1.SelectedNode.Level = 1 Then
                方案名 = TreeView1.SelectedNode.Text
            ElseIf TreeView1.SelectedNode.Level = 2 Then
                方案名 = TreeView1.SelectedNode.Parent.Text
            Else
                Exit Sub
            End If

        End If



        For Each 方案 As 编号方案 In 编号方案列表
            If 方案.编号方案名称 = 方案名 Then
                当前编号方案 = 方案
                选中指定命令(ComboBox1, 方案名)
                选中指定命令(ComboBox2, 当前编号方案.工作表.Name)

                刷新方案信息()

                加载列标题(当前编号方案.工作表, NumericUpDown1.Value, DataGridView1)
                NumericUpDown1.Value = 当前编号方案.工作表的列标题所在行号

                For Each temp As 元组 In 当前编号方案.行对象筛选序列
                    DataGridView1.Rows(temp.健).Cells(1).Value = temp.值
                Next
                Exit Sub
            End If
        Next



    End Sub

    Public Function 获取方案(方案名 As String) As 编号方案
        For Each 方案 As 编号方案 In 编号方案列表
            If 方案.编号方案名称 = 方案名 Then
                Return 方案
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' 显示指定编号规则的设置信息
    ''' </summary>
    ''' <param name="规则">编号规则对象</param>
    Public Function 显示当前规则设置(规则 As Object)
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Clear()



        DataGridView3.DefaultCellStyle.Font = New Drawing.Font("微软雅黑", 12)
        DataGridView3.DefaultCellStyle.ForeColor = System.Drawing.Color.Blue
        DataGridView3.DefaultCellStyle.BackColor = System.Drawing.Color.Beige
        DataGridView3.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Red
        DataGridView3.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Yellow




        For i = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, False)
        Next
        For i = 0 To CheckedListBox2.Items.Count - 1
            CheckedListBox2.SetItemChecked(i, False)
        Next
        'CheckedListBox1.SelectedItem = Nothing


        If 规则.GetType.Name = "标志位规则" Then



            Dim temp As 标志位规则 = CType(规则, 标志位规则)

            CheckedListBox2.SetSelected(2, True)
            CheckedListBox2.SetItemChecked(2, True)
            CheckedListBox1.SetItemChecked(temp.标志位所在列号 - 1, True)
            'CheckedListBox1.SetSelected(temp.标志位所在列号 - 1, True)

            DataGridView3.Columns.Add("序号", "序号")
            DataGridView3.Columns.Add("标志位值", "标志位值")
            DataGridView3.Columns.Add("编号", "编号")
            DataGridView3.Columns.Item(0).Width = 100
            DataGridView3.Columns.Item(1).Width = 300
            DataGridView3.Columns.Item(2).Width = 250



            DataGridView3.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序



            For Each record As 元组 In temp.标志位编号序列
                Dim n As Integer = DataGridView3.Rows.Add()
                写入行(n,, record.健, record.值)
            Next

        ElseIf 规则.GetType.Name = "常量规则" Then

            Dim temp As 常量规则 = CType(规则, 常量规则)
            CheckedListBox2.SetSelected(0, True)
            CheckedListBox2.SetItemChecked(0, True)


            DataGridView3.Columns.Add("序号", "序号")
            DataGridView3.Columns.Add("属性", "属性")
            DataGridView3.Columns.Add("值", "值")
            DataGridView3.Columns.Item(0).Width = 100
            DataGridView3.Columns.Item(1).Width = 300
            DataGridView3.Columns.Item(2).Width = 250



            DataGridView3.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            Dim n As Integer = DataGridView3.Rows.Add()
            写入行(n,, "常量值", temp.常量值)

        ElseIf 规则.GetType.Name = "引用规则" Then

            Dim temp As 引用规则 = CType(规则, 引用规则)
            CheckedListBox2.SetSelected(1, True)
            CheckedListBox2.SetItemChecked(1, True)

            CheckedListBox1.SetItemChecked(temp.引用列所在列号 - 1, True)
            'CheckedListBox1.SetSelected(temp.引用列所在列号 - 1, True)

            DataGridView3.Columns.Add("序号", "序号")
            DataGridView3.Columns.Add("属性", "属性")
            DataGridView3.Columns.Add("值", "值")
            DataGridView3.Columns.Item(0).Width = 100
            DataGridView3.Columns.Item(1).Width = 300
            DataGridView3.Columns.Item(2).Width = 250


            DataGridView3.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            Dim n As Integer = DataGridView3.Rows.Add()
            写入行(n,, "限制位数", temp.引用列限制位数)
            n = DataGridView3.Rows.Add()
            写入行(n,, "补位字符", temp.补位字符)

        ElseIf 规则.GetType.Name = "分类编号规则" Then

            Dim temp As 分类编号规则 = CType(规则, 分类编号规则)
            CheckedListBox2.SetSelected(3, True)
            CheckedListBox2.SetItemChecked(3, True)

            For Each t As Integer In temp.分类列号序列
                CheckedListBox1.SetItemChecked(t - 1, True)
                'CheckedListBox1.SetSelected(t - 1, True)
            Next



            DataGridView3.Columns.Add("序号", "序号")
            DataGridView3.Columns.Add("属性", "属性")
            DataGridView3.Columns.Add("值", "值")
            DataGridView3.Columns.Item(0).Width = 100
            DataGridView3.Columns.Item(1).Width = 300
            DataGridView3.Columns.Item(2).Width = 250



            DataGridView3.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            Dim n As Integer = DataGridView3.Rows.Add()
            写入行(n,, "限制位数", temp.限制位数)
            n = DataGridView3.Rows.Add()
            写入行(n,, "起始编号", temp.起始编号)
            n = DataGridView3.Rows.Add()
            Dim ss As String = ""


            For Each t As Integer In temp.分类列号序列
                ss &= t & ","
                'CheckedListBox1.SetItemChecked(t - 1, True)
            Next
            写入行(n,, "分类列号", ss)
            'Dim indexList As New System.Collections.ArrayList
            'For i = 0 To CheckedListBox1.Items.Count - 1
            '    If CheckedListBox1.GetItemChecked(i) = True Then
            '        indexList.Add(i)
            '    End If
            'Next


        ElseIf 规则.GetType.Name = "整体编号规则" Then

            Dim temp As 整体编号规则 = CType(规则, 整体编号规则)
            CheckedListBox2.SetSelected(4, True)
            CheckedListBox2.SetItemChecked(4, True)

            DataGridView3.Columns.Add("序号", "序号")
            DataGridView3.Columns.Add("属性", "属性")
            DataGridView3.Columns.Add("值", "值")
            DataGridView3.Columns.Item(0).Width = 100
            DataGridView3.Columns.Item(1).Width = 300
            DataGridView3.Columns.Item(2).Width = 250



            DataGridView3.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            DataGridView3.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
            Dim n As Integer = DataGridView3.Rows.Add()
            写入行(n,, "限制位数", temp.限制位数)
            n = DataGridView3.Rows.Add()
            写入行(n,, "起始编号", temp.起始编号)


        End If




        DataGridView3.Columns.Item(0).Width = DataGridView3.Columns.Item(0).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)
        DataGridView3.Columns.Item(1).Width = DataGridView3.Columns.Item(1).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)
        DataGridView3.Columns.Item(2).Width = DataGridView3.Columns.Item(2).GetPreferredWidth(Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells, True)

        'DataGridView3.Columns.Item(0).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        'DataGridView3.Columns.Item(1).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        'DataGridView3.Columns.Item(2).AutoSizeMode = Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells


    End Function


    Public Function 保存当前规则设置(规则 As Object)

        Dim 总行数 As Integer = DataGridView3.RowCount

        If 规则.GetType.Name = "标志位规则" Then

            Dim temp As 标志位规则 = CType(规则, 标志位规则)
            temp.标志位编号序列.Clear()
            For n As Integer = 0 To 总行数 - 1
                If DataGridView3.Rows(n).Cells(1).Value IsNot Nothing Or DataGridView3.Rows(n).Cells(2).Value IsNot Nothing Then
                    temp.add(DataGridView3.Rows(n).Cells(1).Value, DataGridView3.Rows(n).Cells(2).Value)
                End If
            Next

        ElseIf 规则.GetType.Name = "常量规则" Then
            Dim temp As 常量规则 = CType(规则, 常量规则)
            temp.常量值 = DataGridView3.Rows(0).Cells(2).Value

        ElseIf 规则.GetType.Name = "引用规则" Then
            Dim temp As 引用规则 = CType(规则, 引用规则)
            temp.引用列限制位数 = DataGridView3.Rows(0).Cells(2).Value
            temp.补位字符 = DataGridView3.Rows(1).Cells(2).Value

        ElseIf 规则.GetType.Name = "分类编号规则" Then
            Dim temp As 分类编号规则 = CType(规则, 分类编号规则)
            temp.限制位数 = DataGridView3.Rows(0).Cells(2).Value
            temp.起始编号 = DataGridView3.Rows(1).Cells(2).Value


            'For i = 0 To CheckedListBox1.Items.Count - 1
            '    If CheckedListBox1.GetItemChecked(i) = True Then
            '        temp.分类列号序列.Add(i + 1)
            '    End If
            'Next



            'temp.分类列号序列 = DataGridView3.Rows(1).Cells(2).Value




            temp.分类列号序列.Clear()
            Dim str As String = DataGridView3.Rows(2).Cells(2).Value
            Dim tempstr As String = ""
            For Each c As Char In str
                If c <> "," Then
                    tempstr &= c
                Else
                    temp.分类列号序列.Add(tempstr)
                    tempstr = ""
                End If
            Next




        ElseIf 规则.GetType.Name = "整体编号规则" Then
            Dim temp As 整体编号规则 = CType(规则, 整体编号规则)
            temp.限制位数 = DataGridView3.Rows(0).Cells(2).Value
            temp.起始编号 = DataGridView3.Rows(1).Cells(2).Value

        End If




        My.Settings.Save()

    End Function





    Private Sub TreeView1_BeforeSelect(sender As Object, e As Windows.Forms.TreeViewCancelEventArgs) Handles TreeView1.BeforeSelect
        'If TreeView1.SelectedNode IsNot Nothing Then
        '    If TreeView1.SelectedNode.Level = 2 Then
        '          MsgBox(DataGridView3.Rows(0).Cells(2).Value)
        '        保存当前规则设置(当前编号方案.编号规则序列(TreeView1.SelectedNode.Index))
        '    End If

        'End If


    End Sub



    'Private Sub ContextMenuStrip1_Opening(sender As Object, e As ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    'End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        For Each 方案 As 编号方案 In 编号方案列表
            If 冗余行列检查(方案.工作表, True) = True Then
                Exit Sub
            End If
        Next

        Dim n As Integer = 0
        设置为手动计算()
        流水信息.清除信息()
        For Each 方案 As 编号方案 In 编号方案列表
            n += 开始编号(方案)
            '方案.错误管理者.显示错误信息(方案.编号方案名称)
        Next
        流水信息.记录信息("所有方案，合计编号个数 共 " & n & " 个")
        流水信息.显示信息("各编号方案错误信息统计")

        设置为自动计算()

    End Sub



    Public Function 开始编号(方案 As 编号方案) As Integer
        Dim n As Integer = 0
        Dim sheet As Excel.Worksheet = 方案.工作表
        Dim EndCell As Excel.Range = 获取结束单元格(sheet)


        Dim cell_1 As Excel.Range = sheet.UsedRange.Cells(方案.工作表的列标题所在行号 + 1, 1)
        Dim cell_2 As Excel.Range = 获取结束单元格(sheet)
        Dim 数据区 As Excel.Range
        数据区 = MyRange(sheet, cell_1.Row, cell_1.Column, cell_2.Row, cell_2.Column)

        '数据区 = app.Range(cell_1, cell_2)


        Dim 列 As Excel.Range = 插入列(sheet, EndCell.Column + 1, True, "@")
        方案.编号所在的列索引 = 列.Column
        列.Cells(方案.工作表的列标题所在行号, 1).value = "#" & 方案.编号方案名称


        If 方案.行对象筛选序列.Count = 0 Then
            For Each 行 As Excel.Range In 数据区.Rows
                列.Cells(行.Row, 1).value = 方案.申请编号(行)
                n += 1
            Next
        Else
            Dim 特征字符串 As String = ""
            Dim 行特征字符串 As String = ""
            For Each temp As 元组 In 方案.行对象筛选序列
                特征字符串 &= temp.值
            Next
            For Each 行 As Excel.Range In 数据区.Rows
                For Each temp As 元组 In 方案.行对象筛选序列
                    行特征字符串 &= 行.Cells(1, temp.健 + 1).value
                Next
                If 行特征字符串 = 特征字符串 Then
                    列.Cells(行.Row, 1).value = 方案.申请编号(行)
                    n += 1
                End If
                行特征字符串 = ""
            Next

        End If
        列.AutoFit()
        设置内部边框(区域交集(sheet.UsedRange, 列), 2, 2)
        '设置外边框(区域交集(sheet.UsedRange, 列), 2, 2)
        '''''''''''号段统计   号段统计
        '''''''''''号段统计   号段统计

        流水信息.记录信息("方案:" & 方案.编号方案名称 & "  中，共统计 " & 号段统计(方案) & "个号码。" & vbCrLf)

        方案.整体序号 = 0

        Return n
    End Function
    ''' <summary>
    ''' 号段统计
    ''' </summary>
    ''' <param name="方案">编号方案对象</param>
    ''' <returns>统计得到的总数</returns>

    Public Function 号段统计(方案 As 编号方案) As Integer
        Dim 规则号 As Integer = 1
        Dim 总数 As Integer = 0
        For Each 规则 As Object In 方案.编号规则序列
            If 规则.GetType.Name = "分类编号规则" Then
                Dim 当前分类规则 As 分类编号规则 = CType(规则, 分类编号规则)
                Dim 统计表 As Excel.Worksheet = 新建工作表(方案.编号方案名称 & "_号段统计" & 规则号, True)

                规则号 += 1

                Dim column As Integer = 1
                统计表.Cells(1, column).value = "统计序号"
                column += 1

                For Each 列号 As Integer In 当前分类规则.分类列号序列
                    统计表.Cells(1, column).value = 方案.工作表.Cells(方案.工作表的列标题所在行号, 列号).value
                    column += 1
                Next
                统计表.Cells(1, column).value = "起始号码"
                设置单元格格式(统计表.Columns.Item(column), "文本")
                column += 1
                统计表.Cells(1, column).value = "终止号码"
                设置单元格格式(统计表.Columns.Item(column), "文本")
                column += 1




                统计表.Cells(1, column).value = "数量"
                column += 1
                统计表.Cells(1, column).value = "特征码"
                column += 1


                Dim row As Integer = 2
                Dim 统计序号 As Integer = 1
                For Each 段 As System.Collections.ArrayList In 当前分类规则.分类序号记录表 '格式：{行特征代码，当前尾编号，此特征代码的首行索引，此特征代码的尾行索引}
                    column = 1
                    统计表.Cells(row, column).value = 统计序号
                    统计序号 += 1
                    column += 1
                    For Each 列号 As Integer In 当前分类规则.分类列号序列
                        统计表.Cells(row, column).value = 方案.工作表.Cells(段(2), 列号).value '分类列信息
                        column += 1
                    Next

                    统计表.Cells(row, column).value = 方案.工作表.Cells(段(2), 方案.编号所在的列索引).value '号段起始编号
                    column += 1
                    统计表.Cells(row, column).value = 方案.工作表.Cells(段(3), 方案.编号所在的列索引).value '号段结尾编号
                    column += 1
                    统计表.Cells(row, column).value = 段(1) - 当前分类规则.起始编号 + 1 '号段内数量
                    column += 1
                    统计表.Cells(row, column).value = 段(0) '本段特征码
                    column += 1


                    总数 += (段(1) - 当前分类规则.起始编号 + 1)
                    row += 1

                Next
                当前分类规则.分类序号记录表.Clear()
                自动列宽(统计表)
                居中(统计表.UsedRange)
                设置为表头样式(统计表, 1, 1)
            End If

        Next

        Return 总数

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        删除规则()
    End Sub

    Public Sub 删除规则()
        If TreeView1.SelectedNode IsNot Nothing Then
            If TreeView1.SelectedNode.Level = 2 Then
                Dim 方案名 As String = TreeView1.SelectedNode.Parent.Text
                Dim 规则索引 As Integer = TreeView1.SelectedNode.Index
                获取方案(方案名).编号规则序列.RemoveAt(规则索引)
                刷新方案树()
            Else
                MsgBox("请选择要删除的规则节点！")
            End If
        End If

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        If CheckedListBox2.SelectedItem IsNot Nothing Then
            If CheckedListBox2.SelectedItem.ToString = "分类编号规则" Then
                CheckedListBox1.CheckOnClick = True '使其可以实现多选
                创建默认新规则()

            ElseIf CheckedListBox2.SelectedItem.ToString = "引用规则" Then
                创建默认新规则()

            ElseIf CheckedListBox2.SelectedItem.ToString = "标志位规则" Then
                创建默认新规则()


            End If
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        加载列标题(app.Sheets(ComboBox2.Text), NumericUpDown1.Value, DataGridView1)

        For Each 方案 As 编号方案 In 编号方案列表
            If 方案.编号方案名称 = ComboBox1.Text Then
                当前编号方案 = 方案
                选中指定命令(ComboBox2, 当前编号方案.工作表.Name)
                For Each temp As 元组 In 当前编号方案.行对象筛选序列
                    DataGridView1.Rows(temp.健).Cells(1).Value = temp.值
                Next



                刷新方案信息()
            End If
        Next
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 1 Then
            If 编号方案列表.Count < 1 Then
                MsgBox("你还没有添加任何编号方案" & vbCrLf & "请先添加编号方案，然后再开始设置。")
                TabControl1.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub 上移ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 上移ToolStripMenuItem.Click

        Dim index As Integer = TreeView1.SelectedNode.Index
        If index > 0 Then
            If TreeView1.SelectedNode.Level = 2 Then '规则节点
                Dim temp As Object = 当前编号方案.编号规则序列(index)
                当前编号方案.编号规则序列.RemoveAt(index)
                当前编号方案.编号规则序列.Insert(index - 1, temp)
            End If
        End If
        刷新方案树(index - 1)
    End Sub

    Private Sub 下移ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 下移ToolStripMenuItem.Click

        Dim index As Integer = TreeView1.SelectedNode.Index
        If index < 当前编号方案.编号规则序列.Count - 1 Then
            If TreeView1.SelectedNode.Level = 2 Then '规则节点
                Dim temp As Object = 当前编号方案.编号规则序列(index)
                当前编号方案.编号规则序列.RemoveAt(index)
                当前编号方案.编号规则序列.Insert(index + 1, temp)
            End If
        End If
        刷新方案树(index + 1)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim temp As 元组
        当前编号方案.行对象筛选序列.Clear()
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value <> "" Then
                temp = New 元组(i, DataGridView1.Rows(i).Cells(1).Value)
                当前编号方案.行对象筛选序列.Add(temp)
            End If

        Next
    End Sub

    Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click
        ClickTreeView()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        保存方案()
    End Sub
    Public Function 保存方案()
        Dim 编号方案存储表 As Excel.Worksheet = 新建工作表(编号方案保存表名, False, False)
        Dim n As Integer = 1
        For Each 方案 As 编号方案 In 编号方案列表
            编号方案存储表.Cells(n, 1) = 方案.ToString
            n += 1
        Next

        编号方案存储表.Columns(1).ColumnWidth = 100

        ' 设置行高自动调整
        编号方案存储表.Rows.AutoFit()

        ' 设置单元格为自动换行
        编号方案存储表.Cells.WrapText = True


        MsgBox("共成功保存了 " & n - 1 & "个 编号方案规则！")

    End Function
    Public Function 导入方案()
        Dim sheet As Excel.Worksheet
        Dim n As Integer = 0
        Dim m As Integer = 0
        If 是否存在工作表(编号方案保存表名) = True Then
            sheet = app.Worksheets(编号方案保存表名)
            For Each cell As Excel.Range In sheet.UsedRange
                Try
                    编号方案列表.Add(New 编号方案().CreatFromString(cell.Value))
                    n += 1
                Catch ex As Exception
                    m += 1
                End Try
            Next
        Else
            If MsgBox("未发现储存编号规则的工作表" & vbCrLf & "是否尝试从当前单元格中导入储存的编号规则？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

                Dim cell As Excel.Range = app.ActiveCell
                Try
                    编号方案列表.Add(New 编号方案().CreatFromString(cell.Value))
                    n += 1
                Catch ex As Exception
                    m += 1
                End Try
            End If
        End If
            刷新方案树()
        流水信息.记录信息("成功导入 " & n & "个 编号方案规则，" & "失败 " & m & "个")
        流水信息.显示信息()


    End Function
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        导入方案()
    End Sub

    Private Sub 删除编号方案ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除编号方案ToolStripMenuItem.Click
        删除当前编号方案()
    End Sub

    Private Sub 删除规则ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 删除规则ToolStripMenuItem.Click
        删除规则()
    End Sub

    Private Sub Button9_MouseEnter(sender As Object, e As EventArgs) Handles Button9.MouseEnter
        检测是否有可导入的编号规则()
    End Sub

    Private Sub Button9_MouseLeave(sender As Object, e As EventArgs) Handles Button9.MouseLeave
        Button9.ForeColor = Color.Black
        Button9.Text = "导入规则"
    End Sub

    Private Sub DataGridView3_CellEndEdit(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellEndEdit

        If TreeView1.SelectedNode IsNot Nothing Then
            If TreeView1.SelectedNode.Level = 2 Then

                保存当前规则设置(当前编号方案.编号规则序列(TreeView1.SelectedNode.Index))
                刷新方案树()

                ' 获取当前编辑结束的单元格
                Dim currentCell As DataGridViewCell = DataGridView3(e.ColumnIndex, e.RowIndex)

                '' 获取修改后的单元格值
                'Dim newValue As Object = currentCell.Value.ToString()

                '' 现在你可以使用newValue变量进行后续的操作，例如更新数据源或者其他计算
                'Debug.WriteLine("Modified cell value: " & newValue.ToString())

                MsgBox("已修改设置为：" & currentCell.Value.ToString())
            End If

        End If
    End Sub

    Private Sub 刷新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 刷新ToolStripMenuItem.Click, 刷新ToolStripMenuItem1.Click
        刷新方案树()
        If 当前编号方案 IsNot Nothing Then
            加载列标题(当前编号方案.工作表, NumericUpDown1.Value, DataGridView1)
        End If

    End Sub

    Private Sub 多列显示ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 多列显示ToolStripMenuItem.Click
        CheckedListBox2.MultiColumn = Not CheckedListBox2.MultiColumn
        CheckedListBox1.MultiColumn = Not CheckedListBox1.MultiColumn
    End Sub


End Class
