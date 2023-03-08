Public Class 分类包含统计控件

    Public 分类区域 As Excel.Range
    Public 统计值区域 As Excel.Range

    Public bigIntegerValue As System.Numerics.BigInteger


    Class 数对
        Public X As String
        Public Y As String
        Public Sub New(Optional xx As String = "0", Optional yy As String = "0")
            X = xx
            Y = yy
        End Sub

        Public Overrides Function Tostring() As String
            Return "(" & X & "," & Y & ")"
        End Function
    End Class

    Class 单值
        Public 类型 As String = "单值"
        Public 值 As String
        Public 数量 As Integer
        Public Sub New(Optional _值 As String = "", Optional _数量 As Integer = 0)
            值 = _值
            数量 = _数量
        End Sub

        Public Overrides Function Tostring() As String
            Return "(" & 值 & "," & 数量 & ")"
        End Function
    End Class



    Class 号段
        Public 类型 As String = "号段"
        Public 首相 As System.Numerics.BigInteger
        Public 尾项 As System.Numerics.BigInteger

        Public Sub New(_首相 As System.Numerics.BigInteger, _尾项 As System.Numerics.BigInteger)
            首相 = _首相
            尾项 = _尾项
        End Sub


        Public Function 获取项数() As System.Numerics.BigInteger
            Return 尾项 - 首相
        End Function




        Public Overrides Function Tostring() As String
            Return "(" & 首相.ToString & "," & 尾项.ToString & ")"
        End Function
    End Class

    Class 单值记录
        Public 类型 As String = "单值记录"
        Public 分类字符串 As String
        Public 单值序列 As New System.Collections.ArrayList
        Public Sub New(_分类字符串 As String, _值序列 As System.Collections.ArrayList)
            分类字符串 = _分类字符串
            单值序列 = _值序列
        End Sub
        Public Sub New(_分类字符串 As String, _值 As String)
            分类字符串 = _分类字符串
            Dim _单值 As New 单值(_值, 1)
            单值序列.Add(_单值)

        End Sub
        Public Sub New()
            分类字符串 = Nothing
        End Sub

        Public Sub 添加值(值 As String)
            值 = 值.Trim
            Dim 是否存在 As Boolean = False
            For Each ValueClass As 单值 In 单值序列
                If ValueClass.值 = 值 Then
                    ValueClass.数量 += 1
                    是否存在 = True
                    Exit For
                End If
            Next
            If 是否存在 = False Then
                Dim 新单值 As New 单值(值, 1)
                单值序列.Add(新单值)
            End If

        End Sub


        Public Overrides Function ToString() As String
            Dim result As String = ""
            For Each ValueClass As 单值 In 单值序列
                result &= "(" & ValueClass.值 & "," & ValueClass.数量 & "),"
            Next
            Return result
        End Function

        Public Function 表格输出(sheet As Excel.Worksheet, 行号 As Integer, 列号 As Integer)
            Dim 列 As Integer
            Dim num As Integer = 1
            Dim tempRange As Excel.Range
            列 = 列号
            sheet.Cells(行号, 列).value = 分类字符串
            设置外边框(sheet.Cells(行号, 列), 2, 4, RGB(0, 0, 0))
            列 += 1
            For Each value As 单值 In 单值序列
                sheet.Cells(行号, 列).value = value.值
                列 += 1
                sheet.Cells(行号, 列).value = value.数量
                列 += 1
                tempRange = sheet.Cells(行号, 列 - 2).resize(1, 2)
                设置外边框(tempRange, 2, 4, RGB(0, 0, 0))
                设置填充色(tempRange, 获取循环色(num))
                num += 1
            Next

        End Function
    End Class



    Class 号段记录
        Public 类型 As String = "号段记录"
        Public 分类字符串 As String
        Public 号段序列 As New System.Collections.ArrayList
        'Public 错误管理者 As New 错误管理

        Public Sub New(_分类字符串 As String, _号段序列 As System.Collections.ArrayList)
            分类字符串 = _分类字符串
            号段序列 = _号段序列
        End Sub
        Public Sub New(_分类字符串 As String, _号 As System.Numerics.BigInteger)
            分类字符串 = _分类字符串
            Dim _号段 As New 号段(_号, _号)
            号段序列.Add(_号段)

        End Sub
        Public Sub New()
            分类字符串 = Nothing
        End Sub

        Public Function 添加值(号码 As String) As Boolean
            Dim NUM As System.Numerics.BigInteger
            NUM = System.Numerics.BigInteger.Parse(号码)
            添加值(NUM)
        End Function



        Public Function 添加值(号码 As System.Numerics.BigInteger) As Boolean

            Dim 是否存在 As Boolean = False
            Dim 号段0 As 号段
            If 号段序列.Count = 0 Then
                Dim 新号段 As New 号段(号码, 号码)
                号段序列.Add(新号段)
                Return True
            End If
            For i As Integer = 0 To 号段序列.Count - 1
                号段0 = 号段序列(i)
                If 号码 < 号段0.首相 Then
                    If 号码 = 号段0.首相 - 1 Then
                        号段0.首相 = 号码
                        Return True
                    Else
                        Dim 新号段 As New 号段(号码, 号码)
                        号段序列.Insert(i, 新号段)
                        Return True
                    End If






                ElseIf 号段0.首相 <= 号码 And 号段0.尾项 >= 号码 Then

                    流水信息.记录信息("重复号码：" & 号码.ToString)

                    '错误管理者.记录错误信息("重复号码：" & 号码)
                    'MsgBox("重复号码：" & 号码)
                    是否存在 = True
                    Return False





                ElseIf 号码 > 号段0.尾项 Then
                    If 号码 = 号段0.尾项 + 1 Then
                        号段0.尾项 = 号码
                        If i + 1 <= 号段序列.Count - 1 Then '存在后续号段
                            Dim 号段1 As 号段 = 号段序列(i + 1)
                            If 号段0.尾项 + 1 = 号段1.首相 Then
                                号段0.尾项 = 号段1.尾项
                                号段序列.Remove(号段1)
                            End If
                        End If
                        Return True

                    Else

                        If i = 号段序列.Count - 1 Then
                            Dim 新号段 As New 号段(号码, 号码)
                            号段序列.Add(新号段)
                            Return True
                        End If



                    End If



                Else
                    Return False


                End If




            Next


        End Function


        Public Overrides Function ToString() As String
            Dim result As String = ""
            For Each 号段0 As 号段 In 号段序列
                result &= "(" & 号段0.首相.ToString & "," & 号段0.尾项.ToString & "),"
            Next
            Return result
        End Function

        Public Function 表格输出(sheet As Excel.Worksheet, 行号 As Integer, 列号 As Integer)
            Dim 列 As Integer
            Dim num As Integer = 1
            列 = 列号
            sheet.Cells(行号, 列).value = 分类字符串
            设置外边框(sheet.Cells(行号, 列), 2, 4, RGB(0, 0, 0))
            列 += 1
            Dim tempRange As Excel.Range
            For Each 号段0 As 号段 In 号段序列
                sheet.Cells(行号, 列).value = 号段0.首相.ToString
                列 += 1
                sheet.Cells(行号, 列).value = 号段0.尾项.ToString
                列 += 1
                tempRange = sheet.Cells(行号, 列 - 2).resize(1, 2)
                设置外边框(tempRange, 2, 4, RGB(0, 0, 0))
                设置填充色(tempRange, 获取循环色(num))
                num += 1


            Next

        End Function
    End Class



    Private Sub 分类统计值控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"分类区域", "统计值区域"})
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        分类区域 = 区域选择控件1.获取区域(0)
        统计值区域 = 区域选择控件1.获取区域(1)

        If 区域选择控件1.区域校验() = True Then
            分类统计值(分类区域, 统计值区域, NumericUpDown1.Value)
        Else
            区域选择控件1.显示错误()
        End If




    End Sub







    Public Function 分类统计值(分类列区域 As Excel.Range, 统计值列区域 As Excel.Range, Optional 列标题所在行号 As Integer = 1)
        If 分类列区域 Is Nothing Then
            Exit Function
        End If

        If 统计值列区域 Is Nothing Then
            Exit Function
        End If


        If 分类列区域.Worksheet IsNot 统计值列区域.Worksheet Then
            Exit Function
        End If

        If 列标题所在行号 < 0 Then
            列标题所在行号 = 0
        End If


        Dim sheet As Excel.Worksheet = 分类列区域.Worksheet

        Dim 最大行号 As Integer = Math.Min(获取结束行号(分类列区域), 获取结束行号(统计值列区域))

        If 列标题所在行号 >= 最大行号 Then
            Exit Function
        End If



        流水信息.清除信息()
        Dim 分类数据区域 As Excel.Range = 分类列区域.Cells(列标题所在行号 + 1, 1).resize(最大行号 - 列标题所在行号, 1)
        Dim 统计值数据区域 As Excel.Range = 统计值列区域.Cells(列标题所在行号 + 1, 1).resize(最大行号 - 列标题所在行号, 1)

        '设置填充色(分类数据区域, RGB(100, 0, 0))
        '设置填充色(统计值数据区域, RGB(0, 100, 0))

        Dim 是否全为数字 As Boolean = True
        For Each cell As Excel.Range In 统计值数据区域
            If IsNumeric(cell.Value) = False And cell.Value IsNot Nothing Then
                是否全为数字 = False
                Exit For
            End If
            'sheet.Cells(cell.Row, cell.Column + 1).value = IsNumeric(cell.Value)
        Next


        Dim 统计结果序列 As New System.Collections.ArrayList
        Dim 当前分类， 当前值 As String
        Dim 当前号 As System.Numerics.BigInteger

        If 是否全为数字 = True Then

            Dim 是否存在 As Boolean = False
            For Each cel1 As Excel.Range In 分类数据区域
                是否存在 = False
                If cel1.Value Is Nothing Then
                    当前分类 = ""
                Else
                    当前分类 = cel1.Value
                End If

                If sheet.Cells(cel1.Row, 统计值数据区域.Column).value Is Nothing Then

                    流水信息.记录信息("⚠警告：第" & cel1.Row & "行" & 统计值数据区域.Column & "列" & " 为空值")
                    Continue For
                Else
                    当前号 = System.Numerics.BigInteger.Parse(sheet.Cells(cel1.Row, 统计值数据区域.Column).value)
                End If


                For Each 记录 As 号段记录 In 统计结果序列
                    If 记录.分类字符串 = 当前分类 Then
                        记录.添加值(当前号)
                        是否存在 = True
                    End If
                Next
                If 是否存在 = False Then
                    Dim 新记录 As New 号段记录(当前分类, 当前号)
                    统计结果序列.Add(新记录)
                End If

            Next



        Else
            Dim 是否存在 As Boolean = False
            For Each cel1 As Excel.Range In 分类数据区域
                是否存在 = False
                'If cel1.Value Is Nothing Then
                '    当前分类 = ""
                'Else
                '    当前分类 = cel1.Value
                'End If




                'If sheet.Cells(cel1.Row, 统计值数据区域.Column).value Is Nothing Then
                '    当前值 = ""
                'Else
                '    当前值 = sheet.Cells(cel1.Row, 统计值数据区域.Column).value
                'End If

                当前分类 = Module1.单元格的值转字符串(cel1, My.Settings.代表空值的字符串)

                当前值 = Module1.单元格的值转字符串(sheet.Cells(cel1.Row, 统计值数据区域.Column), My.Settings.代表空值的字符串)

                For Each 记录 As 单值记录 In 统计结果序列
                    If 记录.分类字符串 = 当前分类 Then
                        记录.添加值(当前值)
                        是否存在 = True
                    End If
                Next
                If 是否存在 = False Then
                    Dim 新记录 As New 单值记录(当前分类, 当前值)
                    统计结果序列.Add(新记录)
                End If




            Next


        End If


        ''''''''以下进行统计结果输出'''''''
        ''''''''以下进行统计结果输出'''''''
        ''''''''以下进行统计结果输出'''''''
        ''''''''以下进行统计结果输出'''''''
        ''''''''以下进行统计结果输出'''''''
        ''''''''以下进行统计结果输出'''''''

        Dim 列表头列数 As Integer = 0

        If 统计结果序列.Count > 0 Then
            Dim ResultSheet As Excel.Worksheet = 新建工作表("分类统计值结果", True)
            设置单元格格式(ResultSheet.Cells, "文本")
            Dim Row As Integer = 2
            Dim type As String = 统计结果序列(0).类型
            If type = "单值记录" Then
                For Each value As 单值记录 In 统计结果序列
                    value.表格输出(ResultSheet, Row, 1)
                    Row += 1
                    If value.单值序列.Count > 列表头列数 Then
                        列表头列数 = value.单值序列.Count
                    End If

                Next



                ''''''''以下进行列表头输出'''''''



                ResultSheet.Cells(1, 1).value = "#" & 分类列区域.Cells(列标题所在行号, 1).value

                设置为表头样式(ResultSheet, 1, 1)

                For i As Integer = 1 To 列表头列数

                    ResultSheet.Cells(1, 2 * i).value = "#" & 统计值列区域.Cells(列标题所在行号, 1).value
                    ResultSheet.Cells(1, 1 + 2 * i).value = "数量"
                    设置外边框(ResultSheet.Cells(1, 2 * i).resize(1, 2), 2, 4, RGB(0, 0, 0))

                Next













            ElseIf type = "号段记录" Then

                For Each value As 号段记录 In 统计结果序列
                    value.表格输出(ResultSheet, Row, 1)
                    Row += 1
                    If value.号段序列.Count > 列表头列数 Then
                        列表头列数 = value.号段序列.Count
                    End If
                Next

                ''''''''以下进行列表头输出'''''''
                If 列标题所在行号 > 0 Then
                    ResultSheet.Cells(1, 1).value = "#" & 分类列区域.Cells(列标题所在行号, 1).value
                End If

                设置为表头样式(ResultSheet, 1, 1)
                For i As Integer = 1 To 列表头列数
                    ResultSheet.Cells(1, 2 * i).value = "号段首"
                    ResultSheet.Cells(1, 1 + 2 * i).value = "号段尾"
                    设置外边框(ResultSheet.Cells(1, 2 * i).resize(1, 2), 2, 4, RGB(0, 0, 0))

                Next






            End If








            自动列宽(ResultSheet)
            居中(ResultSheet.Cells)
        End If
        流水信息.显示信息()
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("这里还在施工，请绕行！")
    End Sub
End Class
