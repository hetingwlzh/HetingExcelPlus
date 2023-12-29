
Imports System.ComponentModel.Design
Imports System.Data.SqlTypes
Imports System.Net.Http
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports MyExcel插件.编号控件
Imports Newtonsoft.Json

Module Module1
    Public MAXROW As Integer = 1048576
    Public MAXCOLUMN As Integer = 16384
    Public app As Excel.Application = Globals.ThisAddIn.Application
    Public NextBlink As Double
    'The cell that you want to blink
    Public Const BlinkCell As String = "Sheet1!B2"

    Public 考号信息表 As String = "考试信息"
    Public 缓冲表名 As String = "heting插件_缓冲表"
    Public 设置页名 As String = "Heting插件公共设置页"
    Public 设置页标题 As String = "插件公共设置"
    Public 错误信息序列 As New Collections.ArrayList
    Public 当前选区 As Excel.Range '当前所选择的区域
    Public 地址单元格 As Excel.Range '获取的地址要写入的单元格

    Public MySelectForm As New SelectRangeForm '选择操作单元格区域的窗口对象
    'Public MySetControl As New 排名控件 '设置操作参数的控件对象
    Public MyForm1 As New MyForm '我的各种操作窗体的主界面
    Public MyErrorForm As New errorForm  '错误信息显示窗体



    Public 警告色 As Drawing.Color = Drawing.Color.FromArgb(255, 210, 200)
    Public 新创建区域颜色 As Long = RGB(220, 220, 255)

    Public 编号方案列表 As New System.Collections.ArrayList

    Public 特征码连结字符 As String = "/"

    Public 流水信息 As New 信息处理

    Public taskPane As Microsoft.Office.Tools.CustomTaskPane
    Public 是否突出显示相同值 As Boolean = False
    Public 当前选择信息表名 As String = "heting_当前选择信息_1234"
    Public 突出显示相同值条件格式索引号 As Integer


    Public 是否打开聚光灯效果 As Boolean = False

    Public 信息显示控件 As Object
    Public 是否自动统计当前选择信息 As Boolean = False


    Public 剪切板区域 As Excel.Range


    Public 当前表记录 As Excel.Worksheet


    Public 新建工作表标志 As Boolean = False



    Public 二级列表元素分隔符 As String = "█"
    Public 一级列表元素分隔符 As String = "▍"
    Public Class 信息处理
        Public 信息序列 As System.Collections.ArrayList
        Public Sub New()
            信息序列 = New System.Collections.ArrayList
        End Sub
        Public Sub 记录信息(信息 As String)
            信息序列.Add(信息)
        End Sub
        Public Sub 显示信息(Optional 标题 As String = "heting提醒", Optional 是否清除已显示的信息 As Boolean = True)


            MyErrorForm.Text = 标题
            MyErrorForm.Width = 1000
            MyErrorForm.Height = 530
            If 信息序列.Count > 0 Then
                Dim n As Integer = 1
                Dim 总文本 As String = ""
                For Each str As String In 信息序列
                    总文本 &= n & "、" & " " & str & vbCrLf
                    n += 1
                Next
                MyErrorForm.RichTextBox1.Text = 总文本
            Else
                MyErrorForm.RichTextBox1.Text = "成功完成！！"
            End If

            MyErrorForm.Show()
            MyErrorForm.WindowState = FormWindowState.Normal
            If 是否清除已显示的信息 = True Then
                清除信息()
            End If

        End Sub
        Public Sub 清除信息()
            信息序列.Clear()
        End Sub
    End Class






    Public Class 属性与值
        Public 属性 As Collections.ArrayList
        Public 值 As Collections.ArrayList
        Public Sub New()
            属性 = New Collections.ArrayList
            值 = New Collections.ArrayList
        End Sub
        Public Sub 添加属性(属性值)
            属性.Add(属性值)
        End Sub
        Public Sub 添加值(数值)
            值.Add(数值)
        End Sub
        Public Sub 清空属性()
            属性.Clear()
        End Sub
        Public Sub 清空值()
            值.Clear()
        End Sub
        Public Function 属性个数()
            Return 属性.Count
        End Function
        Public Function 值个数()
            Return 值.Count
        End Function
        Public Overrides Function ToString() As String
            Dim result As String = "{"
            For Each t In 属性
                result &= t & ","
            Next
            result &= ":"
            For Each t In 值
                result &= t & ","
            Next
            result &= "}"
            Return result
        End Function
    End Class










    Public Class Common '删除关闭按钮的类
        Private Declare Function GetSystemMenu Lib "User32" (ByVal hwnd As Integer, ByVal bRevert As Long) As Integer
        Private Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
        Private Declare Function DrawMenuBar Lib "User32" (ByVal hwnd As Integer) As Integer
        Private Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Integer) As Integer

        Private Const MF_BYPOSITION = &H400&
        Private Const MF_DISABLED = &H2&


        '禁用窗口右上角的“关闭”按钮
        Public Sub DisableCloseButton(ByVal wnd As Windows.Forms.Form)
            Dim hMenu As Integer, nCount As Integer
            '得到系统Menu
            hMenu = GetSystemMenu(wnd.Handle.ToInt32, 0)
            '得到系统Menu的个数
            nCount = GetMenuItemCount(hMenu)
            '去除系统Menu
            Call RemoveMenu(hMenu, nCount - 1, MF_BYPOSITION Or MF_DISABLED)
            '重画MenuBar
            DrawMenuBar(wnd.Handle.ToInt32)
        End Sub
    End Class







    Public Class 记录
        Public key
        Public 数据列 As System.Collections.ArrayList
        Public Sub New()
            数据列 = New Collections.ArrayList
        End Sub
        Public Sub 添加数据(数据 As 属性与值)
            数据列.Add(数据)
        End Sub
        Public Sub 添加数据(数据 As 表头和数据类)
            数据列.Add(数据)
        End Sub
        Public Function 数据个数()
            Return 数据列.Count
        End Function
        Public Function 均值() As Single
            Try
                Dim sum As Single
                Dim n As Integer = 0
                For Each t In 数据列
                    For Each value In t.值
                        sum += CType(Val(value), Single)
                        n += 1
                    Next
                Next
                Return sum / n
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Sub 写入记录到表(表 As Excel.Worksheet, row_num As Integer)

            Dim column_num As Integer = 1
            表.Cells(row_num, column_num) = key '姓名
            column_num += 1
            If 数据列.Count > 0 Then
                表.Cells(row_num, column_num) = 数据列.Item(0).属性(1) '学科
                column_num += 1
                表.Cells(row_num, column_num) = 均值()
                表.Cells(row_num, column_num).NumberFormatlocal = "0.00" '保留小数位数为2
                column_num += 1
                For Each d In 数据列
                    表.Cells(row_num, column_num) = d.属性.item(0)
                    column_num += 1
                    表.Cells(row_num, column_num) = d.值.item(0)
                    column_num += 1
                Next
            End If
        End Sub
        Public Sub 以表间匹配记录形式写入(表 As Excel.Worksheet, row_num As Integer)
            Dim 数据个数 As Integer = 1
            Dim 偏移量 As Integer = 0
            Dim column_num As Integer = 1
            Dim 属性表行表头数, 属性表列表头数, 数据表行表头数, 数据表列表头数 As Integer
            Dim 属性表头， 数据表头 As 表头


            表.Cells(row_num, 1) = key '姓名

            For Each 数据 As 表头和数据类 In 数据列
                属性表头 = 数据.表头序列(0)
                数据表头 = 数据.表头序列(1)
                属性表行表头数 = 属性表头.获取行表头个数
                属性表列表头数 = 属性表头.获取列表头个数
                数据表行表头数 = 数据表头.获取行表头个数
                数据表列表头数 = 数据表头.获取列表头个数

                If My.Settings.结果页列序号_属性表1级表头行号 > 0 And 属性表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头行号) = 属性表头.ColumnHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头行号 > 0 And 属性表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头行号) = 属性表头.ColumnHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头行号 > 0 And 属性表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头行号) = 属性表头.ColumnHead(2)
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_属性表1级表头列号 > 0 And 属性表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头列号) = 属性表头.RowHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头列号 > 0 And 属性表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头列号) = 属性表头.RowHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头列号 > 0 And 属性表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头列号) = 属性表头.RowHead(2)
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_数据表1级表头行号 > 0 And 数据表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头行号) = 数据表头.ColumnHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表2级表头行号 > 0 And 数据表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头行号) = 数据表头.ColumnHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表3级表头行号 > 0 And 数据表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头行号) = 数据表头.ColumnHead(2)
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_数据表1级表头列号 > 0 And 数据表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头列号) = 数据表头.RowHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表2级表头列号 > 0 And 数据表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头列号) = 数据表头.RowHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表3级表头列号 > 0 And 数据表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头列号) = 数据表头.RowHead(2)
                    偏移量 += 1
                End If
                If 数据.值序列.Count > 0 Then
                    表.Cells(row_num, column_num + 偏移量 + 1) = 数据.值转字符串
                    偏移量 += 1
                End If
                设置填充色(MyRange(表, row_num, column_num + 1, row_num, column_num + 偏移量), 获取循环色((数据个数 Mod 2)))
                column_num += 偏移量
                偏移量 = 0
                数据个数 += 1

            Next


        End Sub

        Public Sub 以查询表头记录形式写入(表 As Excel.Worksheet, row_num As Integer)
            Dim 数据个数 As Integer = 1
            Dim 偏移量 As Integer = 0
            Dim column_num As Integer = 1
            Dim 属性表行表头数, 属性表列表头数, 数据表行表头数, 数据表列表头数 As Integer
            Dim 属性表头 As 表头


            表.Cells(row_num, 1) = key '姓名

            For Each 数据 As 表头和数据类 In 数据列
                属性表头 = 数据.表头序列(0)

                属性表行表头数 = 属性表头.获取行表头个数
                属性表列表头数 = 属性表头.获取列表头个数


                If My.Settings.结果页列序号_属性表1级表头行号 > 0 And 属性表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头行号) = 属性表头.ColumnHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头行号 > 0 And 属性表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头行号) = 属性表头.ColumnHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头行号 > 0 And 属性表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头行号) = 属性表头.ColumnHead(2)
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_属性表1级表头列号 > 0 And 属性表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头列号) = 属性表头.RowHead(0)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头列号 > 0 And 属性表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头列号) = 属性表头.RowHead(1)
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头列号 > 0 And 属性表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头列号) = 属性表头.RowHead(2)
                    偏移量 += 1
                End If




                'If My.Settings.结果页列序号_数据表1级表头行号 > 0 And 数据表列表头数 > 0 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头行号) = 数据表头.ColumnHead(0)
                '    偏移量 += 1
                'End If

                'If My.Settings.结果页列序号_数据表2级表头行号 > 0 And 数据表列表头数 > 1 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头行号) = 数据表头.ColumnHead(1)
                '    偏移量 += 1
                'End If

                'If My.Settings.结果页列序号_数据表3级表头行号 > 0 And 数据表列表头数 > 2 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头行号) = 数据表头.ColumnHead(2)
                '    偏移量 += 1
                'End If




                'If My.Settings.结果页列序号_数据表1级表头列号 > 0 And 数据表行表头数 > 0 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头列号) = 数据表头.RowHead(0)
                '    偏移量 += 1
                'End If

                'If My.Settings.结果页列序号_数据表2级表头列号 > 0 And 数据表行表头数 > 1 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头列号) = 数据表头.RowHead(1)
                '    偏移量 += 1
                'End If

                'If My.Settings.结果页列序号_数据表3级表头列号 > 0 And 数据表行表头数 > 2 Then
                '    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头列号) = 数据表头.RowHead(2)
                '    偏移量 += 1
                'End If
                If 数据.值序列.Count > 0 Then
                    表.Cells(row_num, column_num + 偏移量 + 1) = 数据.值转字符串
                    偏移量 += 1
                End If
                设置填充色(MyRange(表, row_num, column_num + 1, row_num, column_num + 偏移量), 获取循环色((数据个数 Mod 2)))
                column_num += 偏移量
                偏移量 = 0
                数据个数 += 1

            Next


        End Sub


        Public Sub 以表间匹配记录形式写入记录表头(表 As Excel.Worksheet, row_num As Integer)
            Dim 数据个数 As Integer = 1
            Dim 偏移量 As Integer = 0
            Dim column_num As Integer = 1
            Dim 属性表行表头数, 属性表列表头数, 数据表行表头数, 数据表列表头数 As Integer
            Dim 属性表头， 数据表头 As 表头


            表.Cells(row_num, 1) = My.Settings.属性表数据名称  '姓名

            For Each 数据 As 表头和数据类 In 数据列
                属性表头 = 数据.表头序列(0)
                数据表头 = 数据.表头序列(1)
                属性表行表头数 = 属性表头.获取行表头个数
                属性表列表头数 = 属性表头.获取列表头个数
                数据表行表头数 = 数据表头.获取行表头个数
                数据表列表头数 = 数据表头.获取列表头个数

                If My.Settings.结果页列序号_属性表1级表头行号 > 0 And 属性表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头行号) = My.Settings.属性表1级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头行号 > 0 And 属性表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头行号) = My.Settings.属性表2级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头行号 > 0 And 属性表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头行号) = My.Settings.属性表3级表头行名
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_属性表1级表头列号 > 0 And 属性表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头列号) = My.Settings.属性表1级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头列号 > 0 And 属性表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头列号) = My.Settings.属性表2级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头列号 > 0 And 属性表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头列号) = My.Settings.属性表3级表头列名
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_数据表1级表头行号 > 0 And 数据表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头行号) = My.Settings.数据表1级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表2级表头行号 > 0 And 数据表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头行号) = My.Settings.数据表2级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表3级表头行号 > 0 And 数据表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头行号) = My.Settings.数据表3级表头行名
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_数据表1级表头列号 > 0 And 数据表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表1级表头列号) = My.Settings.数据表1级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表2级表头列号 > 0 And 数据表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表2级表头列号) = My.Settings.数据表2级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_数据表3级表头列号 > 0 And 数据表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_数据表3级表头列号) = My.Settings.数据表3级表头列名
                    偏移量 += 1
                End If
                If 数据.值序列.Count > 0 Then
                    表.Cells(row_num, column_num + 偏移量 + 1) = My.Settings.数据表数据名称
                    偏移量 += 1
                End If
                设置填充色(MyRange(表, row_num, column_num + 1, row_num, column_num + 偏移量), RGB(220 + (数据个数 + 1 Mod 2) * 35, 200, 220 + (数据个数 Mod 2) * 35))
                column_num += 偏移量
                偏移量 = 0
                数据个数 += 1

            Next


        End Sub
        Public Sub 以查询表头记录形式写入记录表头(表 As Excel.Worksheet, row_num As Integer)
            Dim 数据个数 As Integer = 1
            Dim 偏移量 As Integer = 0
            Dim column_num As Integer = 1
            Dim 属性表行表头数, 属性表列表头数 As Integer
            Dim 属性表头 As 表头


            表.Cells(row_num, 1) = My.Settings.属性表数据名称  '姓名

            For Each 数据 As 表头和数据类 In 数据列
                属性表头 = 数据.表头序列(0)

                属性表行表头数 = 属性表头.获取行表头个数
                属性表列表头数 = 属性表头.获取列表头个数


                If My.Settings.结果页列序号_属性表1级表头行号 > 0 And 属性表列表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头行号) = My.Settings.属性表1级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头行号 > 0 And 属性表列表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头行号) = My.Settings.属性表2级表头行名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头行号 > 0 And 属性表列表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头行号) = My.Settings.属性表3级表头行名
                    偏移量 += 1
                End If




                If My.Settings.结果页列序号_属性表1级表头列号 > 0 And 属性表行表头数 > 0 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表1级表头列号) = My.Settings.属性表1级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表2级表头列号 > 0 And 属性表行表头数 > 1 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表2级表头列号) = My.Settings.属性表2级表头列名
                    偏移量 += 1
                End If

                If My.Settings.结果页列序号_属性表3级表头列号 > 0 And 属性表行表头数 > 2 Then
                    表.Cells(row_num, column_num + My.Settings.结果页列序号_属性表3级表头列号) = My.Settings.属性表3级表头列名
                    偏移量 += 1
                End If



                If 数据.值序列.Count > 0 Then
                    表.Cells(row_num, column_num + 偏移量 + 1) = My.Settings.数据表数据名称
                    偏移量 += 1
                End If
                设置填充色(MyRange(表, row_num, column_num + 1, row_num, column_num + 偏移量), RGB(220 + (数据个数 + 1 Mod 2) * 35, 200, 220 + (数据个数 Mod 2) * 35))
                column_num += 偏移量
                偏移量 = 0
                数据个数 += 1

            Next


        End Sub




        Public Overrides Function ToString() As String
            Dim result As String = "{" & key & ":"
            For Each t In 数据列
                result &= t.ToString & ","
            Next
            result &= "}"
            Return result
        End Function
    End Class
    Public Class 行列
        Public row As Integer
        Public column As Integer

        Public Sub New(Optional row_ As Integer = 1, Optional column_ As Integer = 1)
            row = row_
            column = column_
            If row < 1 Then
                row = 1
            End If
            If column < 1 Then
                column = 1
            End If
        End Sub
        Public Sub Down(Optional n As Integer = 1)
            row += n
        End Sub

        Public Sub Up(Optional n As Integer = 1)
            row -= n
            If row < 1 Then
                row = 1
            End If
        End Sub

        Public Sub Left(Optional n As Integer = 1)
            column -= n
            If column < 1 Then
                column = 1
            End If
        End Sub

        Public Sub Right(Optional n As Integer = 1)
            column += n
        End Sub
        Public Overrides Function ToString() As String
            Return "(" & row & "," & column & ")"
        End Function


    End Class
    ''' <summary>
    ''' 一个单元格的一个表头可以存放多级表头比如 RowHead=["1班"],ColumnHead=["数学","分数"]，
    ''' RowHead中存放行表头按照表头级别依次组成行表头列表
    ''' ColumnHead中存放列表头按照表头级别依次组成行表头列表
    ''' </summary>
    Public Class 表头
        Public RowHead As System.Collections.ArrayList
        Public ColumnHead As System.Collections.ArrayList
        Public Sub New(Row_Head As System.Collections.ArrayList, Column_Head As System.Collections.ArrayList)
            RowHead = Row_Head
            ColumnHead = Column_Head
        End Sub
        Public Sub New()
            RowHead = New System.Collections.ArrayList
            ColumnHead = New System.Collections.ArrayList
        End Sub
        Public Sub New(Row_Head_String As String, Column_Head_String As String)
            RowHead = New System.Collections.ArrayList
            ColumnHead = New System.Collections.ArrayList
            RowHead.Add(Row_Head_String)
            ColumnHead.Add(Column_Head_String)
        End Sub
        Public Function 添加行表头(Row_Head)
            RowHead.Add(Row_Head)
        End Function
        Public Function 添加列表头(Column_Head)
            ColumnHead.Add(Column_Head)
        End Function
        Public Overrides Function ToString() As String
            Return "[(" & 列表转字符串(RowHead) & "),(" & 列表转字符串(ColumnHead) & ")]"
        End Function
        Public Function 获取行表头个数() As Integer
            Return RowHead.Count
        End Function
        Public Function 获取列表头个数() As Integer
            Return ColumnHead.Count
        End Function

    End Class
    ''' <summary>
    ''' 两个元素的元素对形如(x,y),x,y为字符串类型
    ''' </summary>
    Public Class 数对
        Public x As String
        Public y As String

        Public Sub New(Optional xx As Integer = 0, Optional yy As Integer = 0)
            x = xx
            y = yy
        End Sub
        Public Overrides Function ToString() As String
            Return "(" & x & "," & y & ")"
        End Function
    End Class
    ''' <summary>
    ''' 把列表内的元素用指定的分隔符连接成字符串并返回
    ''' </summary>
    ''' <param name="list">要转换的列表</param>
    ''' <param name="分隔符">元素间的分隔符</param>
    ''' <returns>返回链接后的字符串</returns>
    Public Function 列表转字符串(list As System.Collections.ArrayList, Optional 分隔符 As String = ",")
        Try
            Dim result As String = ""
            For Each t In list
                If t IsNot Nothing Then
                    result &= t.ToString & 分隔符
                End If
            Next
            If result.Substring(result.Length - 分隔符.Length, 分隔符.Length) = 分隔符 Then
                result = result.Substring(0, result.Length - 分隔符.Length)
            End If
            Return result
        Catch ex As Exception
            Return ""
        End Try

    End Function
    ''' <summary>
    ''' 把指定列表元素依次写入指定表格的指定行或列中
    ''' </summary>
    ''' <param name="list">指定的列表</param>
    ''' <param name="sheet">指定的表格</param>
    ''' <param name="RowOrColumn">写成行还是列（RowOrColumn小于等于0 代表行，否则代表列</param>
    ''' <param name="行列号">要写入的行或列号</param>
    ''' <param name="起始序号">指定行列的开始写入位置</param>
    ''' <returns>写入元素的个数</returns>
    Public Function 列表写入表格(list As System.Collections.ArrayList,
                           sheet As Excel.Worksheet,
                           Optional RowOrColumn As Integer = 0,
                           Optional 行列号 As Integer = 1,
                           Optional 起始序号 As Integer = 1
                           ) As Integer

        If RowOrColumn <= 0 Then '写成一行
            For Each t In list
                sheet.Cells(行列号, 起始序号) = t.ToString
                起始序号 += 1
            Next
        Else
            For Each t In list
                sheet.Cells(起始序号, 行列号) = t.ToString
                起始序号 += 1
            Next

        End If

        Return 起始序号
    End Function
    ''' <summary>
    ''' 在指定的区域内查找指定的字符串内容d的单元格，
    ''' 并返回查找结果所有单元格组成的区域
    ''' </summary>
    ''' <param name="range">指定要查找的区域</param>
    ''' <param name="要查找内容">指定要查找的字符串内容</param>
    ''' <returns>以Excel.Range形式返回查找结果单元格的区域</returns>
    Public Function 区域查找(range As Excel.Range, 要查找内容 As String) As Excel.Range
        Dim result As Excel.Range
        If 要查找内容 Is Nothing Then
            要查找内容 = ""
        End If
        If range Is Nothing Then
            range = Nothing
        Else
            For Each cell In range
                If Not cell.value Is Nothing Then
                    If cell.value.ToString = 要查找内容 Then
                        Dim t As New 行列(cell.row, cell.column)
                        If result Is Nothing Then
                            result = cell
                        Else
                            result = app.Union(result, cell)
                        End If
                    End If
                End If
            Next
        End If
        Return result
    End Function

    ''' <summary>
    ''' 由给定的表格内容来获取对应的行列各级表头（也即是行列各级属性），个能有多个匹配结果，所以返回匹配结果的列表
    ''' </summary>
    ''' <param name="要查找的表"></param>
    ''' <param name="要查找内容"></param>
    ''' <param name="一级表头列号"></param>
    ''' <param name="一级表头行号"></param>
    ''' <returns>由表头类对象构成的列表</returns>
    Public Function 由值获取表头(要查找的表 As Excel.Worksheet,
                           要查找内容 As String,
                           Optional 一级表头行号 As Integer = 1，
                           Optional 一级表头列号 As Integer = 1,
                           Optional 二级表头行号 As Integer = -1，
                           Optional 二级表头列号 As Integer = -1,
                           Optional 三级表头行号 As Integer = -1，
                           Optional 三级表头列号 As Integer = -1
                           ) As System.Collections.ArrayList

        If 二级表头列号 = Nothing Then
            二级表头列号 = -1
        End If
        If 二级表头行号 = Nothing Then
            二级表头行号 = -1
        End If


        If 三级表头行号 = Nothing Then
            三级表头行号 = -1
        End If
        If 三级表头行号 = Nothing Then
            三级表头行号 = -1
        End If


        Dim 表头序列 As New System.Collections.ArrayList

        Dim range As Excel.Range = 区域移除列(区域移除行(获取用户区域(要查找的表), 一级表头行号, 二级表头行号, 三级表头行号), 一级表头列号, 二级表头列号, 三级表头列号)
        Dim ResultRange As Excel.Range = 区域查找(range, 要查找内容)
        If ResultRange IsNot Nothing Then
            For Each cell In ResultRange
                Dim 表头值 As New 表头()
                Dim temp As String = Nothing
                表头值.RowHead.Add(要查找的表.Cells(cell.row, 一级表头列号).MergeArea.Cells(1, 1).value)
                表头值.ColumnHead.Add(要查找的表.Cells(一级表头行号, cell.column).MergeArea.Cells(1, 1).value)
                If 二级表头列号 >= 1 Then
                    temp = 要查找的表.Cells(cell.row, 二级表头列号).MergeArea.Cells(1, 1).value
                    If temp <> Nothing Then
                        表头值.RowHead.Add(temp)
                    End If

                End If
                temp = Nothing
                If 二级表头行号 >= 1 Then
                    temp = 要查找的表.Cells(二级表头行号, cell.column).MergeArea.Cells(1, 1).value
                    If temp <> Nothing Then
                        表头值.ColumnHead.Add(temp)
                    End If
                End If
                temp = Nothing
                If 三级表头列号 >= 1 Then
                    temp = 要查找的表.Cells(cell.row, 三级表头列号).MergeArea.Cells(1, 1).value
                    If temp <> Nothing Then
                        表头值.RowHead.Add(temp)
                    End If
                End If
                temp = Nothing
                If 三级表头行号 >= 1 Then
                    temp = 要查找的表.Cells(三级表头行号, cell.column).MergeArea.Cells(1, 1).value
                    If temp <> Nothing Then
                        表头值.ColumnHead.Add(temp)
                    End If
                End If


                If 表头值.RowHead(0) IsNot Nothing And 表头值.ColumnHead(0) IsNot Nothing Then
                    表头序列.Add(表头值)
                End If

            Next
        End If

        Return 表头序列 '由表头类对象构成的列表
    End Function


    Public Function 由表头获取值(要查找的表 As Excel.Worksheet,
                          行表头值 As String,
                          一级列表头值 As String,
                          二级列表头值 As String,
                          Optional 行表头所在列 As Integer = 1,
                          Optional 一级列表头所在行 As Integer = 1,
                          Optional 二级列表头所在行 As Integer = 2
                          ) As String
        Dim MaxRowNum As Integer = 获取用户区域(要查找的表).Rows.Count
        Dim MaxColumnNum As Integer = 获取用户区域(要查找的表).Columns.Count
        Dim range As Excel.Range = 要查找的表.Range(要查找的表.Cells(1, 行表头所在列), 要查找的表.Cells(MaxRowNum, 行表头所在列))

        Dim CellRange As Excel.Range = 区域查找(range, 行表头值)
        Dim 行表头号 As Integer
        If CellRange.Count > 0 Then
            行表头号 = CellRange(1, 1).row
        End If



        range = 要查找的表.Range(要查找的表.Cells(一级列表头所在行， 1), 要查找的表.Cells(一级列表头所在行, MaxColumnNum))
        CellRange = 区域查找(range, 一级列表头值)


        Dim 二级列表头号 As Integer
        If CellRange.Count > 0 Then
            Dim 二级表头开始列 As Integer = CellRange(1, 1).column
            Dim 二级表头结束列 As Integer = 二级表头开始列 + CellRange(1, 1).MergeArea.Columns.Count - 1
            Dim range2 As Excel.Range = 要查找的表.Range(要查找的表.Cells(二级列表头所在行， 二级表头开始列),
                                                    要查找的表.Cells(二级列表头所在行, 二级表头结束列))
            二级列表头号 = 区域查找(range2, 二级列表头值)(1, 1).column
            Return 要查找的表.Cells(行表头号, 二级列表头号).value
        Else
            Return Nothing
        End If




    End Function



    Public Function 由表头获取值test(要查找的表 As Excel.Worksheet,
                          表头值 As 表头,
                          Optional 一级表头行号 As Integer = 1,
                          Optional 一级表头列号 As Integer = 1,
                          Optional 二级表头行号 As Integer = -1，
                          Optional 二级表头列号 As Integer = -1,
                          Optional 三级表头行号 As Integer = -1,
                          Optional 三级表头列号 As Integer = -1
                          ) As Excel.Range
        Dim result As Excel.Range
        Dim 数据区 As Excel.Range = 获取用户区域(要查找的表)
        Dim MaxRowNum As Integer = 数据区.Rows.Count
        Dim MaxColumnNum As Integer = 数据区.Columns.Count
        Dim HeadRowCount As Integer = 表头值.获取行表头个数
        Dim HeadColumnCount As Integer = 表头值.获取列表头个数
        Dim 一级表头行位置, 二级表头行位置, 三级表头行位置, 一级表头列位置, 二级表头列位置, 三级表头列位置 As Excel.Range
        一级表头行位置 = 区域查找(MyRange(要查找的表, 一级表头行号, 1, 一级表头行号, MaxColumnNum), 表头值.ColumnHead(0))
        For Each R1 In 一级表头行位置
            二级表头行位置 = 区域查找(MyRange(要查找的表, 二级表头行号, R1.column, 二级表头行号, R1.column + R1.MergeArea.Columns.Count - 1), 表头值.ColumnHead(1))
            For Each R2 In 二级表头行位置
                三级表头行位置 = 区域查找(MyRange(要查找的表, 三级表头行号, R2.column, 三级表头行号, R2.column + R2.MergeArea.Columns.Count - 1), 表头值.ColumnHead(2))
                For Each R3 In 三级表头行位置





                    一级表头列位置 = 区域查找(MyRange(要查找的表, 1, 一级表头列号, MaxRowNum, 一级表头列号), 表头值.RowHead(0))
                    For Each C1 In 一级表头列位置
                        二级表头列位置 = 区域查找(MyRange(要查找的表, C1.row, 二级表头列号, C1.row + C1.MergeArea.Rows.Count - 1, 二级表头列号), 表头值.RowHead(1))
                        For Each C2 In 二级表头列位置
                            三级表头列位置 = 区域查找(MyRange(要查找的表, C2.row, 三级表头列号, C2.row + C2.MergeArea.Rows.Count - 1, 三级表头列号), 表头值.RowHead(2))
                            For Each C3 In 三级表头列位置
                                If result Is Nothing Then
                                    result = 要查找的表.Cells(C3.row, R3.column)
                                Else
                                    result = app.Union(result, 要查找的表.Cells(C3.row, R3.column))
                                End If

                            Next


                        Next
                    Next






                Next
            Next
        Next




        'Dim range As Excel.Range = 要查找的表.Range(要查找的表.Cells(1, 行表头所在列), 要查找的表.Cells(MaxRowNum, 行表头所在列))


        Return result



    End Function
    ''' <summary>
    ''' 由 表头 类型的对象做参数，在指定的表中查找匹配的单元格，最多支持三级表头，也可更少级别。
    ''' </summary>
    ''' <param name="要查找的表">指定要在其中查找的表的表</param>
    ''' <param name="表头值">要匹配的表头类型的对象</param>
    ''' <param name="一级表头行号">一级表头行所在的行号,默认为第1行，不能小于1</param>
    ''' <param name="一级表头列号">一级表头列所在的列号,默认为第1列，不能小于1</param>
    ''' <param name="二级表头行号">二级表头行所在的行号,默认为-1，小于1表示不存在二级表头行</param>
    ''' <param name="二级表头列号">二级表头行所在的列号,默认为-1，小于1表示不存在二级表头列</param>
    ''' <param name="三级表头行号">三级表头行所在的行号,默认为-1，小于1表示不存在三级表头行</param>
    ''' <param name="三级表头列号">三级表头行所在的列号,默认为-1，小于1表示不存在三级表头列</param>
    ''' <param name="是否涂色返回区域">所返回的区域是否设置背景色以突出显示</param>
    ''' <returns></returns>
    Public Function 由表头获取单元格(要查找的表 As Excel.Worksheet,
                          表头值 As 表头,
                          Optional 一级表头行号 As Integer = 1,
                          Optional 一级表头列号 As Integer = 1,
                          Optional 二级表头行号 As Integer = -1，
                          Optional 二级表头列号 As Integer = -1,
                          Optional 三级表头行号 As Integer = -1,
                          Optional 三级表头列号 As Integer = -1,
                          Optional 是否涂色返回区域 As Boolean = False) As Excel.Range
        Dim result As Excel.Range
        Dim 数据区 As Excel.Range = 获取用户区域(要查找的表)
        Dim MaxRowNum As Integer = 数据区.Rows.Count
        Dim MaxColumnNum As Integer = 数据区.Columns.Count
        Dim HeadRowCount As Integer = 表头值.获取行表头个数
        Dim HeadColumnCount As Integer = 表头值.获取列表头个数
        Dim 一级表头行位置, 二级表头行位置, 三级表头行位置, 一级表头列位置, 二级表头列位置, 三级表头列位置 As Excel.Range
        Dim ColumnRange1, ColumnRange2, ColumnRange3, RowRange1, RowRange2, RowRange3 As Excel.Range

        一级表头行位置 = 区域查找(MyRange(要查找的表, 一级表头行号, 1, 一级表头行号, MaxColumnNum), 表头值.ColumnHead(0))
        If 一级表头行位置 IsNot Nothing Then
            For Each R1 In 一级表头行位置
                If ColumnRange1 Is Nothing Then
                    ColumnRange1 = MyRange(要查找的表, 1, R1.column, MaxRowNum, R1.column + R1.MergeArea.Columns.Count - 1)
                Else
                    ColumnRange1 = app.Union(ColumnRange1, MyRange(要查找的表, 1, R1.column, MaxRowNum, R1.column + R1.MergeArea.Columns.Count - 1))
                End If
            Next
        End If



        If 二级表头行号 > 0 And HeadColumnCount > 1 Then
            二级表头行位置 = 区域查找(MyRange(要查找的表, 二级表头行号, 1, 二级表头行号, MaxColumnNum), 表头值.ColumnHead(1))
            If 二级表头行位置 IsNot Nothing Then
                For Each R2 In 二级表头行位置
                    If ColumnRange2 Is Nothing Then
                        ColumnRange2 = MyRange(要查找的表, 1, R2.column, MaxRowNum, R2.column + R2.MergeArea.Columns.Count - 1)
                    Else
                        ColumnRange2 = app.Union(ColumnRange2, MyRange(要查找的表, 1, R2.column, MaxRowNum, R2.column + R2.MergeArea.Columns.Count - 1))
                    End If
                Next
            End If
        Else
            ColumnRange2 = ColumnRange1
        End If


        If 三级表头行号 > 0 And HeadColumnCount > 2 Then
            三级表头行位置 = 区域查找(MyRange(要查找的表, 三级表头行号, 1, 三级表头行号, MaxColumnNum), 表头值.ColumnHead(2))
            If 三级表头行位置 IsNot Nothing Then
                For Each R3 In 三级表头行位置
                    If ColumnRange3 Is Nothing Then
                        ColumnRange3 = MyRange(要查找的表, 1, R3.column, MaxRowNum, R3.column + R3.MergeArea.Columns.Count - 1)
                    Else
                        ColumnRange3 = app.Union(ColumnRange3, MyRange(要查找的表, 1, R3.column, MaxRowNum, R3.column + R3.MergeArea.Columns.Count - 1))
                    End If
                Next
            End If
        Else
            ColumnRange3 = ColumnRange2

        End If








        一级表头列位置 = 区域查找(MyRange(要查找的表, 1, 一级表头列号, MaxRowNum, 一级表头列号), 表头值.RowHead(0))
        If 一级表头列位置 IsNot Nothing Then
            For Each C1 In 一级表头列位置
                If RowRange1 Is Nothing Then
                    RowRange1 = MyRange(要查找的表, C1.row, 1, C1.row + C1.MergeArea.Rows.Count - 1, MaxColumnNum)
                Else
                    RowRange1 = app.Union(RowRange1, MyRange(要查找的表, C1.row, 1, C1.row + C1.MergeArea.Rows.Count - 1, MaxColumnNum))
                End If
            Next
        End If



        If 二级表头列号 > 0 And HeadRowCount > 1 Then
            二级表头列位置 = 区域查找(MyRange(要查找的表, 1, 二级表头列号, MaxRowNum, 二级表头列号), 表头值.RowHead(1))
            If 二级表头列位置 IsNot Nothing Then
                For Each C2 In 二级表头列位置
                    If RowRange2 Is Nothing Then
                        RowRange2 = MyRange(要查找的表, C2.row, 1, C2.row + C2.MergeArea.Rows.Count - 1, MaxColumnNum)
                    Else
                        RowRange2 = app.Union(RowRange2, MyRange(要查找的表, C2.row, 1, C2.row + C2.MergeArea.Rows.Count - 1, MaxColumnNum))
                    End If
                Next
            End If
        Else
            RowRange2 = RowRange1
        End If


        If 三级表头列号 > 0 And HeadRowCount > 2 Then
            三级表头列位置 = 区域查找(MyRange(要查找的表, 1, 三级表头列号, MaxRowNum, 三级表头列号), 表头值.RowHead(2))
            If 三级表头列位置 IsNot Nothing Then
                For Each C3 In 三级表头列位置
                    If RowRange3 Is Nothing Then
                        RowRange3 = MyRange(要查找的表, C3.row, 1, C3.row + C3.MergeArea.Rows.Count - 1, MaxColumnNum)
                    Else
                        RowRange3 = app.Union(RowRange3, MyRange(要查找的表, C3.row, 1, C3.row + C3.MergeArea.Rows.Count - 1, MaxColumnNum))
                    End If
                Next
            End If
        Else
            RowRange3 = RowRange2
        End If



        Dim ColumnR, RowR, ResultRange As Excel.Range


        If ColumnRange1 IsNot Nothing Then
            ColumnR = ColumnRange1
            If ColumnRange2 IsNot Nothing Then
                ColumnR = app.Intersect(ColumnR, ColumnRange2)
                If ColumnR Is Nothing Then
                    Return Nothing
                Else
                    If ColumnRange3 IsNot Nothing Then
                        ColumnR = app.Intersect(ColumnR, ColumnRange3)
                    Else
                        Return Nothing
                    End If
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If


        If RowRange1 IsNot Nothing Then
            RowR = RowRange1
            If RowRange2 IsNot Nothing Then
                RowR = app.Intersect(RowR, RowRange2)
                If RowR Is Nothing Then
                    Return Nothing
                Else
                    If RowRange3 IsNot Nothing Then
                        RowR = app.Intersect(RowR, RowRange3)
                    Else
                        Return Nothing
                    End If
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If



        If ColumnR IsNot Nothing And RowR IsNot Nothing Then
            ResultRange = app.Intersect(ColumnR, RowR)
        Else
            Return Nothing
        End If

        If 是否涂色返回区域 = True Then
            设置填充色(ResultRange, RGB(255, 200, 200))
        End If

        Return ResultRange



    End Function
    Public Function MyRange(sheet As Excel.Worksheet, x1 As Integer, y1 As Integer, Optional x2 As Integer = 0, Optional y2 As Integer = 0) As Excel.Range
        If x1 > 0 And y1 > 0 And x2 > 0 And y2 > 0 Then
            Return sheet.Range(sheet.Cells(x1, y1), sheet.Cells(x2, y2))
        ElseIf x2 = 0 And y2 = 0 Then
            Return sheet.Cells(x1, y1)
        Else
            错误信息序列.Add("函数MyRange参数出错！")
            Return Nothing
        End If



    End Function
    Public Function 通过属性匹配两表格数据(属性表 As Excel.Worksheet,
                                数据表 As Excel.Worksheet，
                                匹配元素 As String) As 记录
        Dim record As New 记录
        record.key = 匹配元素

        Dim result As New System.Collections.ArrayList

        Dim NumOfRow, NumOfColumn As Integer
        Dim 属性序列 As System.Collections.ArrayList = 由值获取表头(属性表,
                                                          匹配元素,
                                                          My.Settings.属性表1级表头行号，
                                                          My.Settings.属性表1级表头列号，
                                                          My.Settings.属性表2级表头行号，
                                                          My.Settings.属性表2级表头列号，
                                                          My.Settings.属性表3级表头行号，
                                                          My.Settings.属性表3级表头列号
                                                          )
        If 属性序列.Count > 0 Then
            Dim TempString As String = Nothing
            For Each 属性表表头值 As 表头 In 属性序列
                Dim 数据表表头值 As New 表头
                NumOfRow = 属性表表头值.获取行表头个数
                NumOfColumn = 属性表表头值.获取列表头个数

                If My.Settings.数据表1级表头行名 = Nothing Then
                    If NumOfColumn >= 1 Then
                        数据表表头值.添加列表头(属性表表头值.ColumnHead(0))
                    End If
                Else
                    TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表1级表头行名)
                    If TempString IsNot Nothing Then
                        数据表表头值.添加列表头(TempString)
                    End If
                End If

                If My.Settings.数据表2级表头行名 = Nothing Then
                    If NumOfColumn >= 2 Then
                        数据表表头值.添加列表头(属性表表头值.ColumnHead(1))
                    End If
                Else
                    If 数据表表头值.获取列表头个数 = 1 Then
                        TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表2级表头行名)
                        If TempString IsNot Nothing Then
                            数据表表头值.添加列表头(TempString)
                        End If
                    End If
                End If

                If My.Settings.数据表3级表头行名 = Nothing Then
                    If NumOfColumn >= 3 Then
                        数据表表头值.添加列表头(属性表表头值.ColumnHead(2))
                    End If
                Else
                    If 数据表表头值.获取列表头个数 = 2 Then
                        TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表3级表头行名)
                        If TempString IsNot Nothing Then
                            数据表表头值.添加列表头(TempString)
                        End If
                    End If
                End If




                If My.Settings.数据表1级表头列名 = Nothing Then
                    If NumOfRow >= 1 Then
                        数据表表头值.添加行表头(属性表表头值.RowHead(0))
                    End If
                Else
                    TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表1级表头列名)
                    If TempString IsNot Nothing Then
                        数据表表头值.添加行表头(TempString)
                    End If
                End If


                If My.Settings.数据表2级表头列名 = Nothing Then
                    If NumOfRow >= 2 Then
                        数据表表头值.添加行表头(属性表表头值.RowHead(1))
                    End If
                Else
                    If 数据表表头值.获取行表头个数 = 1 Then
                        TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表2级表头列名)
                        If TempString IsNot Nothing Then
                            数据表表头值.添加行表头(TempString)
                        End If
                    End If
                End If


                If My.Settings.数据表3级表头列名 = Nothing Then
                    If NumOfRow >= 3 Then
                        数据表表头值.添加行表头(属性表表头值.RowHead(2))
                    End If
                Else
                    If 数据表表头值.获取行表头个数 = 2 Then
                        TempString = 数据表行列表头值代码转换(属性表表头值, My.Settings.数据表3级表头列名)
                        If TempString IsNot Nothing Then
                            数据表表头值.添加行表头(TempString)
                        End If
                    End If
                End If


                Dim temp, value
                'temp = 由表头获取值(数据表, 属性.RowHead(0), 属性.ColumnHead(0)， 二级列标值, 数据表一级表头列号, 数据表一级表头行号, 数据表二级表头行号)
                temp = 由表头获取单元格(数据表,
                                数据表表头值,
                                My.Settings.数据表1级表头行号,
                                My.Settings.数据表1级表头列号,
                                My.Settings.数据表2级表头行号,
                                My.Settings.数据表2级表头列号,
                                My.Settings.数据表3级表头行号,
                                My.Settings.数据表3级表头列号,
                                True)
                If temp Is Nothing Then
                    value = Nothing
                Else
                    value = 列表转字符串(获取区域值序列(temp), My.Settings.字符串分隔符)
                End If
                'Dim data As New 属性与值
                Dim 表头和数据 As New 表头和数据类(属性表表头值, 数据表表头值, value)

                'data.添加属性(属性.RowHead(0))
                'data.添加属性(属性表表头值.ToString & Chr(10) & 数据表表头值.ToString)
                'data.添加属性(属性表表头值.ColumnHead(0))
                'data.添加值(value)


                record.添加数据(表头和数据)


            Next
        End If

        Return record

    End Function
    Public Class 表头和数据类
        Public 表头序列 As Collections.ArrayList
        Public 值序列 As Collections.ArrayList
        Public Sub New()
            表头序列 = New Collections.ArrayList
            值序列 = New Collections.ArrayList
        End Sub
        Public Sub New(表头1 As 表头)
            Me.New()
            表头序列.Add(表头1)
        End Sub
        Public Sub New(表头1 As 表头, 表头2 As 表头)
            Me.New()
            表头序列.Add(表头1)
            表头序列.Add(表头2)
        End Sub
        Public Sub New(表头1 As 表头, 表头2 As 表头, 值 As String)
            Me.New()
            表头序列.Add(表头1)
            表头序列.Add(表头2)
            值序列.Add(值)
        End Sub
        Public Sub New(表头1 As 表头, 表头2 As 表头, 值1 As String, 值2 As String)
            Me.New()
            表头序列.Add(表头1)
            表头序列.Add(表头2)
            值序列.Add(值1)
            值序列.Add(值2)
        End Sub
        Public Sub New(表头 As Collections.ArrayList, 值 As Collections.ArrayList)
            Me.New()
            表头序列 = 表头
            值序列 = 值
        End Sub
        Public Sub Add(表头 As 表头, 值 As String)
            表头序列.Add(表头)
            值序列.Add(值)
        End Sub
        Public Sub Add(表头 As 表头)
            表头序列.Add(表头)
        End Sub
        Public Sub Add(值 As String)
            值序列.Add(值)
        End Sub
        Public Sub 清理(值 As String)
            表头序列.Clear()
            值序列.Clear()
        End Sub
        Public Function 值转字符串(Optional 分隔符 As String = ";") As String
            Return 列表转字符串(值序列, 分隔符)
        End Function
        Public Function 表头转字符串() As String
            Dim result As String = ""
            For Each head As 表头 In 表头序列
                result &= head.ToString
            Next
            Return result
        End Function
    End Class

    Public Function 数据表行列表头值代码转换(属性表表头值 As 表头， 数据表表头值字符串 As String) As String
        Dim NumOfRow, NumOfColumn As Integer
        NumOfRow = 属性表表头值.获取行表头个数
        NumOfColumn = 属性表表头值.获取列表头个数
        If 数据表表头值字符串 = My.Settings.属性表1级表头行名 Then
            If NumOfColumn > 0 Then
                Return 属性表表头值.ColumnHead(0)
            Else
                Return Nothing
            End If
        ElseIf 数据表表头值字符串 = My.Settings.属性表2级表头行名 Then
            If NumOfColumn > 1 Then
                Return 属性表表头值.ColumnHead(1)
            Else
                Return Nothing
            End If
        ElseIf 数据表表头值字符串 = My.Settings.属性表3级表头行名 Then
            If NumOfColumn > 2 Then
                Return 属性表表头值.ColumnHead(2)
            Else
                Return Nothing
            End If
        ElseIf 数据表表头值字符串 = My.Settings.属性表1级表头列名 Then
            If NumOfRow > 0 Then
                Return 属性表表头值.RowHead(0)
            Else
                Return Nothing
            End If
        ElseIf 数据表表头值字符串 = My.Settings.属性表2级表头列名 Then
            If NumOfRow > 1 Then
                Return 属性表表头值.RowHead(1)
            Else
                Return Nothing
            End If
        ElseIf 数据表表头值字符串 = My.Settings.属性表3级表头列名 Then
            If NumOfRow > 2 Then
                Return 属性表表头值.RowHead(2)
            Else
                Return Nothing
            End If
        Else
            Return 数据表表头值字符串
        End If


    End Function
    Public Function 获取区域值序列(range As Excel.Range) As Collections.ArrayList
        Dim result As New Collections.ArrayList
        If range IsNot Nothing Then
            For Each cell As Excel.Range In range
                result.Add(cell.Value)
            Next
        End If
        Return result
    End Function
    Public Function 表间批量匹配数据(属性表 As Excel.Worksheet,
                             数据表 As Excel.Worksheet,
                             二级列标值 As String,
                             Optional 匹配范围 As Excel.Range = Nothing,
                             Optional 属性表表头行号 As Integer = 1,
                             Optional 属性表表头列号 As Integer = 1,
                             Optional 数据表一级表头行号 As Integer = 1,
                             Optional 数据表一级表头列号 As Integer = 1,
                             Optional 数据表二级表头行号 As Integer = 2
                             ) As Integer

        If 匹配范围 Is Nothing Then
            匹配范围 = app.Selection
        End If

        Dim Result_Sheet As Excel.Worksheet
        If 是否存在工作表("匹配结果") Then
            Result_Sheet = app.Sheets("匹配结果")
        Else
            Result_Sheet = 新建工作表("匹配结果")
        End If

        Result_Sheet.Activate()

        Dim 元素序列 As System.Collections.ArrayList = 枚举选区(匹配范围)
        Dim row_num As Integer = 1
        Dim 一条记录 As 记录

        Result_Sheet.Cells(row_num, 1) = "姓名"
        Result_Sheet.Cells(row_num, 2) = "学科"
        Result_Sheet.Cells(row_num, 3) = "总均分"

        Result_Sheet.Cells(row_num, 4) = "班级"
        Result_Sheet.Cells(row_num, 5) = "均分"

        Result_Sheet.Cells(row_num, 6) = "班级"
        Result_Sheet.Cells(row_num, 7) = "均分"

        Result_Sheet.Cells(row_num, 8) = "班级"
        Result_Sheet.Cells(row_num, 9) = "均分"

        Result_Sheet.Cells(row_num, 10) = "班级"
        Result_Sheet.Cells(row_num, 11) = "均分"
        row_num += 1
        For Each t In 元素序列
            一条记录 = 通过属性匹配两表格数据(属性表, 数据表, t)
            一条记录.以表间匹配记录形式写入(Result_Sheet, row_num)
            row_num += 1
        Next
        Return row_num - 2
    End Function

    Public Function 表间批量匹配数据2(属性表 As Excel.Worksheet, 数据表 As Excel.Worksheet, 元素序列 As System.Collections.ArrayList) As Integer

        'If 匹配范围 Is Nothing Then
        '    匹配范围 = app.Selection
        'End If

        Dim Result_Sheet As Excel.Worksheet
        If 是否存在工作表("匹配结果") Then
            Result_Sheet = app.Sheets("匹配结果")
            Result_Sheet.Cells.Delete()
        Else
            Result_Sheet = 新建工作表("匹配结果")
        End If

        Result_Sheet.Activate()

        'Dim 元素序列 As System.Collections.ArrayList = 枚举区域值(匹配范围)
        Dim row_num As Integer = 1
        Dim 一条记录 As 记录
        Result_Sheet.Cells(2, 3).Select()
        app.ActiveWindow.FreezePanes = True

        row_num += 1
        Dim 数据个数 As Integer = 0
        For Each 元素值 In 元素序列
            一条记录 = 通过属性匹配两表格数据(属性表, 数据表, 元素值)
            If 一条记录.数据个数() > 数据个数 Then
                一条记录.以表间匹配记录形式写入记录表头(Result_Sheet, 1)
                数据个数 = 一条记录.数据个数()
            End If
            一条记录.以表间匹配记录形式写入(Result_Sheet, row_num)
            row_num += 1
        Next
        设置为表头样式(Result_Sheet, 1)
        自动列宽(Result_Sheet)
        居中(Result_Sheet.UsedRange)
        设置内部边框(Result_Sheet.UsedRange, 2, 2, RGB(0, 0, 0))
        设置外边框(Result_Sheet.UsedRange, 2, 4, RGB(0, 0, 0))
        Return row_num - 2
    End Function

    Public Sub StartBlinking()
        With app
            .Goto(.Range("A1"), 1)
            'If the color is red, change the color and text to white
            If .Range(BlinkCell).Interior.ColorIndex = 3 Then
                .Range(BlinkCell).Interior.ColorIndex = 0
                .Range(BlinkCell).Value = "White"
                'If the color is white, change the color and text to red
            Else
                .Range(BlinkCell).Interior.ColorIndex = 3
                .Range(BlinkCell).Value = "Red"
            End If
            'Wait one second before changing the color again

            .Application.OnTime(DateValue(Now + TimeValue("00:00:01")), "StartBlinking")
        End With

    End Sub
    Public Function 插入公式(公式文本 As String, Optional Cell As Excel.Range = Nothing, Optional 是否为数组公式 As Boolean = False)
        Try
            If Cell Is Nothing Then
                Cell = app.ActiveCell
            End If
            If 是否为数组公式 = False Then
                Cell.Formula = 公式文本
            Else
                Cell.FormulaArray = 公式文本

            End If

            Cell.NumberFormatLocal = "G/通用格式" '设置F1单元格为常规格式

        Catch ex As Exception
            Clipboard.SetText(公式文本)
            MsgBox("插入公式出错了,可能原因是公式太长了（有点搞笑,不过网上查阅确实这样）但是却可以手动键入公式，且能正常运行，公式如下：" &
                   vbCrLf & 公式文本 & vbCrLf & vbCrLf & "公式已复制到剪切板！")
        End Try


    End Function

    Public Function 拖拽填充(填充区域 As Excel.Range, Optional 源区域 As Excel.Range = Nothing) As Integer
        If 填充区域 Is Nothing Then
            'MsgBox("拖拽自动填充时，填充区域不能为空！")
            Return 0
        End If
        If 填充区域.Address.Contains(",") Then
            'MsgBox("请选择连续区域！")
            Return 0
        End If


        If 填充区域.Count = 1 Then
            'MsgBox("拖拽自动填充区域单元格个数不能为1！")
            Return 0
        End If

        If 源区域 Is Nothing Then
            源区域 = 填充区域.Cells(1, 1)
        End If


        Dim RowRange, ColumnRange As Excel.Range
        RowRange = 填充区域.Cells(1, 1).resize(源区域.Rows.Count, 填充区域.Columns.Count) '填充区域.Rows.Item(1)
        ColumnRange = 填充区域.Columns.Item(1)

        If 源区域.Count <> RowRange.Count Then
            源区域.AutoFill(Destination:=RowRange, Type:=Excel.XlAutoFillType.xlFillDefault)
        End If

        If RowRange.Count <> 填充区域.Count Then
            RowRange.AutoFill(Destination:=填充区域, Type:=Excel.XlAutoFillType.xlFillDefault)
        End If




        'If 填充区域.Rows.Count > 1 And 填充区域.Columns.Count > 1 Then '填充区域只有一行

        '    'RowRange.FillDown()
        '    '填充区域.FillRight()



        '    '填充区域.Cells(1, 1).AutoFill(Destination:=ColumnRange, Type:=Excel.XlAutoFillType.xlFillDefault)

        '    'RowRange.AutoFill(Destination:=填充区域, Type:=Excel.XlAutoFillType.xlFillDefault)

        'ElseIf 填充区域.Rows.Count = 1 And 填充区域.Columns.Count > 1 Then '填充区域只有一行
        '    '填充区域.Cells(1, 1).AutoFill(Destination:=填充区域, Type:=Excel.XlAutoFillType.xlFillDefault)
        '    '填充区域.FillRight()
        '    填充区域.Cells(1, 1).AutoFill(Destination:=RowRange, Type:=Excel.XlAutoFillType.xlFillDefault)

        'ElseIf 填充区域.Rows.Count > 1 And 填充区域.Columns.Count = 1 Then '填充区域只有一列
        '    '填充区域.Cells(1, 1).AutoFill(Destination:=填充区域, Type:=Excel.XlAutoFillType.xlFillDefault)
        '    '填充区域.FillDown()
        '    填充区域.Cells(1, 1).AutoFill(Destination:=ColumnRange, Type:=Excel.XlAutoFillType.xlFillDefault)

        'ElseIf 填充区域.Rows.Count = 1 And 填充区域.Columns.Count = 1 Then '填充区域只有一个单元格

        'End If
        Return 填充区域.Count



    End Function


    Public Function 获取区域地址(sheet As Excel.Worksheet, 起始行 As Integer, 起始列 As Integer, 结束行 As Integer, 结束列 As Integer) As String
        Dim range As Excel.Range = MyRange(sheet, 起始行, 起始列, 结束行, 结束列)
        Return range.Address
    End Function

    Public Sub 显示插件()
        With app.Worksheets("sheet1")
            .Rows(1).Font.Bold = True
            .Range("a1:d1").Value =
            {"Name", "Full Name", "Title", "Installed"}
            For i = 1 To app.AddIns.Count
                .Cells(i + 1, 1) = app.AddIns(i).Name
                .Cells(i + 1, 2) = app.AddIns(i).FullName
                .Cells(i + 1, 3) = app.AddIns(i).Title
                .Cells(i + 1, 4) = app.AddIns(i).Installed
            Next
            .Range("a1").CurrentRegion.Columns.AutoFit
        End With
    End Sub



    Sub DisplayAddIns()
        app.Worksheets("Sheet1").Activate
        Dim rw = 1
        For Each ad In app.AddIns
            app.Worksheets("Sheet1").Cells(rw, 1) = ad.Name
            app.Worksheets("Sheet1").Cells(rw, 2) = ad.Installed
            rw = rw + 1
        Next
    End Sub

    ''' <summary>
    ''' 创建一个新的名为name的工作表
    ''' </summary>
    ''' <param name="name">表的名字,可以省略</param>
    ''' <returns>返回新建的工作表</returns>
    Function 新建工作表(name As String,
                   Optional 是否自动重命名 As Boolean = False,
                   Optional 是否隐藏 As Boolean = False,
                   Optional After As Object = Nothing,
                   Optional Before As Object = Nothing) As Excel.Worksheet
        Dim Num As Integer = 0
        Dim NewName As String = name
        Dim sheet As Excel.Worksheet
        name = 格式化工作表名(name)
        If name = "" Then
            name = "未命名"
        End If
        If 是否存在工作表(name) Then
            If 是否自动重命名 = True Then
                Do
                    Num += 1
                    NewName = name & "(" & Num & ")"
                Loop Until 是否存在工作表(NewName) = False


                If After IsNot Nothing Then
                    sheet = app.Sheets.Add(After:=After)
                ElseIf Before IsNot Nothing Then
                    sheet = app.Sheets.Add(Before:=Before)
                Else
                    sheet = app.Sheets.Add()
                End If

                sheet.Name = NewName
                Return sheet
            Else
                sheet = app.Sheets(name)
                If MsgBox("已存在名为 """ & name & """ 的工作表！" & Chr(10) & "是否清除其内容？", MsgBoxStyle.YesNo, "是否清除") = MsgBoxResult.Yes Then
                    sheet.Cells.Delete()
                End If
                Return sheet
            End If


        Else
            新建工作表标志 = True
            If After IsNot Nothing Then
                sheet = app.Sheets.Add(After:=After)
            ElseIf Before IsNot Nothing Then
                sheet = app.Sheets.Add(Before:=Before)
            Else
                sheet = app.Sheets.Add()
            End If
            sheet.Name = name

            If 是否隐藏 = True Then
                新建工作表标志 = True
                sheet.Visible = False
            End If
            Return sheet
        End If

    End Function


    Function 是否存在工作表(name As String) As Boolean
        Dim sheet As Excel.Worksheet
        For Each sheet In app.Sheets
            If sheet.Name = name Then
                Return True
            End If

        Next

        Return False
    End Function







    ''' <summary>
    ''' 创建一个新的名为name的工作簿
    ''' </summary>
    ''' <returns>返回新建的工作簿</returns>
    Function 新建工作簿() As Excel.Workbook
        Dim workBook As Excel.Workbook = app.Workbooks.Add()
        Return workBook
    End Function

    ''' <summary>
    ''' 当前所选区域单元个数
    ''' </summary>
    ''' <returns>返回当前所选区域单元个数</returns>

    Function 所选区域单元格个数() As Long
        Return app.Selection.count
    End Function

    ''' <summary>
    ''' 当前所选区域行数
    ''' </summary>
    ''' <returns>返回当前所选区域行数</returns>

    Function 所选区域行数() As Long
        Return app.Selection.rows.count
    End Function




    ''' <summary>
    ''' 当前所选区域列数
    ''' </summary>
    ''' <returns>返回当前所选区域列数</returns>

    Function 所选区域列数() As Long
        Return app.Selection.columns.count
    End Function

    ''' <summary>
    ''' 包含所有数据区域的最小方形区域
    ''' </summary>
    ''' <param name="sheet">指定的工作表</param>
    ''' <returns>返回包含所有数据区域的最小方形区域 Range类型</returns>
    Function 获取用户区域(Optional sheet As Excel.Worksheet = Nothing) As Excel.Range
        If sheet Is Nothing Then
            sheet = app.ActiveSheet
        End If
        Return sheet.UsedRange
    End Function

    ''' <summary>
    '''  获取指定表的第一行第一列的单元格至Sheet.UsedRange的最后一个单元格的连续区域
    ''' </summary>
    ''' <param name="sheet"></param>
    ''' <returns>返回 从 Sheet.Cell(1,1)开始的，包含所有数据区域的最小方形区域</returns>
    Function 获取工作区域(Optional sheet As Excel.Worksheet = Nothing) As Excel.Range
        If sheet Is Nothing Then
            sheet = app.ActiveSheet
        End If
        Return sheet.Range(sheet.Cells(1, 1), 获取结束单元格(sheet))
    End Function


    ''' <summary>
    ''' 已用区域的最左上角单元格
    ''' </summary>
    ''' <param name="sheet">指定的工作表</param>
    ''' <returns>返回已用区域的最左上角单元格 Range类型</returns>
    Function 获取用户区开始单元格(Optional sheet As Excel.Worksheet = Nothing) As Excel.Range

        Return 获取用户区域(sheet).Cells(1, 1)
    End Function


    Public Sub 设置为手动计算()
        app.Calculation = Excel.XlCalculation.xlCalculationManual '改为手动计算，避免页面刷新卡顿
    End Sub
    Public Sub 设置为自动计算()
        app.Calculation = Excel.XlCalculation.xlCalculationAutomatic  '改回自动计算，避免页面刷新卡顿
    End Sub
    Public Sub 开始计算()
        app.Calculate() '开始计算
    End Sub


    Function 枚举选区(Optional range As Excel.Range = Nothing,
                  Optional 是否先行后列 As Boolean = True,
                  Optional 是否剔除重复值 As Boolean = True,
                  Optional 是否忽略空值 As Boolean = True,
                  Optional 是否忽首尾空白字符 As Boolean = False
                  ) As System.Collections.ArrayList
        If range Is Nothing Then
            range = app.Selection
        End If

        Dim result As New System.Collections.ArrayList
        Dim TempStr As String
        If 是否忽首尾空白字符 = True Then
            If 是否先行后列 = True And 是否剔除重复值 = True Then
                For Each cell As Excel.Range In range
                    TempStr = 单元格的值转字符串(cell, My.Settings.代表空值的字符串).Trim
                    If Not result.Contains(TempStr) Then
                        result.Add(TempStr)
                    End If
                Next

            ElseIf 是否先行后列 = True And 是否剔除重复值 = False Then
                For Each cell As Excel.Range In range
                    TempStr = 单元格的值转字符串(cell, My.Settings.代表空值的字符串).Trim
                    result.Add(TempStr)
                Next
            ElseIf 是否先行后列 = False And 是否剔除重复值 = True Then
                For Each column As Excel.Range In range.Columns
                    For Each cell As Excel.Range In column.Cells
                        TempStr = 单元格的值转字符串(cell, My.Settings.代表空值的字符串).Trim
                        If Not result.Contains(TempStr) Then
                            result.Add(TempStr)
                        End If
                    Next
                Next

            ElseIf 是否先行后列 = False And 是否剔除重复值 = False Then
                For Each column As Excel.Range In range.Columns
                    For Each cell As Excel.Range In column.Cells
                        TempStr = 单元格的值转字符串(cell, My.Settings.代表空值的字符串).Trim
                        result.Add(TempStr)
                    Next
                Next

            End If

        Else

            If 是否先行后列 = True And 是否剔除重复值 = True Then
                For Each cell As Excel.Range In range
                    If Not result.Contains(cell.Value) Then
                        result.Add(cell.Value)
                    End If
                Next

            ElseIf 是否先行后列 = True And 是否剔除重复值 = False Then
                For Each cell As Excel.Range In range
                    result.Add(cell.Value)
                Next
            ElseIf 是否先行后列 = False And 是否剔除重复值 = True Then
                For Each column As Excel.Range In range.Columns
                    For Each cell As Excel.Range In column.Cells
                        If Not result.Contains(cell.Value) Then
                            result.Add(cell.Value)
                        End If
                    Next
                Next

            ElseIf 是否先行后列 = False And 是否剔除重复值 = False Then

                For Each column As Excel.Range In range.Columns
                    For Each cell As Excel.Range In column.Cells
                        result.Add(cell.Value)
                    Next
                Next

            End If

        End If



        If 是否忽略空值 = True Then
            Do While result.Contains(Nothing)
                result.Remove(Nothing)
            Loop

            Do While result.Contains("")
                result.Remove("")
            Loop
        End If

        Return result


    End Function

    Public Function 单元格的值转字符串(cell As Excel.Range, Optional 空值时的符串 As String = "") As String
        If cell.Value Is Nothing Then
            Return 空值时的符串
        Else
            Return cell.Value.ToString
        End If
    End Function



    ''' <summary>
    ''' 枚举指定序列中互异的元素
    ''' </summary>
    ''' <param name="序列">要枚举的元素所在序列</param>
    ''' <param name="是否忽略空值">是否忽略空值</param>
    ''' <returns>返回互异元素的新序列</returns>
    Function 枚举序列(序列 As Collections.ArrayList, Optional 是否忽略空值 As Boolean = True) As System.Collections.ArrayList


        Dim list As System.Collections.ArrayList = New System.Collections.ArrayList

        For Each c In 序列
            If Not list.Contains(c) Then
                list.Add(c)
            End If
        Next


        If 是否忽略空值 = True Then
            list.Remove(Nothing)
        End If
        Return list


    End Function


    ''' <summary>
    ''' 枚举列中互异的元素，列可以有列标题，可指明所在行号即可不参与枚举。列区域必须为单列区域否则返回空序列
    ''' </summary>
    ''' <param name="列区域">要枚举的单列区域，可以包含列标题</param>
    ''' <param name="列标题所在行">列标题所在行号</param>
    ''' <returns>返回列标题行下方的列中互异的元素序列，包含空单元格。列区域不为单列区域时返回空序列</returns>
    Public Function 枚举列中互异元素(列区域 As Excel.Range,
                             Optional 列标题所在行 As Integer = 0) As System.Collections.ArrayList
        Dim 互异元素 As New System.Collections.ArrayList
        If 列区域 Is Nothing Then
            Return 互异元素
        ElseIf 列标题所在行 >= 列区域.Count Or 列区域.Columns.Count <> 1 Then
            Return 互异元素
        End If


        For Each cell As Excel.Range In 列区域.Cells

            If Not 互异元素.Contains(cell.Value) Then
                互异元素.Add(cell.Value)
            End If
        Next
        Return 互异元素

    End Function



    Function Zero(n As Integer) As String
        Dim t As String = ""
        For i = 1 To n
            t &= "0"
        Next
        Return t
    End Function


    Function 居中(Optional range As Excel.Range = Nothing)
        If range Is Nothing Then
            range = app.ActiveSheet
        End If
        range.HorizontalAlignment = Excel.Constants.xlCenter
    End Function

    Function 左对齐(Optional range As Excel.Range = Nothing)
        If range Is Nothing Then
            range = app.ActiveSheet
        End If
        range.HorizontalAlignment = Excel.Constants.xlLeft
    End Function
    Function 右对齐(Optional range As Excel.Range = Nothing)
        If range Is Nothing Then
            range = app.ActiveSheet
        End If
        range.HorizontalAlignment = Excel.Constants.xlRight
    End Function
    Function 背景色(Optional Range As Excel.Range = Nothing, Optional RGBValue As Integer = Nothing)
        If Range Is Nothing Then
            Range = app.ActiveSheet
        End If
        If RGBValue = Nothing Then
            RGBValue = RGB(255, 255, 255)
        End If
        Range.Interior.Color = RGBValue
    End Function
    Function 前景色(Optional Range As Excel.Range = Nothing, Optional RGBValue As Integer = Nothing)
        If Range Is Nothing Then
            Range = app.ActiveSheet
        End If
        If RGBValue = Nothing Then
            RGBValue = RGB(255, 255, 255)
        End If
        With Range.Font
            .Name = "微软雅黑"
            .Size = 11
            .Color = RGBValue
        End With
    End Function


    ''' <summary>    ''' 
    ''' 设置指定区域的字体、字号、颜色
    ''' </summary>
    ''' <param name="Range">指定的区域</param>
    ''' <param name="字体名">字体名称（默认为"微软雅黑"）</param>
    ''' <param name="字体大小">字号大小（默认为11）</param>
    ''' <param name="颜色">文字颜色（默认自动黑色）</param>
    ''' <returns></returns>
    Function 设置字体(Optional Range As Excel.Range = Nothing,
                Optional 字体名 As String = "微软雅黑",
                Optional 字体大小 As Integer = 11,
                Optional 颜色 As Integer = Excel.Constants.xlAutomatic
                )
        If Range Is Nothing Then
            Range = app.ActiveSheet
        End If

        With Range.Font
            .Name = 字体名
            .Size = 字体大小
            .Color = 颜色
        End With
    End Function


    ''' <summary>
    ''' 设置指定区域的四周边框样式
    ''' </summary>
    ''' <param name="Range">指定的区域</param>
    ''' <param name="外边框样式">边框线条的样式（2默认实线）；0 无边框；1虚线边框；2 实线框</param>
    ''' <param name="外边框宽度">边框线粗细（2默认细线）；1 极细；2 细线；3 中等；4 较粗</param>
    ''' <param name="外边框颜色">边框线的颜色；默认自动颜色，可通过RGB(255,255,255)来设置颜色</param>
    ''' <returns></returns>
    Function 设置外边框(Optional Range As Excel.Range = Nothing,
                Optional 外边框样式 As Integer = 2,
                Optional 外边框宽度 As Integer = 2,
                Optional 外边框颜色 As Integer = Excel.Constants.xlAutomatic
                )
        If Range Is Nothing Then
            Range = app.ActiveSheet
        End If

        Dim 宽度 As Excel.XlBorderWeight
        Dim 样式 As Excel.XlLineStyle
        Select Case 外边框样式

            Case 0
                样式 = Excel.XlLineStyle.xlLineStyleNone
            Case 1
                样式 = Excel.XlLineStyle.xlDash
            Case 2
                样式 = Excel.XlLineStyle.xlContinuous
            Case Else
                样式 = Excel.XlLineStyle.xlContinuous
        End Select
        Select Case 外边框宽度
            Case 1
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
            Case 2
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            Case 3
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            Case 4
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            Case Else
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
        End Select

        Range.Borders(Excel.XlBordersIndex.xlEdgeLeft).Color = 外边框颜色
        Range.Borders(Excel.XlBordersIndex.xlEdgeRight).Color = 外边框颜色
        Range.Borders(Excel.XlBordersIndex.xlEdgeTop).Color = 外边框颜色
        Range.Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = 外边框颜色


        Range.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = 宽度
        Range.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = 宽度
        Range.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = 宽度
        Range.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 宽度


        Range.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 样式
        Range.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = 样式
        Range.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = 样式
        Range.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 样式


    End Function

    ''' <summary>
    ''' 设置指定区域的内部边框样式
    ''' </summary>
    ''' <param name="Range">指定的区域</param>
    ''' <param name="内边框样式">边框线条的样式（2默认实线）；0 无边框；1虚线边框；2 实线框</param>
    ''' <param name="内边框宽度">边框线粗细（2默认细线）；1 极细；2 细线；3 中等；4 较粗</param>
    ''' <param name="内边框颜色">边框线的颜色；默认自动颜色，可通过RGB(255,255,255)来设置颜色</param>
    ''' <returns></returns>
    Function 设置内部边框(Optional Range As Excel.Range = Nothing,
                Optional 内边框样式 As Integer = 2,
                Optional 内边框宽度 As Integer = 2,
                Optional 内边框颜色 As Integer = 0
                )
        If Range Is Nothing Then
            Range = app.ActiveSheet
        End If

        Dim 宽度 As Excel.XlBorderWeight
        Dim 样式 As Excel.XlLineStyle
        Select Case 内边框样式

            Case 0
                样式 = Excel.XlLineStyle.xlLineStyleNone
            Case 1
                样式 = Excel.XlLineStyle.xlDash
            Case 2
                样式 = Excel.XlLineStyle.xlContinuous
            Case Else
                样式 = Excel.XlLineStyle.xlContinuous
        End Select
        Select Case 内边框宽度
            Case 1
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
            Case 2
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            Case 3
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            Case 4
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            Case Else
                宽度 = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
        End Select

        Range.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Color = 内边框颜色
        Range.Borders(Excel.XlBordersIndex.xlInsideVertical).Color = 内边框颜色



        Range.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = 宽度
        Range.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = 宽度



        Range.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = 样式
        Range.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = 样式



    End Function
    ''' <summary>
    ''' 自动设置指定表格数据区域指定的行和列为表头格式
    ''' </summary>
    ''' <param name="sheet">指定的表</param>
    ''' <param name="Row">表头所在的行号</param>
    ''' <param name="Column">表头所在的列号</param>
    ''' <returns></returns>
    Public Function 设置为表头样式(Optional sheet As Excel.Worksheet = Nothing,
                           Optional Row As Integer = 1,
                           Optional Column As Integer = 1,
                            Optional FontName As String = "微软雅黑",
                            Optional FontSize As Integer = 0,
                            Optional FontColor As Integer = Nothing,
                            Optional BackColor As Integer = Nothing
                         )
        If sheet Is Nothing Then
            sheet = app.ActiveSheet
        End If
        Dim NumRow, NumColumn As Integer
        Dim Range, RowRange, ColumnRangeAs
        Range = 获取用户区域(sheet)

        NumRow = Range.Rows.Count
        NumColumn = Range.Columns.Count

        If FontSize <= 1 Then
            FontSize = 12
        End If

        If FontColor = Nothing Then
            FontColor = RGB(0, 0, 0)
        End If


        If Column > 0 Then
            ColumnRangeAs = sheet.Range(sheet.Cells(Math.Max(Row, 1), Column), sheet.Cells(NumRow, Column))
            设置字体(ColumnRangeAs, FontName, FontSize, FontColor)
            设置外边框(ColumnRangeAs, 1, 3, RGB(0, 0, 0))
            设置内部边框(ColumnRangeAs, 2, 2, RGB(0, 0, 0))
            设置填充色(ColumnRangeAs, My.Settings.表头列背景色)
            居中(ColumnRangeAs)
        End If

        If Row > 0 Then
            RowRange = sheet.Range(sheet.Cells(Row, Math.Max(Column, 1)), sheet.Cells(Row, NumColumn))
            设置字体(RowRange, FontName, FontSize, FontColor)
            设置外边框(RowRange, 2, 3, RGB(0, 0, 0))
            设置内部边框(RowRange, 2, 2, RGB(0, 0, 0))
            设置填充色(RowRange, My.Settings.表头行背景色)
            居中(RowRange)
        End If





        自动列宽(sheet)

    End Function
    Public Function 是否存在并激活(SheetName As String) As Boolean
        If 是否存在工作表(SheetName) And app.ActiveSheet.name = SheetName Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function 打开设置页(Optional 标题 As String = Nothing) As Excel.Worksheet
        If 标题 = Nothing Then
            设置页标题 = "插件公共设置"
        Else
            设置页标题 = 标题
        End If
        Dim SetSheet As Excel.Worksheet
        If 是否存在工作表(设置页名) Then
            SetSheet = app.Sheets(设置页名)
            SetSheet.Delete()
        End If
        SetSheet = 新建工作表(设置页名)

        SetSheet.Activate()
        Dim TempRange As Excel.Range = MyRange(SetSheet, 1, 1, 1, 9)
        TempRange.Merge()
        TempRange.Value = 设置页标题
        设置填充色(TempRange, RGB(160, 200, 255))
        设置字体(TempRange,, 18, RGB(0, 0, 255))
        设置外边框(TempRange, 2, 4, RGB(0, 0, 0))
        居中(TempRange)





        Return SetSheet

    End Function

    ''' <summary>
    ''' 当前是否存在参数"标题"所指定的设置页面表格
    ''' </summary>
    ''' <param name="标题"></param>
    ''' <returns>存在设置页面表就返回此设置表，不存在返回Nothing</returns>
    Public Function 是否已经设置过(标题 As String) As Excel.Worksheet
        Dim SetSheet As Excel.Worksheet
        If 是否存在并激活(设置页名) Then
            SetSheet = app.Sheets(设置页名)
            If SetSheet.Cells(1.1).value = 标题 Then
                Return SetSheet
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Function 合并两个区域(Range1 As Excel.Range, Range2 As Excel.Range) As Excel.Range
        Return app.Union(Range1, Range2)
    End Function

    ''' <summary>
    ''' 对指定区域填充指定颜色,颜色可用函数RGB(Red As integet, Green As integet, Blue As integet)生成
    ''' </summary>
    ''' <param name="Range">要涂色的区域</param>
    ''' <param name="颜色">要涂的颜色，颜色可用函数RGB(Red As integet, Green As integet, Blue As integet)生成</param>
    Public Sub 设置填充色(Optional Range As Excel.Range = Nothing, Optional 颜色 As Integer = 100000)
        Range.Interior.Color = 颜色
    End Sub
    Public Function 自动列宽(Optional sheet As Excel.Worksheet = Nothing)
        If sheet Is Nothing Then
            sheet = app.ActiveSheet
        End If
        sheet.Cells.EntireColumn.AutoFit()
    End Function
    Public Function 自动行高(Optional sheet As Excel.Worksheet = Nothing)
        If sheet Is Nothing Then
            sheet = app.ActiveSheet
        End If
        sheet.Cells.EntireRow.AutoFit()
    End Function
    Public Sub runtimer()
        app.OnTime(System.DateTime.Now.AddSeconds(1), Procedure:="hhtt")

    End Sub

    Public Sub hhtt()
        app.Range("a1").Value = Format(Now(), "hh,mm,ss")

    End Sub

    ''' <summary>
    ''' 获取指定表中第ColumnNum列的整个区域
    ''' </summary>
    ''' <param name="sheet">指定的表</param>
    ''' <param name="ColumnNum">列号</param>
    ''' <returns>返回 指定表中第ColumnNum列的整个区域</returns>
    Public Function 列区域(sheet As Excel.Worksheet, ColumnNum As Integer) As Excel.Range
        Dim EndCell As Excel.Range = 获取结束单元格(sheet)
        Dim MaxNumOfRow, MaxNumOfColumn As Integer
        MaxNumOfRow = EndCell.Row
        MaxNumOfColumn = EndCell.Column


        Return sheet.Range(sheet.Cells(1, ColumnNum), sheet.Cells(MaxNumOfRow, ColumnNum).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
    End Function
    ''' <summary>
    ''' 获取指定表中第RowNum行的整个区域
    ''' </summary>
    ''' <param name="sheet">指定的表</param>
    ''' <param name="RowNum">行号</param>
    ''' <returns>返回 指定表中第RowNum行的整个区域</returns>
    Public Function 行区域(sheet As Excel.Worksheet, RowNum As Integer) As Excel.Range
        Dim EndCell As Excel.Range = 获取结束单元格(sheet)
        Dim MaxNumOfRow, MaxNumOfColumn As Integer
        MaxNumOfRow = EndCell.Row
        MaxNumOfColumn = EndCell.Column

        Return sheet.Range(sheet.Cells(RowNum, 1), sheet.Cells(RowNum, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
    End Function
    ''' <summary>
    ''' 求两个区域的交集区域
    ''' </summary>
    ''' <param name="Range1">第一个区域</param>
    ''' <param name="Range2">第二个区域</param>
    ''' <returns>两个区域的交集区域返回 两个区域的交集区域</returns>
    Public Function 区域交集(Range1 As Excel.Range, Range2 As Excel.Range) As Excel.Range
        Return app.Intersect(Range1, Range2)
    End Function
    ''' <summary>
    ''' 获取区域Range1中舍去属于区域Range2的部分所剩余的区域
    ''' </summary>
    ''' <param name="Range1">源区域</param>
    ''' <param name="RowNum">要舍去的区域</param>
    ''' <returns>返回 区域Range1中舍去属于区域Range2的部分所剩余的区域</returns>
    Public Function 区域移除行(Range1 As Excel.Range, RowNum As Integer) As Excel.Range
        Dim sheet As Excel.Worksheet = Range1.Worksheet
        Dim EndCell As Excel.Range = 获取结束单元格(sheet)
        Dim MaxNumOfRow, MaxNumOfColumn As Integer
        MaxNumOfRow = EndCell.Row
        MaxNumOfColumn = EndCell.Column



        Dim r1, r2, r3 As Excel.Range
        If RowNum > 1 Then
            r1 = sheet.Range(sheet.Cells(1, 1), sheet.Cells(RowNum - 1, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
            r2 = sheet.Range(sheet.Cells(RowNum + 1, 1), sheet.Cells(MaxNumOfRow, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
            r3 = app.Union(r1, r2)
        Else
            r3 = sheet.Range(sheet.Cells(RowNum + 1, 1), sheet.Cells(MaxNumOfRow, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
        End If
        Return app.Intersect(Range1, r3)


    End Function
    Public Function 区域移除行(Range1 As Excel.Range, RowNum1 As Integer, RowNum2 As Integer) As Excel.Range
        Dim range As Excel.Range = Range1
        If RowNum1 > 0 Then
            range = 区域移除行(Range1, RowNum1)
        End If
        If RowNum2 > 0 Then
            range = 区域移除行(range, RowNum2)
        End If
        Return range
    End Function
    Public Function 区域移除行(Range1 As Excel.Range, RowNum1 As Integer, RowNum2 As Integer, RowNum3 As Integer) As Excel.Range
        Dim range As Excel.Range = Range1
        If RowNum1 > 0 Then
            range = 区域移除行(Range1, RowNum1)
        End If
        If RowNum2 > 0 Then
            range = 区域移除行(range, RowNum2)
        End If
        If RowNum3 > 0 Then
            range = 区域移除行(range, RowNum3)
        End If
        Return range
    End Function


    Public Function 区域移除列(Range1 As Excel.Range, ColumnNum As Integer) As Excel.Range
        Dim sheet As Excel.Worksheet = Range1.Worksheet
        Dim EndCell As Excel.Range = 获取结束单元格(sheet)
        Dim MaxNumOfRow, MaxNumOfColumn As Integer
        MaxNumOfRow = EndCell.Row
        MaxNumOfColumn = EndCell.Column



        Dim r1, r2, r3 As Excel.Range
        If ColumnNum > 1 Then
            r1 = sheet.Range(sheet.Cells(1, 1), sheet.Cells(MaxNumOfRow, ColumnNum - 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            r2 = sheet.Range(sheet.Cells(1, ColumnNum + 1), sheet.Cells(MaxNumOfRow, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
            r3 = app.Union(r1, r2)
        Else
            r3 = sheet.Range(sheet.Cells(1, ColumnNum + 1), sheet.Cells(MaxNumOfRow, MaxNumOfColumn).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
        End If
        Return app.Intersect(Range1, r3)


    End Function
    Public Function 区域移除列(Range1 As Excel.Range, ColumnNum1 As Integer, ColumnNum2 As Integer) As Excel.Range
        Dim range As Excel.Range = Range1
        If ColumnNum1 > 0 Then
            range = 区域移除列(Range1, ColumnNum1)
        End If

        If ColumnNum2 > 0 Then
            range = 区域移除列(range, ColumnNum2)
        End If
        Return range
    End Function
    Public Function 区域移除列(Range1 As Excel.Range, ColumnNum1 As Integer, ColumnNum2 As Integer, ColumnNum3 As Integer) As Excel.Range
        Dim range As Excel.Range = Range1
        If ColumnNum1 > 0 Then
            range = 区域移除列(Range1, ColumnNum1)
        End If

        If ColumnNum2 > 0 Then
            range = 区域移除列(range, ColumnNum2)
        End If
        If ColumnNum3 > 0 Then
            range = 区域移除列(range, ColumnNum3)
        End If
        Return range
    End Function



    Public Function 获取循环色(色号 As Integer) As Integer
        If My.Settings.循环色序列 IsNot Nothing Then
            Dim Num As Integer = My.Settings.循环色序列.Count
            If Num > 0 Then
                Return My.Settings.循环色序列(色号 Mod Num)
            Else Return RGB(255, 255, 255)
            End If
        Else
            Return RGB(255, 255, 255)
        End If
    End Function
    Public Function 添加循环色(颜色 As Integer) As Integer
        If My.Settings.循环色序列 Is Nothing Then
            My.Settings.循环色序列 = New Collections.ArrayList
        End If
        My.Settings.循环色序列.Add(颜色)
        Return My.Settings.循环色序列.Count
    End Function


    Public Sub 创建属性表设置页(SetSheet As Excel.Worksheet, 所选区域 As Excel.Range)

        Dim tempRange As Excel.Range
        Dim Green As Integer = RGB(200, 255, 200)
        Dim Blue As Integer = RGB(50, 50, 255)
        My.Settings.枚举序列 = 枚举选区(所选区域)
        Dim row As Integer = 1
        SetSheet.Cells(2, 1) = "枚举序列(共" & My.Settings.枚举序列.Count & "个)"
        设置填充色(SetSheet.Cells(2, 1), Green)
        row = 3
        For Each t In My.Settings.枚举序列
            SetSheet.Cells(row, 1) = t
            row += 1
        Next







        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''以下为创建  属性  表框架'''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        设置字体(MyRange(SetSheet, 2, 3, 12, 10), "微软雅黑", 11, RGB(0, 0, 0))

        tempRange = MyRange(SetSheet, 2, 3, 2, 5)
        tempRange.Merge()
        tempRange.Value = "属性匹配表名"
        设置字体(tempRange, "华文隶书", 11, RGB(0, 0, 0))
        设置填充色(tempRange, RGB(200, 200, 200))

        tempRange = MyRange(SetSheet, 2, 6, 2, 8)
        tempRange.Merge()
        tempRange.Value = My.Settings.匹配属性表
        SetSheet.Cells(2, 9) = "行号"
        SetSheet.Cells(11, 2) = "列号"
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        设置填充色(tempRange, Green)



        tempRange = MyRange(SetSheet, 3, 3, 5, 5)
        tempRange.Merge()
        tempRange.Value = "表头"
        设置填充色(tempRange, Green)

        tempRange = MyRange(SetSheet, 3, 6)
        tempRange.Merge()
        tempRange.Value = "1级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 3, 7, 3, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表1级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 1级行 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)





        tempRange = MyRange(SetSheet, 4, 6)
        tempRange.Merge()
        tempRange.Value = "2级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 4, 7, 4, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表2级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 2级行 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)

        tempRange = MyRange(SetSheet, 5, 6)
        tempRange.Merge()
        tempRange.Value = "3级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 5, 7, 5, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表3级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 3级行 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)


        SetSheet.Cells(3, 9) = My.Settings.属性表1级表头行号
        SetSheet.Cells(4, 9) = My.Settings.属性表2级表头行号
        SetSheet.Cells(5, 9) = My.Settings.属性表3级表头行号

        tempRange = MyRange(SetSheet, 6, 3, 7, 3)
        tempRange.Merge()
        tempRange.Value = "1级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8, 3, 10, 3)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 10, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表1级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 1级列 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)

        tempRange = MyRange(SetSheet, 6, 4, 7, 4)
        tempRange.Merge()
        tempRange.Value = "2级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8, 4, 10, 4)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 10, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表2级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 2级列 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)

        tempRange = MyRange(SetSheet, 6, 5, 7, 5)
        tempRange.Merge()
        tempRange.Value = "3级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8, 5, 10, 5)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 10, RGB(255, 0, 0))
        tempRange.Value = My.Settings.属性表3级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 3级列 字段名" & Chr(10) & "以便与数据表中字段相匹配")
        设置填充色(tempRange, Green)



        tempRange = MyRange(SetSheet, 6, 6, 10, 8)
        tempRange.Merge()
        tempRange.Value = My.Settings.属性表数据名称
        设置字体(tempRange, "微软雅黑", 16, RGB(255, 0, 0))



        SetSheet.Cells(11, 3) = My.Settings.属性表1级表头列号
        SetSheet.Cells(11, 4) = My.Settings.属性表2级表头列号
        SetSheet.Cells(11, 5) = My.Settings.属性表3级表头列号

        MyRange(SetSheet, 6, 3, 10, 5).Orientation = -85


        设置字体(MyRange(SetSheet, 11, 3, 11, 5), "微软雅黑", 11, RGB(255, 0, 0))
        设置字体(MyRange(SetSheet, 3, 9, 5, 9), "微软雅黑", 11, RGB(255, 0, 0))
        设置字体(MyRange(SetSheet, 3, 3, 5, 5), "微软雅黑", 16, RGB(0, 0, 0))


        设置内部边框(MyRange(SetSheet, 3, 3, 10, 8), 2, 2, RGB(0, 0, 0))
        设置外边框(MyRange(SetSheet, 3, 3, 10, 8), 2, 4, RGB(0, 0, 0))


        MyRange(SetSheet, 2, 10).Value = "结果页表头列序号"
        MyRange(SetSheet, 3, 10).Value = My.Settings.结果页列序号_属性表1级表头行号
        MyRange(SetSheet, 4, 10).Value = My.Settings.结果页列序号_属性表2级表头行号
        MyRange(SetSheet, 5, 10).Value = My.Settings.结果页列序号_属性表3级表头行号

        MyRange(SetSheet, 12, 2).Value = "结果页表头列序号"
        MyRange(SetSheet, 12, 3).Value = My.Settings.结果页列序号_属性表1级表头列号
        MyRange(SetSheet, 12, 4).Value = My.Settings.结果页列序号_属性表2级表头列号
        MyRange(SetSheet, 12, 5).Value = My.Settings.结果页列序号_属性表3级表头列号


        tempRange = app.Union(MyRange(SetSheet, 3, 10, 5, 10), (MyRange(SetSheet, 12, 3, 12, 5)))
        设置字体(tempRange, "微软雅黑", 11, Blue)


        tempRange = app.Union(MyRange(SetSheet, 3, 9, 5, 10),
                                  MyRange(SetSheet, 11, 3, 12, 5))
        With tempRange.Validation
            .Delete()
            .Add(Excel.XlDVType.xlValidateWholeNumber,, , Formula1:="0", Formula2:="1000000")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "超出范围"
            .InputMessage = ""
            .ErrorMessage = "必须输入0~1000000的整数"
            .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With









        SetSheet.Cells.EntireColumn.AutoFit() '自动列宽
        '设置为表头样式(SetSheet, 2, 0)
        SetSheet.Cells(3, 1).Select()


        居中(SetSheet.UsedRange)
        右对齐(列区域(SetSheet, 2))



    End Sub


    Public Function 添加按钮(ButtonRange As Excel.Range,
                         Optional Text As String = "按钮",
                         Optional ButtonName As String = "ButtonName",
                         Optional 附加信息 As String = "") As Windows.Forms.Button



        Dim worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(ButtonRange.Worksheet)


        If Not (ButtonRange Is Nothing) Then
            Dim button As New Windows.Forms.Button
            worksheet.Controls.AddControl(button, ButtonRange, ButtonName & Now.Ticks)
            button.Text = Text
            button.Tag = 附加信息 '存放单元格地址信息
            button.Font = New Drawing.Font("微软雅黑", 12)
            button.BackColor = Drawing.Color.FromArgb(100, 200, 200, 250)
            button.AutoSize = True

            AddHandler button.Click, AddressOf MyButton_Click
            AddHandler button.MouseEnter, AddressOf MyButton_MouseEnter
            AddHandler button.MouseLeave, AddressOf MyButton_MouseLeave
            'button.Left = 200
            'button.Top = 300
            Return button
        Else
            Return Nothing
        End If



    End Function



    Private Sub MyButton_Click(sender As Object, e As EventArgs)
        Try
            Dim button As Windows.Forms.Button = sender
            Dim sheet As Excel.Worksheet = app.ActiveSheet
            Dim 标题 As String = sheet.Cells(1, 1).value
            If button.Text = "选择区域" Then
                地址单元格 = app.ActiveSheet.range(button.Tag)
                MySelectForm.Show()

            Else
                If 标题 = "表间匹配" Then
                    执行表间匹配()
                ElseIf 标题 = "批量匹配属性" Then
                    执行批量匹配属性()
                ElseIf 标题 = "插件系统设置" Then
                    执行插件系统设置()
                ElseIf 标题 = "生成编号" Then
                    执行插生成编号()

                ElseIf 标题 = "开始涂色" Then
                    设置填充色(app.Range(app.Sheets(设置页名).cells(2, 2).value), RGB(122, 0, 0))
                Else

                End If

            End If



        Catch ex As Exception
            错误信息序列.Add("执行操作时出现错误，当前操作意外终止！")
        End Try



        If 错误信息序列.Count > 0 Then
            MsgBox(列表转字符串(错误信息序列, Chr(10)))
            错误信息序列.Clear()
        End If

    End Sub

    Private Sub MyButton_MouseEnter(sender As Object, e As EventArgs)
        Dim button As Windows.Forms.Button = sender
        button.ForeColor = Drawing.Color.Red
        button.BackColor = Drawing.Color.FromArgb(100, 170, 255, 180)
        button.Font = New Drawing.Font("微软雅黑", 12)
    End Sub

    Private Sub MyButton_MouseLeave(sender As Object, e As EventArgs)
        Dim button As Windows.Forms.Button = sender
        button.ForeColor = Drawing.Color.Black
        button.BackColor = Drawing.Color.FromArgb(100, 200, 200, 250)
        button.Font = New Drawing.Font("微软雅黑", 12)
    End Sub





    Public Function 创建结果页并显示结果(Optional 结果页名 As String = "插件执行结果",
                               Optional 缓冲表 As Excel.Worksheet = Nothing) As Excel.Worksheet


        If 缓冲表 Is Nothing Then
            If 是否存在工作表(缓冲表名) Then
                缓冲表 = app.Sheets(缓冲表名)
            Else
                错误信息序列.Add("显示结果页时，没有指明缓冲表，也找不到缓冲表！")
                Return Nothing
            End If
        End If
        'MsgBox("1 缓冲表名字:" & 缓冲表.Name)
        '缓冲表.Copy(app.Sheets(1))
        'MsgBox("2 缓冲表名字:" & 缓冲表.Name)
        'Dim ResultSheet As Excel.Worksheet = app.Sheets(1)
        'MsgBox("3 ResultSheet名字:" & ResultSheet.Name)
        If 是否存在工作表(结果页名) Then
            app.DisplayAlerts = False '删除工作表不再提示
            app.Sheets(结果页名).Delete() '删除工作表
            app.DisplayAlerts = True  '恢复删除工作表提示
        End If
        缓冲表.Name = 结果页名
        缓冲表.Activate()
        'MsgBox("4 缓冲表名字:" & 缓冲表.Name)
        'MsgBox("3 ResultSheet名字:" & ResultSheet.Name)

        'If 是否删除缓冲表 = True Then
        '    app.DisplayAlerts = False '删除工作表不再提示
        '    缓冲表.Delete()
        '    app.DisplayAlerts = True  '恢复删除工作表提示
        'End If


        Return 缓冲表

    End Function
    Public Function 执行插生成编号() As Excel.Worksheet

        Dim Setsheet As Excel.Worksheet = app.ActiveSheet

        Dim CurrentRow As Integer = 3
        Dim MaxRow = 获取用户区域(Setsheet).Rows.Count
        Dim 前缀, 后缀, 班级， 班级代码 As String
        Dim 个人起始号, 个号位数， 班号位数， 人数 As Integer

        For i = 3 To MaxRow
            If Setsheet.Cells(i, 1).Value <> Nothing Then
                班级 = Setsheet.Cells(i, 1).Value
            Else
                Exit For
            End If

            If Setsheet.Cells(i, 2).Value <> Nothing Then
                人数 = Setsheet.Cells(i, 2).Value
            Else
                Exit For
            End If

            If Setsheet.Cells(i, 3).Value <> Nothing Then
                前缀 = Setsheet.Cells(i, 3).Value
            Else
                前缀 = ""
            End If

            If Setsheet.Cells(i, 4).Value <> Nothing Then
                班号位数 = Setsheet.Cells(i, 4).Value
                班级代码 = Format$(Int(班级)， Zero(班号位数))
            Else
                班级代码 = ""
            End If


            If Setsheet.Cells(i, 5).Value <> Nothing Then
                个人起始号 = Setsheet.Cells(i, 5).Value
            Else
                个人起始号 = 1
            End If

            If Setsheet.Cells(i, 6).Value <> Nothing Then
                个号位数 = Setsheet.Cells(i, 6).Value
            Else
                个号位数 = 3
            End If

            If Setsheet.Cells(i, 7).Value <> Nothing Then
                后缀 = Setsheet.Cells(i, 7).Value
            Else
                后缀 = ""
            End If



            For n = 1 To Int(Setsheet.Cells(i, 2).Value)
                Setsheet.Cells(CurrentRow, 9) = CurrentRow - 2
                Setsheet.Cells(CurrentRow, 10) = 班级
                Setsheet.Cells(CurrentRow, 11) = 前缀 & 班级代码 & Format$(个人起始号 + n - 1, Zero(个号位数)) & 后缀
                Setsheet.Cells(CurrentRow, 12) = n
                CurrentRow += 1
            Next

        Next
        Setsheet.Cells.EntireColumn.AutoFit() '自动列宽
        居中(Setsheet.UsedRange)
    End Function

    Public Function 执行插件系统设置() As Excel.Worksheet
        Dim SetSheet As Excel.Worksheet
        If 是否存在工作表(设置页名) Then
            SetSheet = app.Sheets(设置页名)
        Else
            错误信息序列.Add("未找到设置页！")
            Return Nothing
        End If



        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        '''
        For i = 1 To My.Settings.循环色序列.Count
            My.Settings.循环色序列(i - 1) = SetSheet.Cells(3, i + 1).Interior.Color
        Next
        My.Settings.表头行背景色 = SetSheet.Cells(2， 2).Interior.Color
        My.Settings.表头列背景色 = SetSheet.Cells(3， 1).Interior.Color

        My.Settings.字符串分隔符 = MyRange(SetSheet, 11, 2).Value

        My.Settings.是否剔除重复值 = MyRange(SetSheet, 12, 2).Value

        My.Settings.是否忽略空值 = MyRange(SetSheet, 13, 2).Value

        My.Settings.是否忽略首尾空白字符 = MyRange(SetSheet, 14, 2).Value





        My.Settings.Save()

        app.DisplayAlerts = False '删除工作表不再提示
        SetSheet.Delete() '删除工作表

        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码


        app.DisplayAlerts = True  '删除工作表回复提示



    End Function
    Public Function 执行批量匹配属性() As Excel.Worksheet

        Dim SetSheet As Excel.Worksheet
        If 是否存在工作表(设置页名) Then
            SetSheet = app.Sheets(设置页名)
        Else
            错误信息序列.Add("未找到设置页！")
            Return Nothing
        End If

        读取属性表设置(SetSheet)

        app.DisplayAlerts = False '删除工作表不再提示
        SetSheet.Delete() '删除工作表

        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码

        SetSheet = 新建工作表("属性匹配结果")
        SetSheet.Cells.Delete()

        Dim i As Integer = 2
        Dim j As Integer = 1
        Dim HeadList As Collections.ArrayList
        Dim 一条记录 As New 记录

        For Each NameText In My.Settings.枚举序列

            一条记录.key = NameText
            HeadList = 由值获取表头(app.Sheets(My.Settings.匹配属性表),
                                                       NameText,
                                                       My.Settings.属性表1级表头行号,
                                                       My.Settings.属性表1级表头列号，
                                                       My.Settings.属性表2级表头行号，
                                                       My.Settings.属性表2级表头列号，
                                                       My.Settings.属性表3级表头行号，
                                                       My.Settings.属性表3级表头列号)


            Dim data As New Collections.ArrayList
            Dim 数据个数 As Integer = 0
            data.Add(HeadList.Count)

            For Each head As 表头 In HeadList
                Dim 表头和数据 As New 表头和数据类
                表头和数据.Add(head)
                一条记录.添加数据(表头和数据)
            Next
            If 一条记录.数据个数() > 数据个数 Then
                一条记录.以查询表头记录形式写入记录表头(SetSheet, 1)
                数据个数 = 一条记录.数据个数()
            End If

            一条记录.以查询表头记录形式写入(SetSheet, i)

            'SetSheet.Cells(i, 1) = NameText
            'SetSheet.Cells(i, 2) = HeadList.Count

            '列表写入表格(HeadList, SetSheet, 0, i, 3)
            一条记录.数据列.Clear()
            i += 1
        Next





        SetSheet.Cells.EntireColumn.AutoFit() '自动列宽
        SetSheet.Cells.EntireRow.AutoFit() '自动行高宽
        设置为表头样式(SetSheet, 1, 1)
        居中(SetSheet.Cells)
        设置内部边框(SetSheet.UsedRange, 2, 2, RGB(0, 0, 0))
        设置外边框(SetSheet.UsedRange, 2, 4, RGB(0, 0, 0))
        app.DisplayAlerts = True  '删除工作表回复提示
    End Function


    Public Function 执行表间匹配() As Excel.Worksheet


        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        ''''''这里是 读取设置数据的代码
        '''
        Dim SetSheet As Excel.Worksheet
        If 是否存在工作表(设置页名) Then
            SetSheet = app.Sheets(设置页名)
        Else
            错误信息序列.Add("未找到设置页！")
            Return Nothing
        End If

        读取属性表设置(SetSheet)
        读取数据表设置(SetSheet)




        app.DisplayAlerts = False '删除工作表不再提示
        SetSheet.Delete() '删除工作表

        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码

        Dim result As Integer
        result = 表间批量匹配数据2(app.Sheets(My.Settings.匹配属性表),
                           app.Sheets(My.Settings.匹配数据表),
                           My.Settings.枚举序列)

        MsgBox("共统计了 " & result & "个数据记录。")
        My.Settings.枚举序列.Clear()

        app.DisplayAlerts = True  '删除工作表回复提示
        Return SetSheet
    End Function
    Public Sub 创建数据表设置页(SetSheet As Excel.Worksheet)
        Dim tempRange As Excel.Range
        Dim Green As Integer = RGB(200, 255, 200)
        Dim Blue As Integer = RGB(50, 50, 255)


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''以下为创建  数据  表框架'''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim n As Integer = 12
        设置字体(MyRange(SetSheet, 2 + n, 3, 12 + n, 10), "微软雅黑", 11, RGB(0, 0, 0))

        tempRange = MyRange(SetSheet, 2 + n, 3, 2 + n, 5)
        tempRange.Merge()
        tempRange.Value = "数据匹配表名"
        设置字体(tempRange, "华文隶书", 11, RGB(0, 0, 0))
        设置填充色(tempRange, RGB(200, 200, 200))

        tempRange = MyRange(SetSheet, 2 + n, 6, 2 + n, 8)
        tempRange.Merge()
        tempRange.Value = My.Settings.匹配数据表
        SetSheet.Cells(2 + n, 9) = "行号"
        'SetSheet.Cells(2 + n, 10) = "表头行值"
        SetSheet.Cells(11 + n, 2) = "列号"
        'SetSheet.Cells(12 + n, 2) = "表头列值"
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        设置填充色(tempRange, RGB(200, 200, 200))



        tempRange = MyRange(SetSheet, 3 + n, 3, 5 + n, 5)
        tempRange.Merge()
        tempRange.Value = "表头"
        设置填充色(tempRange, Green)



        tempRange = MyRange(SetSheet, 3 + n, 6)
        tempRange.Merge()
        tempRange.Value = "1级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 3 + n, 7, 3 + n, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表1级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 1级行 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)


        tempRange = MyRange(SetSheet, 4 + n, 6)
        tempRange.Merge()
        tempRange.Value = "2级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 4 + n, 7, 4 + n, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表2级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 2级行 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)


        tempRange = MyRange(SetSheet, 5 + n, 6)
        tempRange.Merge()
        tempRange.Value = "3级行"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 5 + n, 7, 5 + n, 8)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表3级表头行名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 3级行 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)


        SetSheet.Cells(3 + n, 9) = My.Settings.数据表1级表头行号
        SetSheet.Cells(4 + n, 9) = My.Settings.数据表2级表头行号
        SetSheet.Cells(5 + n, 9) = My.Settings.数据表3级表头行号


        'SetSheet.Cells(3 + n, 10) = My.Settings.数据表1级表头行名
        'SetSheet.Cells(4 + n, 10) = My.Settings.数据表2级表头行名
        'SetSheet.Cells(5 + n, 10) = My.Settings.数据表3级表头行名


        tempRange = MyRange(SetSheet, 6 + n, 3, 7 + n, 3)
        tempRange.Merge()
        tempRange.Value = "1级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8 + n, 3, 10 + n, 3)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表1级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 1级列 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)


        tempRange = MyRange(SetSheet, 6 + n, 4, 7 + n, 4)
        tempRange.Merge()
        tempRange.Value = "2级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8 + n, 4, 10 + n, 4)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表2级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 2级列 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)


        tempRange = MyRange(SetSheet, 6 + n, 5, 7 + n, 5)
        tempRange.Merge()
        tempRange.Value = "3级列"
        设置填充色(tempRange, Green)
        tempRange = MyRange(SetSheet, 8 + n, 5, 10 + n, 5)
        tempRange.Merge()
        设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))
        tempRange.Value = My.Settings.数据表3级表头列名
        tempRange(1, 1).AddComment("heting:[选填内容]" & Chr(10) & "在这里输入 3级列 字段名" & Chr(10) & "可以与属性表表中字段相匹配")
        设置填充色(tempRange, Green)



        tempRange = MyRange(SetSheet, 6 + n, 6, 10 + n, 8)
        tempRange.Merge()
        tempRange.Value = My.Settings.数据表数据名称
        设置字体(tempRange, "微软雅黑", 16, RGB(255, 0, 0))





        SetSheet.Cells(11 + n, 3) = My.Settings.数据表1级表头列号
        SetSheet.Cells(11 + n, 4) = My.Settings.数据表2级表头列号
        SetSheet.Cells(11 + n, 5) = My.Settings.数据表3级表头列号

        'SetSheet.Cells(12 + n, 3) = My.Settings.数据表1级表头列名
        'SetSheet.Cells(12 + n, 4) = My.Settings.数据表2级表头列名
        'SetSheet.Cells(12 + n, 5) = My.Settings.数据表3级表头列名

        MyRange(SetSheet, 6 + n, 3, 10 + n, 5).Orientation = -85


        设置字体(MyRange(SetSheet, 11 + n, 3, 11 + n, 5), "微软雅黑", 11, RGB(255, 0, 0))
        设置字体(MyRange(SetSheet, 3 + n, 9, 5 + n, 9), "微软雅黑", 11, RGB(255, 0, 0))
        设置字体(MyRange(SetSheet, 3 + n, 3, 5 + n, 5), "微软雅黑", 16, RGB(0, 0, 0))

        设置内部边框(MyRange(SetSheet, 3 + n, 3, 10 + n, 8), 2, 2, RGB(0, 0, 0))
        设置外边框(MyRange(SetSheet, 3 + n, 3, 10 + n, 8), 2, 4, RGB(0, 0, 0))


        MyRange(SetSheet, 2 + n, 10).Value = "结果页表头列序号"
        MyRange(SetSheet, 3 + n, 10).Value = My.Settings.结果页列序号_数据表1级表头行号
        MyRange(SetSheet, 4 + n, 10).Value = My.Settings.结果页列序号_数据表2级表头行号
        MyRange(SetSheet, 5 + n, 10).Value = My.Settings.结果页列序号_数据表3级表头行号

        MyRange(SetSheet, 12 + n, 2).Value = "结果页表头列序号"
        MyRange(SetSheet, 12 + n, 3).Value = My.Settings.结果页列序号_数据表1级表头列号
        MyRange(SetSheet, 12 + n, 4).Value = My.Settings.结果页列序号_数据表2级表头列号
        MyRange(SetSheet, 12 + n, 5).Value = My.Settings.结果页列序号_数据表3级表头列号

        tempRange = app.Union(MyRange(SetSheet, 3 + n, 10, 5 + n, 10), MyRange(SetSheet, 12 + n, 3, 12 + n, 5))
        设置字体(tempRange, "微软雅黑", 11, Blue)


        tempRange = app.Union(MyRange(SetSheet, 3 + n, 9, 5 + n, 10),
                              MyRange(SetSheet, 11 + n, 3, 12 + n, 5))
        With tempRange.Validation
            .Delete()
            .Add(Excel.XlDVType.xlValidateWholeNumber,, , Formula1:="0", Formula2:="1000000")
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = "超出范围"
            .InputMessage = ""
            .ErrorMessage = "必须输入0~1000000的整数"
            .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With





        SetSheet.Cells.EntireColumn.AutoFit() '自动列宽
        '设置为表头样式(SetSheet, 2, 0)
        SetSheet.Cells(3, 1).Select()


        居中(SetSheet.UsedRange)
        右对齐(列区域(SetSheet, 2))


    End Sub



    Public Function ToStr(cell As Excel.Range)
        If cell.Value = Nothing Then
            Return ""
        Else
            Return cell.Value.ToString
        End If
    End Function
    Public Sub 读取属性表设置(SetSheet As Excel.Worksheet)
        My.Settings.匹配属性表 = SetSheet.Cells(2, 6).value


        My.Settings.属性表1级表头行号 = SetSheet.Cells(3, 9).value
        My.Settings.属性表2级表头行号 = SetSheet.Cells(4, 9).value
        My.Settings.属性表3级表头行号 = SetSheet.Cells(5, 9).value


        My.Settings.属性表1级表头列号 = SetSheet.Cells(11, 3).value
        My.Settings.属性表2级表头列号 = SetSheet.Cells(11, 4).value
        My.Settings.属性表3级表头列号 = SetSheet.Cells(11, 5).value

        My.Settings.属性表1级表头行名 = SetSheet.Cells(3, 7).value
        My.Settings.属性表2级表头行名 = SetSheet.Cells(4, 7).value
        My.Settings.属性表3级表头行名 = SetSheet.Cells(5, 7).value

        My.Settings.属性表1级表头列名 = SetSheet.Cells(8, 3).value
        My.Settings.属性表2级表头列名 = SetSheet.Cells(8, 4).value
        My.Settings.属性表3级表头列名 = SetSheet.Cells(8, 5).value

        My.Settings.结果页列序号_属性表1级表头行号 = MyRange(SetSheet, 3, 10).Value
        My.Settings.结果页列序号_属性表2级表头行号 = MyRange(SetSheet, 4, 10).Value
        My.Settings.结果页列序号_属性表3级表头行号 = MyRange(SetSheet, 5, 10).Value

        My.Settings.结果页列序号_属性表1级表头列号 = MyRange(SetSheet, 12, 3).Value
        My.Settings.结果页列序号_属性表2级表头列号 = MyRange(SetSheet, 12, 4).Value
        My.Settings.结果页列序号_属性表3级表头列号 = MyRange(SetSheet, 12, 5).Value



        My.Settings.属性表数据名称 = SetSheet.Cells(6, 6).value



        My.Settings.Save()
    End Sub

    Public Sub 读取数据表设置(SetSheet As Excel.Worksheet)
        Dim n As Integer = 12
        My.Settings.匹配数据表 = SetSheet.Cells(2 + n, 6).value


        My.Settings.数据表1级表头行号 = SetSheet.Cells(3 + n, 9).value
        My.Settings.数据表2级表头行号 = SetSheet.Cells(4 + n, 9).value
        My.Settings.数据表3级表头行号 = SetSheet.Cells(5 + n, 9).value


        My.Settings.数据表1级表头列号 = SetSheet.Cells(11 + n, 3).value
        My.Settings.数据表2级表头列号 = SetSheet.Cells(11 + n, 4).value
        My.Settings.数据表3级表头列号 = SetSheet.Cells(11 + n, 5).value


        My.Settings.数据表1级表头行名 = SetSheet.Cells(3 + n, 7).value
        My.Settings.数据表2级表头行名 = SetSheet.Cells(4 + n, 7).value
        My.Settings.数据表3级表头行名 = SetSheet.Cells(5 + n, 7).value

        My.Settings.数据表1级表头列名 = SetSheet.Cells(8 + n, 3).value
        My.Settings.数据表2级表头列名 = SetSheet.Cells(8 + n, 4).value
        My.Settings.数据表3级表头列名 = SetSheet.Cells(8 + n, 5).value


        My.Settings.结果页列序号_数据表1级表头行号 = MyRange(SetSheet, 3 + n, 10).Value
        My.Settings.结果页列序号_数据表2级表头行号 = MyRange(SetSheet, 4 + n, 10).Value
        My.Settings.结果页列序号_数据表3级表头行号 = MyRange(SetSheet, 5 + n, 10).Value

        My.Settings.结果页列序号_数据表1级表头列号 = MyRange(SetSheet, 12 + n, 3).Value
        My.Settings.结果页列序号_数据表2级表头列号 = MyRange(SetSheet, 12 + n, 4).Value
        My.Settings.结果页列序号_数据表3级表头列号 = MyRange(SetSheet, 12 + n, 5).Value


        My.Settings.数据表数据名称 = SetSheet.Cells(6 + n, 6).value

        My.Settings.Save()
    End Sub

    Public Function 分类计数(分类区域1 As Excel.Range,
                         分类区域2 As Excel.Range,
                         分类区域3 As Excel.Range,
                         分类区域4 As Excel.Range,
                         Optional 是否忽略空值 As Boolean = True,
                         Optional 编号非计数 As Boolean = True) As Integer


        Dim 分类条件1, 分类条件2, 分类条件3, 分类条件4, 总条件, 空值条件 As String
        Dim cell1, cell2, cell3, cell4 As Excel.Range
        Dim 填充区域首单元格, 填充区域 As Excel.Range
        Dim 插入区域 As Excel.Range
        Dim address1, address2 As String





        If 分类区域1 Is Nothing Then
            MsgBox("分类区域1 不能为空，请重新选择！")
            Return 0
        End If
        Dim sheet As Excel.Worksheet = 分类区域1.Worksheet
        '分类区域1 = app.Intersect(分类区域1, sheet.UsedRange)



        address1 = 分类区域1.Address

        'If 是否忽略空值 = True Then

        '    For Each cell As Excel.Range In 分类区域
        '        插入公式("=IF(" & cell.Address(False, False) & "="""","""",COUNTIF(" & 分类区域.Cells(1, 1).Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell.Address(False, True) & "," & cell.Address(False, False) & "))",
        '             sheet.Cells(cell.Row, cell.Column + 1))

        '    Next
        'Else

        '    For Each cell As Excel.Range In 分类区域
        '        插入公式("=COUNTIF(" & 分类区域.Cells(1, 1).Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell.Address(False, True) & "," & """=""&" & cell.Address(False, False) & ")",
        '             sheet.Cells(cell.Row, cell.Column + 1))

        '    Next
        'End If




        If 分类区域1 Is Nothing Then
            MsgBox("分类区域1 不能为空，请重新选择！")
            Return 0
        End If
        Dim ResultRow, ResultColumn As Integer
        ResultColumn = 1
        ResultRow = 1


        cell1 = 分类区域1.Cells(1, 1)
        空值条件 = "OR(" & cell1.Address(False, False) & "="""""
        If 编号非计数 = True Then
            分类条件1 = cell1.Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell1.Address(False, False) & "," & cell1.Address(False, False)
        Else
            分类条件1 = 分类区域1.Address(RowAbsolute:=True, ColumnAbsolute:=True) & "," & cell1.Address(False, False)
        End If


        If ResultRow < cell1.Row Then
            ResultRow = cell1.Row
        End If
        If ResultColumn < cell1.Column Then
            ResultColumn = cell1.Column
        End If


        If 分类区域2 IsNot Nothing Then
            cell2 = 分类区域2.Cells(1, 1)
            空值条件 &= "," & cell2.Address(False, False) & "="""""
            If 编号非计数 = True Then
                分类条件2 = "," & cell2.Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell2.Address(False, False) & "," & cell2.Address(False, False)
            Else
                分类条件2 = "," & 分类区域2.Address(RowAbsolute:=True, ColumnAbsolute:=True) & "," & cell2.Address(False, False)
            End If

            If ResultRow < cell2.Row Then
                ResultRow = cell2.Row
            End If
            If ResultColumn < cell2.Column Then
                ResultColumn = cell2.Column
            End If
        Else
            分类条件2 = ""
        End If

        If 分类区域3 IsNot Nothing Then
            cell3 = 分类区域3.Cells(1, 1)
            空值条件 &= "," & cell3.Address(False, False) & "="""""
            If 编号非计数 = True Then
                分类条件3 = "," & cell3.Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell3.Address(False, False) & "," & cell3.Address(False, False)
            Else
                分类条件3 = "," & 分类区域3.Address(RowAbsolute:=True, ColumnAbsolute:=True) & "," & cell3.Address(False, False)
            End If

            If ResultRow < cell3.Row Then
                ResultRow = cell3.Row
            End If
            If ResultColumn < cell3.Column Then
                ResultColumn = cell3.Column
            End If
        Else
            分类条件3 = ""
        End If

        If 分类区域4 IsNot Nothing Then
            cell4 = 分类区域4.Cells(1, 1)
            空值条件 &= "," & cell4.Address(False, False) & "="""""
            If 编号非计数 = True Then
                分类条件4 = "," & cell4.Address(RowAbsolute:=True, ColumnAbsolute:=True) & ":" & cell4.Address(False, False) & "," & cell4.Address(False, False)
            Else
                分类条件4 = "," & 分类区域4.Address(RowAbsolute:=True, ColumnAbsolute:=True) & "," & cell4.Address(False, False)
            End If

            If ResultRow < cell4.Row Then
                ResultRow = cell4.Row
            End If
            If ResultColumn < cell4.Column Then
                ResultColumn = cell4.Column
            End If
        Else
            分类条件4 = ""
        End If

        总条件 = 分类条件1 & 分类条件2 & 分类条件3 & 分类条件4

        空值条件 &= ")"






        If 分类区域1.Columns.Count = 1 Then
            插入区域 = 插入列(sheet, ResultColumn + 1, False)
            插入区域.NumberFormatLocal = "G/通用格式"
            填充区域首单元格 = sheet.Cells(cell1.Row, ResultColumn + 1)
            填充区域 = 填充区域首单元格.Resize(分类区域1.Rows.Count, 1)
            If 是否忽略空值 = True Then
                插入公式("=IF(" & 空值条件 & ","""",COUNTIFS(" & 总条件 & "))", 填充区域首单元格)
                填充区域.FillDown()
            Else
                插入公式("=COUNTIFS(" & 总条件 & ")", 填充区域首单元格)
                填充区域.FillDown()

            End If

        ElseIf 分类区域1.Rows.Count = 1 Then
            插入区域 = 插入行(sheet, ResultRow + 1, False)
            插入区域.NumberFormatLocal = "G/通用格式"
            填充区域首单元格 = sheet.Cells(ResultRow + 1, cell1.Column)
            填充区域 = 填充区域首单元格.Resize(1, 分类区域1.Columns.Count)
            If 是否忽略空值 = True Then
                插入公式("=IF(" & 空值条件 & ","""",COUNTIF(" & 总条件 & "))",
                            填充区域首单元格)
                填充区域.FillRight()
            Else
                插入公式("=COUNTIF(" & 总条件 & ")",
                            填充区域首单元格)
                填充区域.FillRight()

            End If
        Else
            填充区域 = Nothing
            MsgBox("数据区域 必需是 单行或单列 区域！")
        End If





        If 填充区域 IsNot Nothing Then
            设置填充色(填充区域, 新创建区域颜色)
        End If











    End Function
    ''' <summary>
    ''' 对指定的统计表“数据表”，按照指定列的不同值交叉分类统计数目，并生成统计表。
    ''' </summary>
    ''' <param name="数据表">要统计的数据表</param>
    ''' <param name="行分类字段列号">生成的统计表中，行标题所在原表中的列号</param>
    ''' <param name="列分类字段列号1">生成的统计表中，与行标题交叉统计的列标题1所在原表中的列号</param>
    ''' <param name="列分类字段列号2">生成的统计表中，与行标题交叉统计的列标题2所在原表中的列号</param>
    ''' <param name="列分类字段列号3">生成的统计表中，与行标题交叉统计的列标题3所在原表中的列号</param>
    ''' <param name="列标题所在行号">原数据表中列标题所在的行号</param>
    ''' <returns>返回：对指定的统计表“数据表”，按照指定列的不同值交叉分类统计数目所生成的统计表。</returns>
    Public Function 生成分类统计表(数据表 As Excel.Worksheet,
                            行分类字段列号 As Integer，
                            列分类字段列号1 As Integer,
                            Optional 列分类字段列号2 As Integer = 0,
                            Optional 列分类字段列号3 As Integer = 0,
                            Optional 列标题所在行号 As Integer = 1,
                            Optional 是否忽略空值 As Boolean = True)

        '设置为手动计算()



        Dim 分类统计表 As Excel.Worksheet = 新建工作表("分类统计表", True)
        If 列标题所在行号 >= 数据表.UsedRange.Rows.Count Then
            Exit Function
        End If
        Dim 行分类字段数据区域, 列分类字段数据区域1, 列分类字段数据区域2, 列分类字段数据区域3 As Excel.Range
        Dim 行分类序列, 列分类序列1, 列分类序列2, 列分类序列3 As New System.Collections.ArrayList

        If 行分类字段列号 > 0 Then
            行分类字段数据区域 = 数据表.Cells(列标题所在行号 + 1, 行分类字段列号).Resize(数据表.UsedRange.Rows.Count - 列标题所在行号, 1)
            行分类序列 = 枚举选区(行分类字段数据区域, True, True, 是否忽略空值, My.Settings.是否忽略首尾空白字符)
            '设置填充色(行分类字段数据区域, RGB(255, 200, 200))
        Else
            Exit Function
        End If

        If 列分类字段列号1 > 0 Then
            列分类字段数据区域1 = 数据表.Cells(列标题所在行号 + 1, 列分类字段列号1).resize(数据表.UsedRange.Rows.Count - 列标题所在行号, 1)
            列分类序列1 = 枚举选区(列分类字段数据区域1, True, True, 是否忽略空值, My.Settings.是否忽略首尾空白字符)
            '设置填充色(列分类字段数据区域1, RGB(255, 200, 200))
        Else
            Exit Function
        End If

        If 列分类字段列号2 > 0 Then
            列分类字段数据区域2 = 数据表.Cells(列标题所在行号 + 1, 列分类字段列号2).resize(数据表.UsedRange.Rows.Count - 列标题所在行号, 1)
            列分类序列2 = 枚举选区(列分类字段数据区域2, True, True, 是否忽略空值, My.Settings.是否忽略首尾空白字符)
            '设置填充色(列分类字段数据区域2, RGB(255, 200, 200))
        End If

        If 列分类字段列号3 > 0 Then
            列分类字段数据区域3 = 数据表.Cells(列标题所在行号 + 1, 列分类字段列号3).resize(数据表.UsedRange.Rows.Count - 列标题所在行号, 1)
            列分类序列3 = 枚举选区(列分类字段数据区域3, True, True, 是否忽略空值, My.Settings.是否忽略首尾空白字符)
            '设置填充色(列分类字段数据区域3, RGB(255, 200, 200))
        End If










        按列显示序列(行分类序列, 分类统计表, 2, 1)
        Dim 乘积序列 As New System.Collections.ArrayList




        For Each temp1 As String In 列分类序列1
            If 列分类序列2.Count = 0 Then
                Dim 分类 As New System.Collections.ArrayList
                分类.Add(temp1)
                乘积序列.Add(分类)
            Else
                For Each temp2 As String In 列分类序列2
                    If 列分类序列3.Count = 0 Then
                        Dim 分类 As New System.Collections.ArrayList
                        分类.Add(temp1)
                        分类.Add(temp2)
                        乘积序列.Add(分类)
                    Else
                        For Each temp3 As String In 列分类序列3
                            Dim 分类 As New System.Collections.ArrayList
                            分类.Add(temp1)
                            分类.Add(temp2)
                            分类.Add(temp3)
                            乘积序列.Add(分类)
                        Next
                    End If

                Next

            End If

        Next


        Dim lastColumn As Integer = 2
        Dim lastRow As Integer = 1
        For Each 分类 As System.Collections.ArrayList In 乘积序列
            Dim 列标题 As String = ""
            For Each str As String In 分类
                列标题 &= str & "."
            Next
            If 列标题.EndsWith(".") Then
                列标题 = 列标题.Substring(0, 列标题.Length - 1)
            End If
            分类统计表.Cells(1, lastColumn).value = 列标题
            lastColumn += 1
        Next
        分类统计表.Cells(1, lastColumn).value = "合计"

        'MsgBox("'" & 分类统计表.Name & "'!" & 分类统计表.Columns(行分类字段行号).Address)


        Dim 总条件， 行条件, 列条件, 列条件1, 列条件2, 列条件3 As String
        列条件 = ""

        For row As Integer = 0 To 行分类序列.Count - 1
            '行条件 = 获取列地址(数据表, 行分类字段列号) & ",""" & 行分类序列(row) & """"

            行条件 = 获取地址(MyRange(数据表, 列标题所在行号 + 1, 行分类字段列号, MAXROW, 行分类字段列号), False, False, True) &
                       ",""" & 行分类序列(row) & """"
            For column As Integer = 0 To 乘积序列.Count - 1
                Dim 列号 As Integer
                For n As Integer = 0 To 乘积序列(column).Count - 1
                    If n = 0 Then
                        列号 = 列分类字段列号1
                    ElseIf n = 1 Then
                        列号 = 列分类字段列号2
                    ElseIf n = 2 Then
                        列号 = 列分类字段列号3
                    End If
                    '列条件 &= 获取列地址(数据表, 列号) & ",""" & 乘积序列(column)(n) & """" & ","
                    列条件 &= 获取地址(MyRange(数据表, 列标题所在行号 + 1, 列号, MAXROW, 列号), False, False, True) &
                        ",""" & 乘积序列(column)(n) & """" & ","
                Next

                If 列条件.EndsWith(",") Then
                    列条件 = 列条件.Substring(0, 列条件.Length - 1)
                End If

                总条件 = 行条件 & "," & 列条件
                '插入公式("=COUNTIFS(" & 总条件 & ")", 分类统计表.Cells(row + 2, column + 2))

                分类统计表.Cells(row + 2, column + 2).Formula = "=COUNTIFS(" & 总条件 & ")"

                lastRow = row + 2
                列条件 = ""
            Next
            行条件 = ""
        Next


        Dim 左上单元格 As Excel.Range = 分类统计表.Cells(2, 2)
        分类统计表.Cells(行分类序列.Count + 2, 1).value = "合计"
        lastRow += 1
        插入公式("=SUM(" & 左上单元格.Resize(1, lastColumn - 2).Address(False, False) & ")", 分类统计表.Cells(2, lastColumn))
        拖拽填充(分类统计表.Cells(2, lastColumn).Resize(lastRow - 2, 1))

        插入公式("=SUM(" & 左上单元格.Resize(lastRow - 2, 1).Address(False, False) & ")", 分类统计表.Cells(lastRow, 2))
        拖拽填充(分类统计表.Cells(lastRow, 2).Resize(1, lastColumn - 1))




        自动列宽(分类统计表)
        设置为表头样式(分类统计表, 1, 1)
        居中(分类统计表.Cells)
        '设置为自动计算()

    End Function

    Public Function 获取列地址(sheet As Excel.Worksheet,
                          column As Integer,
                          Optional 是否带有表名 As Boolean = True,
                          Optional 是否绝对地址 As Boolean = True)
        Dim 表名 As String = ""
        If 是否带有表名 = True Then
            表名 = "'" & sheet.Name & "'!"
        Else
            表名 = ""
        End If
        If 是否绝对地址 = True Then
            Return 表名 & sheet.Columns(column).address
        Else
            Return 表名 & sheet.Columns(column).address(False, False)
        End If
    End Function

    Public Function 获取地址(range As Excel.Range,
                           Optional 是否锁定行 As Boolean = False,
                           Optional 是否锁定列 As Boolean = False，
                           Optional 是否锁定表 As Boolean = False) As String


        If range Is Nothing Then
            Return ""
        Else
            Dim sheetName As String = ""
            If 是否锁定表 = True Then
                sheetName = "'" & range.Worksheet.Name & "'!"
            End If
            Return sheetName & range.Address(是否锁定行, 是否锁定列)
        End If

    End Function
    Public Function 按行显示序列(list As System.Collections.ArrayList, sheet As Excel.Worksheet, 行 As Integer, 列 As Integer)

        For Each temp As String In list
            sheet.Cells(行, 列).value = temp
            列 += 1
        Next

    End Function
    Public Function 按列显示序列(list As System.Collections.ArrayList, sheet As Excel.Worksheet, 行 As Integer, 列 As Integer)

        For Each temp As String In list
            sheet.Cells(行, 列).value = temp
            行 += 1
        Next

    End Function

    Public Function 取数字(区域 As Excel.Range) As Integer

        Dim sheet As Excel.Worksheet = 区域.Worksheet
        区域 = app.Intersect(区域, sheet.UsedRange)

        Dim address1, address2 As String
        If 区域 IsNot Nothing Then
            If 区域.Columns.Count = 1 Then
                Dim range As Excel.Range = 插入列(sheet, 区域.Cells(1, 1).column + 1)
                address1 = 区域.Cells(1, 1).Address(False, False)

                插入公式("=-Lookup(0, -Mid(SUBSTITUTE(" & address1 & ", ""e"", ""a""), MIN(FIND(ROW($1:$10)-1,SUBSTITUTE(" & address1 & ",""e"",""a"")&1/17)),ROW($1:$10000)))",
                     sheet.Cells(区域.Cells(1, 1).Row, 区域.Cells(1, 1).Column + 1), True)
                拖拽填充(MyRange(sheet, 区域.Cells(1, 1).Row, 区域.Cells(1, 1).Column + 1, 区域.Cells(1, 1).Row + 区域.Rows.Count - 1, 区域.Cells(1, 1).Column + 1))

            ElseIf 区域.Rows.Count = 1 Then
                Dim range As Excel.Range = 插入行(sheet, 区域.Cells(1, 1).row + 1)
                address1 = 区域.Cells(1, 1).Address(False, False)

                插入公式("=-Lookup(0, -Mid(SUBSTITUTE(" & address1 & ", ""e"", ""a""), MIN(FIND(ROW($1:$10)-1,SUBSTITUTE(" & address1 & ",""e"",""a"")&1/17)),ROW($1:$10000)))",
                     sheet.Cells(区域.Cells(1, 1).Row + 1, 区域.Cells(1, 1).Column), True)
                拖拽填充(MyRange(sheet, 区域.Cells(1, 1).Row + 1, 区域.Cells(1, 1).Column, 区域.Cells(1, 1).Row + 1, 区域.Cells(1, 1).Column + 区域.Columns.Count - 1))
            Else
                MsgBox("必须选择单行或单列区域" & vbCrLf & "请重新选择！")
            End If
        End If



        'For Each cell As Excel.Range In 区域
        '    address1 = cell.Address(False, False)
        '    插入公式("=-Lookup(0, -Mid(SUBSTITUTE(" & address1 & ", ""e"", ""a""), MIN(FIND(ROW($1:$10)-1,SUBSTITUTE(" & address1 & ",""e"",""a"")&1/17)),ROW($1:$10000)))",
        '         sheet.Cells(cell.Row, cell.Column + 1), True)

        'Next

    End Function


    Public Function 字符串逆序(区域 As Excel.Range) As Integer

        Dim sheet As Excel.Worksheet = 区域.Worksheet
        区域 = app.Intersect(区域, sheet.UsedRange)


        Dim temp As String
        If 区域 IsNot Nothing Then
            If 区域.Columns.Count = 1 Then
                Dim range As Excel.Range = 插入列(sheet, 区域.Cells(1, 1).column + 1)
                设置单元格格式(range, "文本")
                For Each cell As Excel.Range In 区域
                    If cell.Value <> Nothing Then
                        temp = ""
                        For Each c As Char In cell.Value.ToString.Reverse
                            temp &= c
                        Next
                        sheet.Cells(cell.Row, cell.Column + 1).value = temp
                    End If

                Next

            ElseIf 区域.Rows.Count = 1 Then
                Dim range As Excel.Range = 插入行(sheet, 区域.Cells(1, 1).row + 1)

                设置单元格格式(range, "文本")
                For Each cell As Excel.Range In 区域
                    If cell.Value <> Nothing Then
                        temp = ""
                        For Each c As Char In cell.Value.ToString.Reverse
                            temp &= c
                        Next
                        sheet.Cells(cell.Row + 1, cell.Column).value = temp
                    End If

                Next


            Else
                MsgBox("必须选择单行或单列区域" & vbCrLf & "请重新选择！")
            End If
        End If




    End Function
    Public Function 字符串第n次匹配的位置(区域 As Excel.Range)
        Dim sheet As Excel.Worksheet = 区域.Worksheet
        区域 = app.Intersect(区域, sheet.UsedRange)

        Dim 要匹配的字符串 As String = InputBox("请输入要匹配的字符串", "输入")
        Dim input_str As String = InputBox("请输入要查找第几次匹配", "输入次数")
        Dim 第n次匹配 As Integer
        If IsNumeric(input_str) Then
            第n次匹配 = CInt(input_str)
        Else
            MsgBox("输入的不是数字，请重设置！")
            Exit Function
        End If


        Dim range, 填充区域 As Excel.Range
        Dim 源字符串地址 As String



        If 区域 IsNot Nothing Then
            If 区域.Columns.Count = 1 Then
                range = 插入列(sheet, 区域.Cells(1, 1).column + 1)
                填充区域 = range.Cells(1, 1).resize(区域.Rows.Count， 1)
                源字符串地址 = 获取地址(区域.Cells(1, 1), False, True)


                'For Each cell As Excel.Range In 区域
                '    If cell.Value <> Nothing Then
                '        temp = ""
                '        For Each c As Char In cell.Value.ToString.Reverse
                '            temp &= c
                '        Next
                '        sheet.Cells(cell.Row, cell.Column + 1).value = temp
                '    End If

                'Next

            ElseIf 区域.Rows.Count = 1 Then
                range = 插入行(sheet, 区域.Cells(1, 1).row + 1)
                填充区域 = range.Cells(1, 1).resize(1， 区域.Columns.Count)
                源字符串地址 = 获取地址(区域.Cells(1, 1), True, False)
                '设置单元格格式(range, "通用")
                'For Each cell As Excel.Range In 区域
                '    If cell.Value <> Nothing Then
                '        temp = ""
                '        For Each c As Char In cell.Value.ToString.Reverse
                '            temp &= c
                '        Next
                '        sheet.Cells(cell.Row + 1, cell.Column).value = temp
                '    End If

                'Next


            Else
                MsgBox("必须选择单行或单列区域" & vbCrLf & "请重新选择！")
                Return Nothing
            End If
            设置单元格格式(range, "通用")
            ' =FIND(CHAR(160),SUBSTITUTE(A1,"a",CHAR(160),B1))

            插入公式("=FIND(CHAR(160),SUBSTITUTE(" & 源字符串地址 & ",""" & 要匹配的字符串 & """,CHAR(160)," & 第n次匹配 & "))", range.Cells(1, 1), False)
            拖拽填充(填充区域)
        End If


















    End Function



    Public Function 字符串拆分(区域 As Excel.Range) As Integer

        Dim sheet As Excel.Worksheet = 区域.Worksheet
        区域 = app.Intersect(区域, sheet.UsedRange)

        Dim 分隔符 As String = InputBox("请输入分隔符", "输入")
        Dim charArray As Char()
        ReDim charArray(分隔符.Length - 1)
        For t = 0 To 分隔符.Length - 1
            charArray(t) = 分隔符(t)
        Next




        Dim endColumn As Integer = 获取结束列号(sheet)
        Dim currentColumn As Integer = endColumn + 1
        Dim endRow As Integer = 获取结束行号(sheet)
        Dim currentRow As Integer = endRow + 1


        If 区域 IsNot Nothing Then
            If 区域.Columns.Count = 1 Then
                'Dim range As Excel.Range = 插入列(sheet, 区域.Cells(1, 1).column + 1)
                '设置单元格格式(range, "文本")
                For Each cell As Excel.Range In 区域
                    If cell.Value <> Nothing Then
                        currentColumn = endColumn + 1
                        For Each str As String In cell.Value.ToString.Split(charArray)
                            sheet.Cells(cell.Row, currentColumn).value = str
                            currentColumn += 1
                        Next

                    End If

                Next

            ElseIf 区域.Rows.Count = 1 Then
                'Dim range As Excel.Range = 插入行(sheet, 区域.Cells(1, 1).row + 1)

                '设置单元格格式(range, "文本")
                For Each cell As Excel.Range In 区域
                    If cell.Value <> Nothing Then
                        currentRow = endRow + 1
                        For Each str As String In cell.Value.ToString.Split(charArray)
                            sheet.Cells(currentRow, cell.Column).value = str
                            currentRow += 1
                        Next

                    End If

                Next


            Else
                MsgBox("必须选择单行或单列区域" & vbCrLf & "请重新选择！")
            End If
        End If




    End Function

    Public Function 排名(数据区域 As Excel.Range,
                       Optional 是否为升序 As Boolean = True,
                       Optional 同数据处理方式 As Integer = 1,
                       Optional 是否忽略空值 As Boolean = True,
                       Optional 分类区域1 As Excel.Range = Nothing,
                       Optional 分类区域2 As Excel.Range = Nothing,
                       Optional 分类区域3 As Excel.Range = Nothing) As Integer

        If 数据区域 Is Nothing Then
            MsgBox("数据区域 不能为空，请重新选择！")
            Return 0
        End If
        If 数据区域.Columns.Count <> 1 And 数据区域.Rows.Count <> 1 Then
            MsgBox("数据区域 必须是 单行或单列区域，请重新选择！")
            Return 0
        End If

        Dim sheet As Excel.Worksheet = 数据区域.Worksheet
        Dim 分类地址1, 分类地址2, 分类地址3, 数据地址 As String
        Dim 分类列号1, 分类列号2， 分类列号3, 分类行号1, 分类行号2， 分类行号3 As Integer

        分类地址1 = ""
        分类地址2 = ""
        分类地址3 = ""
        If 分类区域1 IsNot Nothing Then
            分类地址1 = 分类区域1.Address(True, True)
            'If 分类区域1.Columns.Count <> 1 Then
            '    MsgBox("分类区域1 必须是单列区域，请重新选择！")
            '    Return 0
            'End If
            分类列号1 = 分类区域1.Column
            分类行号1 = 分类区域1.Row
        End If

        If 分类区域2 IsNot Nothing Then
            分类地址2 = 分类区域2.Address(True, True)
            'If 分类区域2.Columns.Count <> 1 Then
            '    MsgBox("分类区域2 必须是单列区域，请重新选择！")
            '    Return 0
            'End If
            分类列号2 = 分类区域2.Column
            分类行号2 = 分类区域2.Row
        End If

        If 分类区域3 IsNot Nothing Then
            分类地址3 = 分类区域3.Address(True, True)
            'If 分类区域3.Columns.Count <> 1 Then
            '    MsgBox("分类区域3 必须是单列区域，请重新选择！")
            '    Return 0
            'End If
            分类列号3 = 分类区域3.Column
            分类行号3 = 分类区域3.Row
        End If
        'If 数据区域.Columns.Count <> 1 Then
        '    MsgBox("数据区域 必须是单列区域，请重新选择！")
        '    Return 0
        'End If

        数据地址 = 数据区域.Address(True, True)



        Dim 偏移行号 As Integer
        偏移行号 = 0
        Dim 不等号, 等号 As String
        Dim 排序方式符号 As String = ">"
        If 是否为升序 = True Then
            不等号 = "<"
        Else
            不等号 = ">"
        End If

        If 同数据处理方式 = 1 Then
            等号 = ""
            偏移行号 = 1
        ElseIf 同数据处理方式 = -1 Then
            等号 = "="
        ElseIf 同数据处理方式 = 0 Then
            等号 = ""
        Else
            等号 = ""
        End If
        排序方式符号 = 不等号 & 等号



        Dim 分类1单元格地址, 分类2单元格地址, 分类3单元格地址, 数据单元格地址 As String

        Dim 分类条件1, 分类条件2, 分类条件3, 分类总条件, 同数据处理条件, 空值处理 As String
        分类条件1 = ""
        分类条件2 = ""
        分类条件3 = ""
        同数据处理条件 = ""
        分类总条件 = 分类条件1 & 分类条件2 & 分类条件3



        Dim Cell As Excel.Range = 数据区域.Cells(1, 1)
        数据单元格地址 = Cell.Address(False, False)
        If 分类区域1 Is Nothing Then
            分类条件1 = ""
            分类条件2 = ""
            分类条件3 = ""
        ElseIf 分类区域2 Is Nothing Then
            If 数据区域.Columns.Count = 1 Then
                分类1单元格地址 = sheet.Cells(Cell.Row, 分类列号1).Address(False, False)
            ElseIf 数据区域.Rows.Count = 1 Then
                分类1单元格地址 = sheet.Cells(分类行号1, Cell.Column).Address(False, False)
            End If

            分类条件1 = "(" & 分类地址1 & "=" & 分类1单元格地址 & ")*"
            分类条件2 = ""
            分类条件3 = ""
        ElseIf 分类区域3 Is Nothing Then
            If 数据区域.Columns.Count = 1 Then
                分类1单元格地址 = sheet.Cells(Cell.Row, 分类列号1).Address(False, False)
                分类2单元格地址 = sheet.Cells(Cell.Row, 分类列号2).Address(False, False)
            ElseIf 数据区域.Rows.Count = 1 Then
                分类1单元格地址 = sheet.Cells(分类行号1, Cell.Column).Address(False, False)
                分类2单元格地址 = sheet.Cells(分类行号2, Cell.Column).Address(False, False)
            End If

            分类条件1 = "(" & 分类地址1 & "=" & 分类1单元格地址 & ")*"
            分类条件2 = "(" & 分类地址2 & "=" & 分类2单元格地址 & ")*"
            分类条件3 = ""
        Else
            If 数据区域.Columns.Count = 1 Then
                分类1单元格地址 = sheet.Cells(Cell.Row, 分类列号1).Address(False, False)
                分类2单元格地址 = sheet.Cells(Cell.Row, 分类列号2).Address(False, False)
                分类3单元格地址 = sheet.Cells(Cell.Row, 分类列号3).Address(False, False)
            ElseIf 数据区域.Rows.Count = 1 Then
                分类1单元格地址 = sheet.Cells(分类行号1, Cell.Column).Address(False, False)
                分类2单元格地址 = sheet.Cells(分类行号2, Cell.Column).Address(False, False)
                分类3单元格地址 = sheet.Cells(分类行号3, Cell.Column).Address(False, False)
            End If

            分类条件1 = "(" & 分类地址1 & "=" & 分类1单元格地址 & ")*"
            分类条件2 = "(" & 分类地址2 & "=" & 分类2单元格地址 & ")*"
            分类条件3 = "(" & 分类地址3 & "=" & 分类3单元格地址 & ")*"

        End If

        If 是否忽略空值 = True Then
            空值处理 = "(" & 数据地址 & "<>"""")*"
        Else
            空值处理 = ""
        End If


        分类总条件 = 分类条件1 & 分类条件2 & 分类条件3 & 空值处理

        '=SUMPRODUCT(($A$2:$A$25=A2)*($B$2:$B$25>B2))+1

        If 同数据处理方式 = 1 Or 同数据处理方式 = -1 Then
            同数据处理条件 = ""
        ElseIf 同数据处理方式 = 0 Then
            同数据处理条件 = "+" & 分类总条件 & "(ROW(" & 数据地址 & ")<=ROW())*(" & 数据地址 & "=" & 数据单元格地址 & ")"
        End If







        Dim range, 填充区域 As Excel.Range
        If 数据区域.Columns.Count = 1 Then
            range = 插入列(sheet, Cell.Column + 1)
            range.NumberFormatLocal = "G/通用格式"
            If 是否忽略空值 = True Then
                插入公式("=IF(" & 数据单元格地址 & "="""","""",SUMPRODUCT(1*" & 分类总条件 & "(" & 数据地址 & 排序方式符号 & 数据单元格地址 & ")" &
                 同数据处理条件 & ")+" & 偏移行号 & ")",
                 sheet.Cells(Cell.Row, Cell.Column + 1))
            Else
                插入公式("=SUMPRODUCT(1*" & 分类总条件 & "(" & 数据地址 & 排序方式符号 & 数据单元格地址 & ")" &
                 同数据处理条件 & ")+" & 偏移行号,
                 sheet.Cells(Cell.Row, Cell.Column + 1))
            End If

            填充区域 = MyRange(sheet, Cell.Row, Cell.Column + 1, Cell.Row + 数据区域.Rows.Count - 1, Cell.Column + 1)
            填充区域.FillDown()
        ElseIf 数据区域.Rows.Count = 1 Then
            range = 插入行(sheet, Cell.Row + 1)
            range.NumberFormatLocal = "G/通用格式"
            If 是否忽略空值 = True Then
                插入公式("=IF(" & 数据单元格地址 & "="""","""",SUMPRODUCT(1*" & 分类总条件 & "(" & 数据地址 & 排序方式符号 & 数据单元格地址 & ")" &
                 同数据处理条件 & ")+" & 偏移行号 & ")",
                 sheet.Cells(Cell.Row + 1, Cell.Column))
            Else
                插入公式("=SUMPRODUCT(1*" & 分类总条件 & "(" & 数据地址 & 排序方式符号 & 数据单元格地址 & ")" &
                 同数据处理条件 & ")+" & 偏移行号,
                 sheet.Cells(Cell.Row + 1, Cell.Column))
            End If

            填充区域 = MyRange(sheet, Cell.Row + 1, Cell.Column, Cell.Row + 1, Cell.Column + 数据区域.Columns.Count - 1)
            填充区域.FillRight()
        Else
            MsgBox("数据区域只能是单行或单列的区域！")
            Return 0
        End If


        'If range IsNot Nothing Then
        '    设置填充色(range, RGB(255, 230, 230))
        'End If

        居中(sheet.UsedRange)
        Return 数据区域.Count
    End Function
    Public Sub 设置单元格格式(要设置的区域 As Excel.Range, Optional 格式 As String = "通用")
        If 格式.Contains("通用") Then
            要设置的区域.NumberFormatLocal = "G/通用格式"
        ElseIf 格式.Contains("文本") Then
            要设置的区域.NumberFormatLocal = "@"
        End If

    End Sub

    ''' <summary>
    ''' 在指定表格的指定列的左边插入一列
    ''' </summary>
    ''' <param name="sheet">指定表格</param>
    ''' <param name="NumOfColumn">指定列的列号</param>
    ''' <param name="是否设置背景色">是否设置背景颜色</param>
    ''' <param name="单元格格式">设置单元格格式  常规-"G/通用格式";文本-"@"</param>
    ''' <returns>返回所插入的列区域</returns>
    Public Function 插入列(sheet As Excel.Worksheet,
                        NumOfColumn As Integer,
                        Optional 是否设置背景色 As Boolean = True,
                        Optional 单元格格式 As String = "G/通用格式") As Excel.Range

        sheet.Columns.Item(NumOfColumn).Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        Dim 列 As Excel.Range = sheet.Columns.Item(NumOfColumn)
        列.NumberFormatLocal = 单元格格式
        If 是否设置背景色 = True Then
            Dim n As Integer = 获取结束行号(sheet)
            设置填充色(列.Cells(1, 1).resize(n, 1), 新创建区域颜色)
        End If
        Return 列
    End Function

    Public Function 插入行(sheet As Excel.Worksheet, NumOfRow As Integer, Optional 是否设置背景色 As Boolean = True) As Excel.Range
        sheet.Rows.Item(NumOfRow).Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
        If 是否设置背景色 = True Then
            设置填充色(sheet.Rows.Item(NumOfRow), 新创建区域颜色)
        End If
        Return sheet.Rows.Item(NumOfRow)
    End Function


    ''' <summary>
    ''' 指定区域的值按照指定的顺序装换成序列（序列的中的值可以是相同的，不是枚举）
    ''' </summary>
    ''' <param name="range">指定的区域</param>
    ''' <param name="是否先行后列">指定是否是先行后列</param>
    ''' <returns>返回转换后的序列</returns>

    Public Function 区域转序列(range As Excel.Range, 是否先行后列 As Boolean) As Collections.ArrayList
        Dim result As New Collections.ArrayList
        If 是否先行后列 = True Then
            For Each cell As Excel.Range In range
                result.Add(cell.Value)
            Next
        Else
            For Each column As Excel.Range In range.Columns
                For Each cell As Excel.Range In column.Cells
                    result.Add(cell.Value)
                Next
            Next
        End If
        Return result
    End Function
    Public Function 获取缓冲表() As Excel.Worksheet
        Dim sheet As Excel.Worksheet
        If 是否存在工作表(缓冲表名) Then
            sheet = app.Sheets(缓冲表名)
            sheet.Cells.Delete()
        Else
            sheet = 新建工作表(缓冲表名)
        End If
        'sheet.Visible = False

        Return sheet
    End Function

    Public Function CustomTaskPanesIndex(name As String)
        Dim i As Integer
        For i = 0 To Globals.ThisAddIn.CustomTaskPanes.Count - 1
            If Globals.ThisAddIn.CustomTaskPanes.Item(i).Title = name Then
                Return i
            End If
        Next
        Return -1
    End Function

    Private Sub CustomPaneVisibleChanged(sender As Object， e As System.EventArgs)
        Dim temp As Microsoft.Office.Tools.CustomTaskPane = sender
        If temp.Visible = False Then '用户关闭任务窗格
            'MsgBox("关闭")
            Globals.ThisAddIn.CustomTaskPanes.Remove(temp)
        Else '用户打开任务窗格
            'MsgBox("打开")
        End If

    End Sub








    Public Function 添加或显示功能控件(Control1 As Windows.Forms.Control,
                              Optional 标题 As String = "无标题",
                              Optional 是否填充容器 As Boolean = False) As Microsoft.Office.Tools.CustomTaskPane
        Dim index As Integer

        Dim width As Integer = Control1.Width
        Dim height As Integer = Control1.Height



        If My.Settings.是否启用任务窗格 = True Then
            'index = CustomTaskPanesIndex(标题) '查找是否已经有同标题的任务窗格

            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Control1, 标题)
            AddHandler taskPane.VisibleChanged, AddressOf CustomPaneVisibleChanged

            taskPane.Width = width

            taskPane.Visible = True
            Return taskPane
            'If index < 0 Then

            'Else
            '    Globals.ThisAddIn.CustomTaskPanes.Item(index).Visible = True
            'End If



        Else


            index = TabIndex(标题)
            If index < 0 Then
                MyForm1.TabControl1.TabPages.Add(标题, 标题)
                Dim Tab As Windows.Forms.TabPage = MyForm1.TabControl1.TabPages.Item(MyForm1.TabControl1.TabPages.Count - 1)




                Tab.Controls.Add(Control1)

                'Tab.BackColor = Drawing.Color.FromArgb(230, 230, 255)
                'Tab.Font = New Drawing.Font("微软雅黑", 12)
                MyForm1.TabControl1.SelectedIndex = MyForm1.TabControl1.TabPages.Count - 1
                Control1.Top = 0
                Control1.Left = 0
                MyForm1.Left = My.Computer.Screen.WorkingArea.Width - MyForm1.Width
                MyForm1.更新大小()
                If 是否填充容器 = True Then
                    Control1.Dock = Windows.Forms.DockStyle.Fill
                End If

                MyForm1.Show()
            Else
                MyForm1.TabControl1.SelectedIndex = index
            End If

            Return Nothing
        End If




    End Function

    Public Sub 显示插件设置()
        Dim 插件设置 As New 插件设置控件
        添加或显示功能控件(插件设置, "插件设置")
    End Sub

    Public Sub 插件设置()
        当前选区 = app.Selection
        Dim 设置标题 As String = "插件系统设置"
        Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        If SetSheet Is Nothing Then '还没完成设置，以下开始设置
            SetSheet = 打开设置页(设置标题)

            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码

            Dim tempRange As Excel.Range
            tempRange = MyRange(SetSheet, 2, 2, 2, 5)
            tempRange.Merge()
            tempRange.Value = "表头行"



            tempRange = MyRange(SetSheet, 3, 1, 8, 1)
            tempRange.Merge()
            tempRange.Value = "表头列"



            For i = 1 To My.Settings.循环色序列.Count
                tempRange = MyRange(SetSheet, 3, 1 + i, 8, 1 + i)
                tempRange.Merge()
                tempRange.Value = "循环色" & i
                tempRange.Interior.Color = My.Settings.循环色序列(i - 1)
            Next





            设置为表头样式(SetSheet, 2, 1)

            tempRange = MyRange(SetSheet, 11, 1)
            tempRange.Value = "字符串分隔符"
            设置字体(tempRange, "微软雅黑", 11, RGB(0, 0, 0))

            tempRange = MyRange(SetSheet, 11, 2)
            tempRange.Value = My.Settings.字符串分隔符
            设置字体(tempRange, "微软雅黑", 11, RGB(255, 0, 0))







            tempRange = MyRange(SetSheet, 12, 1)
            tempRange.Value = "枚举时是否剔除重复值"
            设置字体(tempRange, "微软雅黑", 11, RGB(0, 0, 0))

            tempRange = MyRange(SetSheet, 12, 2)
            tempRange.Value = My.Settings.是否剔除重复值
            设置字体(tempRange, "枚举时是否剔除重复值", 11, RGB(255, 0, 0))


            With tempRange.Validation
                .Delete()
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                     AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                     Operator:=Excel.XlFormatConditionOperator.xlBetween,
                     Formula1:="True,False")
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl  'xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With












            tempRange = MyRange(SetSheet, 13, 1)
            tempRange.Value = "枚举时是否忽略空值"
            设置字体(tempRange, "微软雅黑", 11, RGB(0, 0, 0))

            tempRange = MyRange(SetSheet, 13, 2)
            tempRange.Value = My.Settings.是否忽略空值
            设置字体(tempRange, "枚举时是否忽略空值", 11, RGB(255, 0, 0))

            With tempRange.Validation
                .Delete()
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                     AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                     Operator:=Excel.XlFormatConditionOperator.xlBetween,
                     Formula1:="True,False")
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl  'xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With








            tempRange = MyRange(SetSheet, 14, 1)
            tempRange.Value = "枚举时是否忽略首尾空白字符"
            设置字体(tempRange, "微软雅黑", 11, RGB(0, 0, 0))

            tempRange = MyRange(SetSheet, 14, 2)
            tempRange.Value = My.Settings.是否忽略首尾空白字符
            设置字体(tempRange, "枚举时是否忽略首尾空白字符", 11, RGB(255, 0, 0))

            With tempRange.Validation
                .Delete()
                .Add(Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                     AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                     Operator:=Excel.XlFormatConditionOperator.xlBetween,
                     Formula1:="True,False")
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl  'xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With



            SetSheet.Cells.EntireColumn.AutoFit() '自动列宽
            Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
            添加按钮(按钮位置, "保存设置", "Start")
        End If
    End Sub






    Public Sub 颜色代码()
        当前选区 = app.Selection
        Dim 设置标题 As String = "我的设置标题"
        Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        If SetSheet Is Nothing Then '还没完成设置，以下开始设置
            SetSheet = 打开设置页(设置标题)
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码

            Dim tempRange As Excel.Range
            Dim R, G, B As Integer

            For n As Integer = 1 To 10
                For i As Integer = 1 To 15
                    tempRange = MyRange(SetSheet, i + 1, 2 * n - 1)
                    tempRange.Value = i
                    设置字体(tempRange, "微软雅黑", 16)
                    R = Int(255 / 10 * n) 'Int(Rnd() * 255)
                    G = Int(255 / 15 * i)  'Int(Rnd() * 255)
                    B = 255 - Int(255 / 25 * (i + n))
                    设置填充色(tempRange, RGB(R, G, B))
                    MyRange(SetSheet, i + 1, 2 * n).Value = RGB(R, G, B)
                Next

            Next



            Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
            添加按钮(按钮位置, "开始执行", "Start")
            SetSheet.Cells.EntireColumn.AutoFit() '自动列宽

        End If

    End Sub




    Public Function TabIndex(Name As String) As Integer

        For i = 0 To MyForm1.TabControl1.TabPages.Count - 1
            If MyForm1.TabControl1.TabPages.Item(i).Name = Name Then
                Return i
            End If
        Next
        Return -1
    End Function
    Public Function 添加控件到工作表(sheet As Excel.Worksheet,
                             MyControl As Windows.Forms.Control,
                             Range As Excel.Range,
                             Optional name As String = "MyControl")
        Dim worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        worksheet.Controls.AddControl(MyControl, Range, name & Now.Ticks)
    End Function
    Public Function 获取单元格个数(区域 As Excel.Range) As Double
        Dim len, a, b As Double

        Try
            Return 区域.Count
        Catch ex As Exception
            a = 区域.Rows.Count
            b = 区域.Columns.Count
            len = (a * b)
            Return len
        End Try
    End Function



    ''' <summary>
    ''' 获取UsedRange的最后一个单元格
    ''' </summary>
    ''' <param name="Sheet">指定的表</param>
    ''' <returns>返回自定标的UsedRange区域的最后一个单元格</returns>
    Public Function 获取结束单元格(Sheet As Excel.Worksheet) As Excel.Range
        Return Sheet.UsedRange(Sheet.UsedRange.Rows.Count, Sheet.UsedRange.Columns.Count)
    End Function

    Public Function 获取结束单元格(区域 As Excel.Range) As Excel.Range
        Return 区域(区域.Rows.Count, 区域.Columns.Count)
    End Function



    Public Function 获取结束行号(Sheet As Excel.Worksheet) As Integer
        Return 获取结束单元格(Sheet).Row
    End Function

    Public Function 获取结束行号(区域 As Excel.Range) As Integer
        Return 获取结束单元格(区域).Row
    End Function
    Public Function 获取结束列号(Sheet As Excel.Worksheet) As Integer
        Return 获取结束单元格(Sheet).Column
    End Function

    Public Function 获取结束列号(区域 As Excel.Range) As Integer
        Return 获取结束单元格(区域).Column
    End Function
    ''' <summary>
    ''' 获取一列中最后一个非空单元格的行号
    ''' </summary>
    ''' <param name="Sheet">指定表</param>
    ''' <param name="列号">指定列</param>
    ''' <returns>返回一列中非空最后一个非空单元格的行号</returns>
    Public Function 获取非空列数据尾行号(Sheet As Excel.Worksheet, 列号 As Integer) As Integer

        Dim 尾行号 As Integer = 获取结束行号(Sheet)
        While Sheet.Cells(尾行号, 列号).value Is Nothing And 尾行号 <> 1
            尾行号 -= 1
        End While
        Return 尾行号
    End Function

    Public Function 获取非空列数据首行号(Sheet As Excel.Worksheet, 列号 As Integer) As Integer
        Dim 尾行号 As Integer = 获取结束行号(Sheet)
        Dim 首行号 As Integer = 1
        While Sheet.Cells(首行号, 列号).value Is Nothing And 首行号 < 尾行号
            首行号 += 1
        End While
        Return 首行号
    End Function




    ''' <summary>
    ''' 获取一行中最后一个非空单元格的列号
    ''' </summary>
    ''' <param name="Sheet"></param>
    ''' <param name="行号"></param>
    ''' <returns>返回 获取一行中最后一个非空单元格的列号</returns>
    Public Function 获取非空行数据尾列号(Sheet As Excel.Worksheet, 行号 As Integer) As Integer

        Dim 尾列号 As Integer = 获取结束列号(Sheet)
        While Sheet.Cells(行号, 尾列号).value Is Nothing And 尾列号 <> 1
            尾列号 -= 1
        End While
        Return 尾列号
    End Function


    Public Function 获取非空行数据首列号(Sheet As Excel.Worksheet, 行号 As Integer) As Integer
        Dim 尾列号 As Integer = 获取结束列号(Sheet)
        Dim 首列号 As Integer = 1
        While Sheet.Cells(行号， 首列号).value Is Nothing And 首列号 < 尾列号
            首列号 += 1
        End While
        Return 首列号
    End Function

    Public Function 添加探照灯效果() As Boolean
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range
        Dim selectRange As Excel.Range = app.Selection

        设置为手动计算()


        Dim rowMin, rowMax, columnMin, columnMax As Integer
        rowMin = selectRange.Row
        rowMax = selectRange.Row + selectRange.Rows.Count - 1
        columnMin = selectRange.Column
        columnMax = selectRange.Column + selectRange.Columns.Count - 1

        'range = app.Union(sheet.Cells(rowMin, 1).resize(selectRange.Rows.Count, 16384),
        '                  sheet.Cells(1, columnMin).resize(1048576, selectRange.Columns.Count))
        range = sheet.Cells

        range.FormatConditions.Add(Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression,
                                   Formula1:="=OR(AND(ROW()>=" & rowMin & ",ROW()<=" & rowMax & "),AND(COLUMN()>=" & columnMin & ",COLUMN()<=" & columnMax & "))")

        '''' =OR(AND(COLUMN()>=5,COLUMN()<=9),AND(ROW()>=12,ROW()<=22))
        '''' "=OR(CELL(""col"")=COLUMN(),CELL(""row"")=ROW())"



        range.FormatConditions(range.FormatConditions.Count).SetFirstPriority

        With range.FormatConditions(1).Interior
            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4
            .TintAndShade = 0.8
        End With
        range.FormatConditions(1).StopIfTrue = False
        设置为自动计算()

    End Function


    Public Function 刷新聚光灯效果() As Boolean
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range
        Dim selectRange As Excel.Range = app.Selection
        If 是否打开聚光灯效果 = True Then

            'range.FormatConditions.Add(Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
            '                           ,
            '                           Formula1:="=OR(CELL(""col"")=COLUMN(),CELL(""row"")=ROW())")

            'sheet.Cells.FormatConditions.Delete()

            设置为手动计算()
            Dim n As Integer = sheet.Cells.FormatConditions.Count
            Try
                For i = 1 To n
                    If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression Then
                        If sheet.Cells.FormatConditions(i).Formula1.ToString.StartsWith("=OR(AND(ROW()>=") Then
                            'sheet.Cells.FormatConditions(i).Delete()

                            Dim ff As Excel.FormatCondition
                            ff = sheet.Cells.FormatConditions(i)

                            Dim rowMin, rowMax, columnMin, columnMax As Integer
                            rowMin = selectRange.Row
                            rowMax = selectRange.Row + selectRange.Rows.Count - 1
                            columnMin = selectRange.Column
                            columnMax = selectRange.Column + selectRange.Columns.Count - 1

                            'range = app.Union(sheet.Cells(rowMin, 1).resize(selectRange.Rows.Count, 16384),
                            '                  sheet.Cells(1, columnMin).resize(1048576, selectRange.Columns.Count))


                            'ff.ModifyAppliesToRange(range)
                            ff.ModifyEx(Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression,
                                        Formula1:="=OR(AND(ROW()>=" & rowMin & ",ROW()<=" & rowMax & "),AND(COLUMN()>=" & columnMin & ",COLUMN()<=" & columnMax & "))")
                        End If
                    End If
                Next
            Catch ex As Exception
                设置为自动计算()
            End Try







            'Dim rowMin, rowMax, columnMin, columnMax As Integer
            'rowMin = selectRange.Row
            'rowMax = selectRange.Row + selectRange.Rows.Count - 1
            'columnMin = selectRange.Column
            'columnMax = selectRange.Column + selectRange.Columns.Count - 1

            'range = app.Union(sheet.Cells(rowMin, 1).resize(selectRange.Rows.Count, 16384),
            '                  sheet.Cells(1, columnMin).resize(1048576, selectRange.Columns.Count))

            'range.FormatConditions.Add(Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression,
            '                           Formula1:="=OR(AND(ROW()>=" & rowMin & ",ROW()<=" & rowMax & "),AND(COLUMN()>=" & columnMin & ",COLUMN()<=" & columnMax & "))")

            ''''' =OR(AND(COLUMN()>=5,COLUMN()<=9),AND(ROW()>=12,ROW()<=22))
            ''''' "=OR(CELL(""col"")=COLUMN(),CELL(""row"")=ROW())"



            'range.FormatConditions(range.FormatConditions.Count).SetFirstPriority

            'With range.FormatConditions(1).Interior
            '    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            '    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent4
            '    .TintAndShade = 0.799981688894314
            'End With
            'range.FormatConditions(1).StopIfTrue = False

            设置为自动计算()
            Return True


        Else


            Return False


        End If
    End Function

    Public Function 创建当前选择信息表() As Boolean
        If 是否存在工作表(当前选择信息表名) = False Then
            Dim sheet As Excel.Worksheet = app.ActiveSheet
            Dim range As Excel.Range = sheet.Cells

            Dim tempSheet As Excel.Worksheet = 新建工作表(当前选择信息表名, False, True)




            tempSheet.Cells(1, 1).value = "活动表名"
            tempSheet.Cells(2, 1).value = "活动单元格行号"
            tempSheet.Cells(3, 1).value = "活动单元格列号"
            tempSheet.Cells(4, 1).value = "活动单元格的值"

            tempSheet.Cells(5, 1).value = "选区最小行号"
            tempSheet.Cells(6, 1).value = "选区最大行号"
            tempSheet.Cells(7, 1).value = "选区最小列号"
            tempSheet.Cells(8, 1).value = "选区最大列号"



            设置单元格格式(tempSheet.Cells(4, 2), "文本")
            tempSheet.Cells(4, 2).value = "★NULL空[空]Nothing[]0★"
            'Application.CutCopyMode = False


            Return True

        Else
            Return False
        End If

    End Function

    Public Function 刷新选择信息表() As Boolean
        Try
            设置为手动计算()
            Dim rowMin, rowMax, columnMin, columnMax As Integer
            Dim selectRange As Excel.Range = app.Selection
            rowMin = selectRange.Row
            rowMax = selectRange.Row + selectRange.Rows.Count - 1
            columnMin = selectRange.Column
            columnMax = selectRange.Column + selectRange.Columns.Count - 1

            Dim tempSheet As Excel.Worksheet = app.Sheets(当前选择信息表名)

            tempSheet.Cells(1, 2).value = app.ActiveSheet.Name
            tempSheet.Cells(2, 2).value = app.ActiveCell.Row
            tempSheet.Cells(3, 2).value = app.ActiveCell.Column



            If app.ActiveCell.Value <> Nothing Then
                tempSheet.Cells(4, 2).value = app.ActiveCell.Value
            Else
                tempSheet.Cells(4, 2).value = "★NULL空[空]Nothing[]0★"
            End If
            tempSheet.Cells(5, 2).value = rowMin
            tempSheet.Cells(6, 2).value = rowMax
            tempSheet.Cells(7, 2).value = columnMin
            tempSheet.Cells(8, 2).value = columnMax
            设置为自动计算()
        Catch ex As Exception
            设置为自动计算()
        End Try

    End Function



    Public Function 统计信息() As String
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range = app.Selection
        Dim cell As Excel.Range = app.ActiveCell
        Dim currentSheetNamme As String = sheet.Name
        Dim currentRow As Integer = cell.Row
        Dim currentColumn As Integer = cell.Column
        Dim currentValue As String = cell.Value
        Dim equalNum As Integer = 0

        For Each cel As Excel.Range In sheet.UsedRange


            If cel.Value IsNot Nothing And cell.Value IsNot Nothing Then
                If cel.Value.ToString = cell.Value.ToString Then
                    equalNum += 1
                End If
            End If
        Next
        Dim cellNum As String
        Try
            cellNum = range.Count.ToString
        Catch ex As Exception
            cellNum = range.Rows.Count & "*" & range.Columns.Count
        End Try

        Return "表名：" & currentSheetNamme & vbCrLf &
            "选区计数：" & cellNum & vbCrLf &
            "当前位置：" & "(" & currentRow & "," & currentColumn & ")" & vbCrLf &
            "当前内容：" & currentValue & vbCrLf &
            "相同个数：" & equalNum & vbCrLf



    End Function



    Public Function 删除关注重复值条件格式(sheet As Excel.Worksheet, Optional 是否删除后台数据表 As Boolean = True) As Boolean
        '''''''''删除条件格式

        'Dim sheet As Excel.Worksheet = app.ActiveSheet
        Try
            Dim range As Excel.Range = sheet.Cells
            Dim n As Integer = range.FormatConditions.Count
            Try
                For i = 1 To n
                    If range.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue Then
                        If range.FormatConditions(i).Formula1 = "=" & 当前选择信息表名 & "!$B$4" Then
                            range.FormatConditions(i).Delete()
                        End If
                    End If

                Next
            Catch ex As Exception

            End Try




            ''''''''删除表格
            If 是否删除后台数据表 = True Then
                app.Sheets(当前选择信息表名).Cells(1, 1).value = "~~~@@@@@#####￥￥￥￥￥￥￥￥%%%%唯一不重复的值"
                app.DisplayAlerts = False '删除工作表不再提示
                app.Sheets(当前选择信息表名).Delete() '删除工作表
                app.DisplayAlerts = True  '删除工作表提示
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try


    End Function

    Public Function 添加关注相同值条件格式(sheet As Excel.Worksheet)

        创建当前选择信息表()

        'Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range = sheet.Cells
        Dim n As Integer = sheet.Cells.FormatConditions.Count
        Try
            For i = 1 To n
                If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue Then
                    If sheet.Cells.FormatConditions(i).Formula1.ToString = "='" & 当前选择信息表名 & "'!$B$4" Then
                        sheet.Cells.FormatConditions(i).Delete()
                    End If
                End If

            Next
        Catch ex As Exception

        End Try









        range.FormatConditions.Add(Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                   Operator:=Excel.XlFormatConditionOperator.xlEqual,
                                   Formula1:="='" & 当前选择信息表名 & "'!$B$4")

        range.FormatConditions(range.FormatConditions.Count).SetFirstPriority


        'With range.FormatConditions(1).Font
        '    .Color = -16383844
        '    .TintAndShade = 0
        'End With
        With range.FormatConditions(1).Interior
            .PatternColorIndex = Excel.Constants.xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        range.FormatConditions(1).StopIfTrue = False



        'With range.FormatConditions(1).Interior
        '    .Pattern = Excel.XlPattern.xlPatternRectangularGradient  'xlPatternRectangularGradient
        '    .Gradient.RectangleLeft = 0
        '    .Gradient.RectangleLeft = 0.5
        '    .Gradient.RectangleRight = 0.5
        '    .Gradient.RectangleTop = 0.5
        '    .Gradient.RectangleBottom = 0.5
        '    .Gradient.ColorStops.Clear
        'End With
        'With range.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        '    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
        '    .TintAndShade = 0
        'End With
        'With range.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        '    .Color = 14452223 '3932159
        '    .TintAndShade = 0
        'End With


        'With range.FormatConditions(1).Interior
        '    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
        '    .Color = 16753663
        '    .TintAndShade = 0
        'End With
        'range.FormatConditions(1).StopIfTrue = False

    End Function


    Public Function 删除聚光灯效果条件格式(sheet As Excel.Worksheet) As Boolean
        Dim n As Integer = sheet.Cells.FormatConditions.Count
        Try
            For i = 1 To n
                If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression Then
                    If sheet.Cells.FormatConditions(i).Formula1.ToString.StartsWith("=OR(AND(ROW()>=") Then
                        sheet.Cells.FormatConditions(i).Delete()
                    End If
                End If

            Next
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function 开启聚光灯效果(sheet As Excel.Worksheet) As Boolean
        Dim n As Integer = sheet.Cells.FormatConditions.Count

        Dim 是否存在聚光灯效果 As Boolean = False
        Try
            For i = 1 To n
                If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression Then
                    If sheet.Cells.FormatConditions(i).Formula1.ToString.StartsWith("=OR(AND(ROW()>=") Then
                        是否存在聚光灯效果 = True
                    End If
                End If
            Next

            If 是否存在聚光灯效果 = False Then
                添加探照灯效果()
            End If

        Catch ex As Exception

        End Try

    End Function




    Public Sub CreateRectangle(ByVal xlWorksheet As Excel.Worksheet, ByVal xlRange As Excel.Range)

        ' 获取工作表的形状集合
        Dim xlShapes = xlWorksheet.Shapes

        ' 获取范围的左上角坐标和宽度和高度
        Dim cellLeft As Single = xlWorksheet.Cells(1, xlRange.Column).Left
        Dim cellTop As Single = xlWorksheet.Cells(1, xlRange.Column).Top
        Dim cellWidth As Single = xlRange.Width
        Dim cellHeight As Single = xlRange.Height





        ' 使用坐标参数向工作表添加一个长方形，并返回一个Shape对象
        Dim xlRectangle = CType(xlShapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, cellLeft, cellTop, cellWidth, cellHeight), Excel.Shape)

        ' 设置长方形对象的一些属性，如名称、填充颜色、透明度、边框颜色等
        xlRectangle.Name = "MyRectangle"
        xlRectangle.Fill.ForeColor.RGB = RGB(255, 0, 0) ' 醒目色为红色，你也可以改成其他颜色
        xlRectangle.Fill.Transparency = 0.5F ' 透明度为50%
        xlRectangle.Line.ForeColor.RGB = RGB(0, 0, 0) ' 边框颜色为黑色

    End Sub



    Public Sub HighlightRowAndColumn(ByVal xlWorksheet As Excel.Worksheet, ByVal xlRange As Excel.Range)

        ' 获取工作表的形状集合
        Dim xlShapes = xlWorksheet.Shapes

        ' 获取范围的左上角坐标和宽度和高度
        Dim cellLeft As Single = xlRange.Left
        Dim cellTop As Single = xlRange.Top
        Dim cellWidth As Single = xlRange.Width
        Dim cellHeight As Single = xlRange.Height

        ' 使用坐标参数向工作表添加两个长方形，并返回两个Shape对象
        ' 第一个长方形覆盖整个行，第二个长方形覆盖整个列
        Dim xlRowRectangle = CType(xlShapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, cellTop, Globals.ThisAddIn.Application.ActiveWindow.VisibleRange.Width, cellHeight), Excel.Shape)
        Dim xlColumnRectangle = CType(xlShapes.AddShape(MsoAutoShapeType.msoShapeRectangle, cellLeft, 0, cellWidth, Globals.ThisAddIn.Application.ActiveWindow.VisibleRange.Height), Excel.Shape)

        ' 设置两个长方形对象的一些属性，如名称、填充颜色、透明度、边框颜色等
        ' 这里使用GDI模式，即将RGB值转换为十进制数值，例如RGB(255,0,0)转换为-16776961（255 + 256 * (255 + 256 * 255)）
        ' 参考链接：https://docs.microsoft.com/en-us/office/vba/api/excel.colorformat.forecolor.rgb
        xlRowRectangle.Name = "MyRowRectangle"
        xlRowRectangle.Fill.ForeColor.RGB = -16776961 ' 醒目色为红色，你也可以改成其他颜色
        xlRowRectangle.Fill.Transparency = 0.9F ' 透明度为50%
        'xlRowRectangle.Line.ForeColor.RGB = -16777216 ' 边框颜色为黑色
        xlRowRectangle.Line.Visible = False

        xlColumnRectangle.Name = "MyColumnRectangle"
        xlColumnRectangle.Fill.ForeColor.RGB = -16776961 ' 醒目色为红色，你也可以改成其他颜色
        xlColumnRectangle.Fill.Transparency = 0.9F ' 透明度为50%
        'xlColumnRectangle.Line.ForeColor.RGB = -16777216 ' 边框颜色为黑色
        xlColumnRectangle.Line.Visible = False
    End Sub

    Public Function 加载表(ComboBox As Windows.Forms.ComboBox) As Integer
        ComboBox.Items.Clear()

        For Each sheet As Excel.Worksheet In app.Worksheets

            ComboBox.Items.Add(sheet.Name)

        Next
        Return app.Worksheets.Count
    End Function
    ''' <summary>
    ''' 结果返回待检查的工作表中是否存在冗余数据
    ''' </summary>
    ''' <returns></returns>
    Public Function 冗余行列检查(sheer As Excel.Worksheet,
                           Optional 是否弹出提示框 As Boolean = True,
                           Optional MaxRowNum As Integer = 10000,
                           Optional MaxColumnNum As Integer = 1000) As Boolean
        Try
            Dim r As Integer = 获取结束单元格(sheer).Row
            Dim c As Integer = 获取结束单元格(sheer).Column
            If r > MaxRowNum Or c > MaxColumnNum Then
                If 是否弹出提示框 = True Then
                    If MsgBox("工作表  """ & sheer.Name & """ 有 " & r & "行  " & c & "列" & " 数据" & vbCrLf &
                                 "可能有大量 空行、空列！建议删除冗余数据。" & vbCrLf &
                                 "若继续操作，可能要耗费很长时间，也可能程序崩溃！" & vbCrLf &
                                 "你确认要继续操作吗？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Return False
                    Else
                        Return True
                    End If
                Else
                    Return True
                End If
            Else
                Return False

            End If
        Catch ex As Exception
            Return True
        End Try



    End Function
    ''' <summary>
    ''' 结果返回待检查的工作表中是否存在冗余数据
    ''' </summary>
    ''' <returns></returns>
    Public Function 冗余行列检查(sheerList As List(Of Excel.Worksheet),
                           Optional 是否弹出提示框 As Boolean = True,
                           Optional MaxRowNum As Integer = 10000,
                           Optional MaxColumnNum As Integer = 1000) As Boolean
        For Each sheet As Excel.Worksheet In sheerList
            If 冗余行列检查(sheet, 是否弹出提示框, MaxRowNum, MaxColumnNum) = True Then
                Return True
            End If
        Next
        Return False
    End Function



    Function 格式化工作表名(SheetName As String) As String
        If SheetName = Nothing Then
            Return Nothing
        End If

        Dim pattern As String = "[\:\/\\?\*\\[\\]\r\n\t]|[：：]"

        ' 使用正则表达式替换特殊字符
        SheetName = Regex.Replace(SheetName, pattern, "")

        ' 使用 Trim() 删除开头和结尾的空白字符
        SheetName = SheetName.Trim()

        ' 删除重复的空白字符 
        SheetName = Regex.Replace(SheetName, "\s+", " ")

        If SheetName.Length > 31 Then
            SheetName = SheetName.Substring(0, 31)
        End If

        Return SheetName

    End Function

    Public Function GetValueFromString(str As String, key As String) As String
        If str IsNot Nothing And key IsNot Nothing Then
            If Not key.StartsWith("<") Then
                key = "<" & key
            End If
            If Not key.EndsWith(">") Then
                key = key & ">"
            End If
            Dim startIndex, overIndex As Integer
            startIndex = str.IndexOf(key)
            If startIndex < 0 Then
                流水信息.记录信息("警告！  获取 " & key & " 失败")
                Return Nothing
            End If
            startIndex += key.Length

            key = key.Trim("<").Trim(">")
            key = "</" & key & ">"
            overIndex = str.IndexOf(key)
            If overIndex < 0 Then
                流水信息.记录信息("警告！  获取 " & key & " 失败")
                Return Nothing
            End If
            overIndex -= 1
            Return str.Substring(startIndex, overIndex - startIndex + 1)
        Else
            流水信息.记录信息("警告！  获取 " & key & " 失败")
            Return Nothing
        End If

    End Function
    Public Function ListToString(序列 As List(Of Object), 分隔符 As String) As String
        Dim result As String = ""

        For Each element As Object In 序列
            result &= element.ToString() & 分隔符
        Next
        If result.EndsWith(分隔符) Then
            result.Trim(分隔符)
        End If
        Return result
    End Function
    Public Function ListToString(序列 As Collections.ArrayList, 分隔符 As String) As String
        Dim result As String = ""

        For Each element As Object In 序列
            result &= element.ToString() & 分隔符
        Next
        If result.EndsWith(分隔符) Then
            result.Trim(分隔符)
        End If
        Return result
    End Function

    Public Function StringToList(str As String, 分隔符 As String) As String()
        Try
            If str Is Nothing Then
                Return Nothing
            End If
            Dim result() As String = str.Split(分隔符)
            Dim n As Integer = 0



            Dim newList As New List(Of String)()

            ' 遍历原始数组，将非空字符串添加到新列表中
            For Each s As String In result
                If Not String.IsNullOrEmpty(s) Then
                    newList.Add(s)
                End If
            Next

            ' 将新列表转换为数组
            Dim newArr() As String = newList.ToArray()

            ' 输出结果


            Return newArr
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Sub 加载列标题到数据表(sheet As Excel.Worksheet,
                         标题行行号 As Integer,
                         MyDataGridView As Windows.Forms.DataGridView,
                         Optional 是否只读 As Boolean = True)


        If sheet IsNot Nothing Then
            MyDataGridView.Columns.Clear()
            MyDataGridView.Columns.Add(sheet.Name, sheet.Name)
            'MyDataGridView.Columns.Add("匹配值", "匹配值")
            MyDataGridView.Columns(0).ReadOnly = 是否只读
            Dim str As String = ""

            If 标题行行号 > sheet.UsedRange.Rows.Count Then
                MsgBox("标题行行号超出使用区范围，请重新设置标题行所在区域的行号。")
                Exit Sub
            End If
            Dim n As Integer = 0
            For Each cell As Excel.Range In sheet.UsedRange.Rows.Item(标题行行号).Cells
                n = MyDataGridView.Rows.Add()
                MyDataGridView.Rows(n).Cells(0).Value = ToStr(cell)
                'MyDataGridView.Rows(n).Cells(1).Value = ""
            Next
        End If
    End Sub

End Module
