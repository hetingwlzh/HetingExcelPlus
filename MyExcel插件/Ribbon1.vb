Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools.Excel
Imports Microsoft.Office.Tools.Excel.Extensions


Imports System.Deployment.Application

''' <summary>
''' '''
''' </summary>
''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Class Ribbon1
    Public Function 功能设置页代码框架() As Excel.Worksheet

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







            Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
            添加按钮(按钮位置, "开始执行", "Start")
            SetSheet.Cells.EntireColumn.AutoFit() '自动列宽

            Return SetSheet
        End If
    End Function
    Public Function 功能执行页代码示例()

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





        app.DisplayAlerts = False '删除工作表不再提示
        SetSheet.Delete() '删除工作表
        Dim 缓冲表 As Excel.Worksheet = 获取缓冲表() '操作结果暂时写入缓冲表，完成后再拷贝
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码
        ''''''这里是 功能实现代码 功能实现代码







        自动列宽(缓冲表)
        居中(缓冲表.UsedRange)
        创建结果页并显示结果("执行结果页", 缓冲表)
        app.DisplayAlerts = True  '删除工作表回复提示
    End Function

    Public Class record
        Public key As String
        Public value1 As String
        Public value2 As String
        Public Sub New(keyText As String, v1 As String, Optional v2 As String = "")
            key = keyText
            value1 = v1
            value2 = v2
        End Sub
    End Class
    Public Sub 初始化后台数据()
        If MsgBox("警告！！！真的要初始化所有后台设置数据吗？", MsgBoxStyle.YesNo, "初始化数据！") = MsgBoxResult.Yes Then


            If My.Settings.枚举序列 Is Nothing Then
                My.Settings.枚举序列 = New System.Collections.ArrayList
            End If

            If My.Settings.循环色序列 Is Nothing Then
                My.Settings.循环色序列 = New System.Collections.ArrayList
            Else
                My.Settings.循环色序列.Clear()
            End If



            My.Settings.循环色序列.Add(RGB(204, 239, 252))
            My.Settings.循环色序列.Add(RGB(233, 246, 220))
            My.Settings.循环色序列.Add(RGB(244, 222, 254))
            My.Settings.循环色序列.Add(RGB(255, 242, 204))
            My.Settings.循环色序列.Add(RGB(226, 214, 236))
            My.Settings.循环色序列.Add(RGB(242, 204, 204))
            My.Settings.循环色序列.Add(RGB(204, 226, 242))
            My.Settings.循环色序列.Add(RGB(255, 255, 204))
            My.Settings.循环色序列.Add(RGB(204, 239, 220))
            My.Settings.循环色序列.Add(RGB(204, 201, 255))



            My.Settings.枚举序列.Clear()






            My.Settings.Save()

        End If

    End Sub
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

        CheckBox1.Checked = My.Settings.是否以非空数据为边
        CheckBox2.Checked = My.Settings.是否启用任务窗格
        CheckBox3.Checked = 是否突出显示相同值
        'My.Settings.list1.Clear()

        My.Settings.Save()
        If My.Settings.枚举序列 Is Nothing Or
           My.Settings.循环色序列 Is Nothing Then

            初始化后台数据()
        End If






    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        'MsgBox(app.ActiveSheet.name)
        获取用户区域().Select()

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        sheet.UsedRange.Cells(1, 1).Select()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        获取结束单元格(app.ActiveSheet).Select()
    End Sub











    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        app.ActiveSheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        app.ActiveSheet.Cells.EntireRow.AutoFit()
    End Sub







    Private Sub Ribbon1_Close(sender As Object, e As EventArgs) Handles MyBase.Close

        My.Settings.Save()
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        初始化后台数据()
    End Sub



    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        app.Selection.NumberFormatLocal = "@" '设置单元格为文本格式
    End Sub



    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs)
        Dim DataSheet As Excel.Worksheet '后台数据表
        Dim is_new As Boolean = False
        If 是否存在工作表(考号信息表) = False Then
            DataSheet = 新建工作表(考号信息表)
            is_new = True

        Else
            DataSheet = app.Worksheets(考号信息表)
            is_new = False
        End If
        DataSheet.Cells(1, 1) = "班级"
        DataSheet.Cells(1, 2) = "人数"
        DataSheet.Cells(1, 3) = "前缀"
        DataSheet.Cells(1, 4) = "班号位数"
        DataSheet.Cells(1, 5) = "个人起始号"
        DataSheet.Cells(1, 6) = "个号位数"
        DataSheet.Cells(1, 7) = "后缀"

        DataSheet.Cells(1, 9) = "总序号"
        DataSheet.Cells(1, 10) = "班级"
        DataSheet.Cells(1, 11) = "考号"
        DataSheet.Cells(1, 12) = "班级序号"


        DataSheet.Cells(2, 1).Select()
        app.ActiveWindow.FreezePanes = True
        DataSheet.Range(DataSheet.Cells(1, 1), DataSheet.Cells(1, 7)).Interior.Color = RGB(200, 255, 200)
        DataSheet.Range(DataSheet.Cells(1, 9), DataSheet.Cells(1, 12)).Interior.Color = RGB(255, 200, 200)

        DataSheet.Columns(1).NumberFormatLocal = "@" '设置单元格为文本格式
        DataSheet.Columns(3).NumberFormatLocal = "@" '设置单元格为文本格式
        DataSheet.Columns(7).NumberFormatLocal = "@" '设置单元格为文本格式
        DataSheet.Columns(10).NumberFormatLocal = "@" '设置单元格为文本格式
        DataSheet.Columns(11).NumberFormatLocal = "@" '设置单元格为文本格式



        If is_new = False Then
            Dim CurrentRow As Integer = 2
            Dim MaxRow = 获取用户区域(DataSheet).Rows.Count
            Dim 前缀, 后缀, 班级， 班级代码 As String
            Dim 个人起始号, 个号位数， 班号位数， 人数 As Integer

            For i = 2 To MaxRow
                If DataSheet.Cells(i, 1).Value <> Nothing Then
                    班级 = DataSheet.Cells(i, 1).Value
                Else
                    Exit For
                End If

                If DataSheet.Cells(i, 2).Value <> Nothing Then
                    人数 = DataSheet.Cells(i, 2).Value
                Else
                    Exit For
                End If

                If DataSheet.Cells(i, 3).Value <> Nothing Then
                    前缀 = DataSheet.Cells(i, 3).Value
                Else
                    前缀 = ""
                End If

                If DataSheet.Cells(i, 4).Value <> Nothing Then
                    班号位数 = DataSheet.Cells(i, 4).Value
                    班级代码 = Format$(Int(班级)， Zero(班号位数))
                Else
                    班级代码 = ""
                End If


                If DataSheet.Cells(i, 5).Value <> Nothing Then
                    个人起始号 = DataSheet.Cells(i, 5).Value
                Else
                    个人起始号 = 1
                End If

                If DataSheet.Cells(i, 6).Value <> Nothing Then
                    个号位数 = DataSheet.Cells(i, 6).Value
                Else
                    个号位数 = 3
                End If

                If DataSheet.Cells(i, 7).Value <> Nothing Then
                    后缀 = DataSheet.Cells(i, 7).Value
                Else
                    后缀 = ""
                End If



                For n = 1 To Int(DataSheet.Cells(i, 2).Value)
                    DataSheet.Cells(CurrentRow, 9) = CurrentRow - 1
                    DataSheet.Cells(CurrentRow, 10) = 班级
                    DataSheet.Cells(CurrentRow, 11) = 前缀 & 班级代码 & Format$(个人起始号 + n - 1, Zero(个号位数)) & 后缀
                    DataSheet.Cells(CurrentRow, 12) = n
                    CurrentRow += 1
                Next

            Next



        End If


        DataSheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub Button15_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click

        '当前选区 = 区域交集(app.Selection, app.Selection.Worksheet.UsedRange)
        Dim 枚举 As New 枚举控件
        添加或显示功能控件(枚举, "枚举")


        'Dim 设置标题 As String = "枚举"
        'Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        'If SetSheet Is Nothing Then '还没完成设置，以下开始设置
        '    Dim range As Excel.Range = app.Selection
        '    SetSheet = 打开设置页(设置标题)

        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码


        '    SetSheet.Cells(2, 1).value = "行数"
        '    SetSheet.Cells(3, 1).value = "列数"
        '    SetSheet.Cells(4, 1).value = "读取顺序"
        '    SetSheet.Cells(5, 1).value = "是否剔除重复值"
        '    SetSheet.Cells(6, 1).value = "是否忽略空值"

        '    SetSheet.Cells(2, 2).value = ""
        '    SetSheet.Cells(3, 2).value = "1"

        '    With SetSheet.Cells(4, 2).Validation
        '        .Delete
        '        .Add(Excel.XlDVType.xlValidateList,,, Formula1:="先行后列,先列后行")
        '        .IgnoreBlank = True
        '        .InCellDropdown = True
        '        .InputTitle = ""
        '        .ErrorTitle = ""
        '        .InputMessage = ""
        '        .ErrorMessage = ""
        '        .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
        '        .ShowInput = True
        '        .ShowError = True
        '    End With
        '    If My.Settings.是否先行后列 = True Then
        '        SetSheet.Cells(4, 2).value = "先行后列"
        '    Else
        '        SetSheet.Cells(4, 2).value = "先列后行"
        '    End If

        '    With SetSheet.Cells(5, 2).Validation
        '        .Delete
        '        .Add(Excel.XlDVType.xlValidateList,,, Formula1:="是,否")
        '        .IgnoreBlank = True
        '        .InCellDropdown = True
        '        .InputTitle = ""
        '        .ErrorTitle = ""
        '        .InputMessage = ""
        '        .ErrorMessage = ""
        '        .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
        '        .ShowInput = True
        '        .ShowError = True
        '    End With
        '    If My.Settings.是否剔除重复值 = True Then
        '        SetSheet.Cells(5, 2).value = "是"
        '    Else
        '        SetSheet.Cells(5, 2).value = "否"
        '    End If


        '    With SetSheet.Cells(6, 2).Validation
        '        .Delete
        '        .Add(Excel.XlDVType.xlValidateList,,, Formula1:="是,否")
        '        .IgnoreBlank = True
        '        .InCellDropdown = True
        '        .InputTitle = ""
        '        .ErrorTitle = ""
        '        .InputMessage = ""
        '        .ErrorMessage = ""
        '        .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
        '        .ShowInput = True
        '        .ShowError = True
        '    End With
        '    If My.Settings.是否忽略空值 = True Then
        '        SetSheet.Cells(6, 2).value = "是"
        '    Else
        '        SetSheet.Cells(6, 2).value = "否"
        '    End If

        '    Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
        '    添加按钮(按钮位置, "开始枚举", "Start")
        '    设置为表头样式(SetSheet, 0, 1)
        '    自动列宽(SetSheet)

        'End If








    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click

        Dim 所选区域 As Excel.Range = app.Selection
        Dim 设置标题 As String = "表间匹配"
        Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        If SetSheet Is Nothing Then '还没完成设置，以下开始设置
            SetSheet = 打开设置页(设置标题)

            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            '''



            创建属性表设置页(SetSheet, 所选区域)
            创建数据表设置页(SetSheet)



            Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
            添加按钮(按钮位置, "开始执行", "Start")
            'Else '已经完成设置，一下开始执行功能操作


        End If


    End Sub



    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        Dim 设置标题 As String = "生成编号"
        Dim Setsheet As Excel.Worksheet = 是否已经设置过(设置标题)
        If Setsheet Is Nothing Then '还没完成设置，以下开始设置
            'MsgBox("还没设置")
            Setsheet = 打开设置页(设置标题)

            Setsheet.Cells(2, 1) = "类别"
            Setsheet.Cells(2, 2) = "个数"
            Setsheet.Cells(2, 3) = "前缀"
            Setsheet.Cells(2, 4) = "类别位数"
            Setsheet.Cells(2, 5) = "个体起始编号"
            Setsheet.Cells(2, 6) = "个号位数"
            Setsheet.Cells(2, 7) = "后缀"

            Setsheet.Cells(2, 9) = "总序号"
            Setsheet.Cells(2, 10) = "类别"
            Setsheet.Cells(2, 11) = "编号"
            Setsheet.Cells(2, 12) = "类中序号"


            Setsheet.Cells(2, 1).Select()
            app.ActiveWindow.FreezePanes = True
            Setsheet.Range(Setsheet.Cells(2, 1), Setsheet.Cells(2, 7)).Interior.Color = RGB(200, 255, 200)
            Setsheet.Range(Setsheet.Cells(2, 9), Setsheet.Cells(2, 12)).Interior.Color = RGB(255, 200, 200)

            Setsheet.Columns(1).NumberFormatLocal = "@" '设置单元格为文本格式
            Setsheet.Columns(3).NumberFormatLocal = "@" '设置单元格为文本格式
            Setsheet.Columns(7).NumberFormatLocal = "@" '设置单元格为文本格式
            Setsheet.Columns(10).NumberFormatLocal = "@" '设置单元格为文本格式
            Setsheet.Columns(11).NumberFormatLocal = "@" '设置单元格为文本格式

            设置为表头样式(Setsheet, 2, 0,,,,)

            Dim 按钮位置 As Excel.Range = Setsheet.Cells(1, 10)
            添加按钮(按钮位置, "开始编号", "Start")
            居中(Setsheet.UsedRange)







        End If
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click



        Dim 所选区域 As Excel.Range = app.Selection
        Dim 设置标题 As String = "批量匹配属性"
        Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        If SetSheet Is Nothing Then '还没完成设置，以下开始设置
            Dim range As Excel.Range = app.Selection
            SetSheet = 打开设置页(设置标题)

            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码
            ''''''这里是 创建设置页面内容的代码



            创建属性表设置页(SetSheet, 所选区域)




            Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
            添加按钮(按钮位置, "开始执行", "Start")



        End If


    End Sub




    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click



        '当前选区 = app.Selection
        'Dim 设置标题 As String = "合并表"
        'Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
        'If SetSheet Is Nothing Then '还没完成设置，以下开始设置
        '    Dim range As Excel.Range = app.Selection
        '    SetSheet = 打开设置页(设置标题)

        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码
        '    ''''''这里是 创建设置页面内容的代码



        '    SetSheet.Cells(2, 1).value = "合并那些工作表"




        '    With SetSheet.Cells(2, 2).Validation
        '        .Delete
        '        .Add(Excel.XlDVType.xlValidateList,,, Formula1:="左边表格,右边表格")
        '        .IgnoreBlank = True
        '        .InCellDropdown = True
        '        .InputTitle = ""
        '        .ErrorTitle = ""
        '        .InputMessage = ""
        '        .ErrorMessage = ""
        '        .IMEMode = Excel.XlIMEMode.xlIMEModeNoControl
        '        .ShowInput = True
        '        .ShowError = True
        '    End With

        '    SetSheet.Cells(2, 2).value = "左边表格"


        '    Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
        '    添加按钮(按钮位置, "开始合并")

        '    居中(SetSheet.UsedRange)
        '    自动列宽(SetSheet)
        'End If
        Dim 合并对象 As New 合并工作表控件

        'MyForm1.GroupBox1.Controls.Add(合并对象)
        添加或显示功能控件(合并对象, "合并表格")


    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        'Dim sheet As Excel.Worksheet = app.ActiveSheet
        'Dim range, sel As Excel.Range
        'sel = app.Intersect(app.Selection, sheet.UsedRange)
        '排名(sel)

        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim 排名 As New 排名控件
        'Dim worksheet As Microsoft.Office.Tools.Excel.Worksheet = Globals.Factory.GetVstoObject(sheet)
        'worksheet.Controls.AddControl(SetControl1, sheet.Cells(2, 8), "SetControl" & Now.Ticks)

        添加或显示功能控件(排名, "排名方式")
        '添加控件到工作表(sheet, 排名, sheet.Cells(5, 3))

    End Sub



    Private Sub Gallery1_ButtonClick(sender As Object, e As RibbonControlEventArgs) Handles Gallery1.ButtonClick
        If sender.label = "设置" Then
            插件设置()
        ElseIf sender.Label = "颜色代码" Then
            颜色代码()
        End If


    End Sub










    Private Sub ToggleButton1_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton1.Click
        MySelectForm.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        当前选区.Worksheet.Activate()
        当前选区.Select()
    End Sub

    'Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
    '    当前选区 = app.Selection
    '    Dim 设置标题 As String = "开始涂色"
    '    Dim SetSheet As Excel.Worksheet = 是否已经设置过(设置标题)
    '    If SetSheet Is Nothing Then '还没完成设置，以下开始设置
    '        SetSheet = 打开设置页(设置标题)
    '        ''''''这里是 创建设置页面内容的代码
    '        ''''''这里是 创建设置页面内容的代码
    '        ''''''这里是 创建设置页面内容的代码
    '        ''''''这里是 创建设置页面内容的代码
    '        ''''''这里是 创建设置页面内容的代码
    '        ''''''这里是 创建设置页面内容的代码

    '        SetSheet.Cells(2, 1).value = "地址"
    '        SetSheet.Cells(2, 2).WrapText = True
    '        添加按钮(SetSheet.Cells(2, 3), "选择区域", "区域1", "B2")



    '        SetSheet.Cells(3, 1).value = "地址"
    '        SetSheet.Cells(3, 2).WrapText = True
    '        添加按钮(SetSheet.Cells(3, 3), "选择区域", "区域2", "B3")



    '        设置为表头样式(SetSheet, 0, 1)
    '        Dim 按钮位置 As Excel.Range = SetSheet.Cells(1, 10)
    '        添加按钮(按钮位置, "开始涂色", "Start")
    '        SetSheet.Cells.EntireColumn.AutoFit() '自动列宽

    '    End If
    'End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        'Dim sheet As Excel.Worksheet = app.ActiveSheet
        'Dim range As Excel.Range
        'range = app.Selection
        '分类计数(range)
        Dim 分类计数 As New 分类编号控件
        添加或显示功能控件(分类计数, "分类编号")
    End Sub



    Private Sub Button29_Click(sender As Object, e As RibbonControlEventArgs) Handles Button29.Click
        Dim 生成随机数 As New 随机数控件
        添加或显示功能控件(生成随机数, "生成随机数")
    End Sub





    Private Sub Button32_Click(sender As Object, e As RibbonControlEventArgs) Handles Button32.Click
        Dim 按行匹配 As New 行匹配
        'MyForm1.GroupBox1.Controls.Add(合并对象)
        添加或显示功能控件(按行匹配, "按行匹配")
    End Sub

    Private Sub Button33_Click(sender As Object, e As RibbonControlEventArgs) Handles Button33.Click
        Dim 按列汇总 As New 列汇总控件
        'MyForm1.GroupBox1.Controls.Add(合并对象)
        添加或显示功能控件(按列汇总, "按列汇总")
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        Dim 循环色 As New 循环色控件
        添加或显示功能控件(循环色, "循环色")
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        Dim 条件编号 As New 编号控件
        添加或显示功能控件(条件编号, "条件编号", True)
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        生成分类统计表(app.ActiveSheet, 1, 11, ,, 1)
    End Sub

    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click
        Dim 实验窗口 As New Form1
        实验窗口.Show()
    End Sub

    Private Sub Button35_Click(sender As Object, e As RibbonControlEventArgs) Handles Button35.Click
        Dim 分类数量统计 As New 分类数量统计表控件
        'MyForm1.GroupBox1.Controls.Add(合并对象)
        添加或显示功能控件(分类数量统计, "分类数量统计")
    End Sub

    Private Sub Button36_Click(sender As Object, e As RibbonControlEventArgs) Handles Button36.Click
        '选取拖拽填充实验
        拖拽填充(app.Selection)
    End Sub

    Private Sub Button37_Click(sender As Object, e As RibbonControlEventArgs) Handles Button37.Click
        Dim 校验 As New 校验控件
        添加或显示功能控件(校验, "校验"， True)
    End Sub

    Private Sub Button38_Click(sender As Object, e As RibbonControlEventArgs) Handles Button38.Click
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range = app.Selection
        MsgBox(range.Count)

    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs) Handles Button10.Click
        Dim 实验 As New 实验控件
        添加或显示功能控件(实验, "实验")
    End Sub

    Private Sub Gallery1_Click(sender As Object, e As RibbonControlEventArgs) Handles Gallery1.Click

    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        Dim 分类包含统计 As New 分类包含统计控件
        添加或显示功能控件(分类包含统计, "分类包含统计")

    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        显示插件设置()
    End Sub

    Private Sub Button39_Click(sender As Object, e As RibbonControlEventArgs) Handles Button39.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim currentCell As Excel.Range = app.ActiveCell
        Dim 当前列号 As Integer = currentCell.Column
        Dim 当前行号 As Integer = currentCell.Row
        Dim 数据结尾行号 As Integer = 获取非空列数据尾行号(Sheet, 当前列号)
        Dim 用户区结尾行号 As Integer = 获取结束行号(Sheet)
        If 当前行号 = 数据结尾行号 Then
            Sheet.Cells(用户区结尾行号, 当前列号).Select()
        ElseIf 当前行号 = 用户区结尾行号 Then
            Sheet.Cells(数据结尾行号, 当前列号).Select()
        Else
            Sheet.Cells(数据结尾行号, 当前列号).Select()
        End If






    End Sub

    Private Sub Button43_Click(sender As Object, e As RibbonControlEventArgs) Handles Button43.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim selectRange As Excel.Range = app.Selection

        Dim 左上单元格 As Excel.Range = selectRange.Cells(1, 1)




        If My.Settings.是否以非空数据为边 = True Then
            Dim 最大尾行号 As Integer = Sheet.UsedRange.Row
            Dim 当前尾行号 As Integer = 最大尾行号
            For Each column As Excel.Range In selectRange.Columns
                当前尾行号 = 获取非空列数据尾行号(Sheet, column.Column)
                If 最大尾行号 < 当前尾行号 Then
                    最大尾行号 = 当前尾行号
                End If
            Next
            'If 最大尾行号 >= 当前行号 Then
            '    selectRange.Cells(1, 1).Resize(最大尾行号 - 当前行号 + 1, selectRange.Columns.Count).Select()
            'End If
            MyRange(Sheet, 左上单元格.Row, 左上单元格.Column, 最大尾行号, 左上单元格.Column + selectRange.Columns.Count - 1).Select()
        Else
            MyRange(Sheet, 左上单元格.Row, 左上单元格.Column, 获取结束行号(Sheet), 左上单元格.Column + selectRange.Columns.Count - 1).Select()
            'Dim 末尾行号 As Integer = 获取结束行号(Sheet)
            'If 末尾行号 >= 当前行号 Then
            '    selectRange.Cells(1, 1).Resize(获取结束行号(Sheet) - 当前行号 + 1, selectRange.Columns.Count).Select()
            'End If

        End If







    End Sub

    Private Sub Button40_Click(sender As Object, e As RibbonControlEventArgs) Handles Button40.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim 当前列号 As Integer = app.ActiveCell.Column

        Sheet.Cells(1, 当前列号).Select()


    End Sub

    Private Sub Button41_Click(sender As Object, e As RibbonControlEventArgs) Handles Button41.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim 当前列号 As Integer = app.ActiveCell.Column
        Dim 当前行号 As Integer = app.ActiveCell.Row
        Sheet.Cells(当前行号, 1).Select()

    End Sub

    Private Sub Button42_Click(sender As Object, e As RibbonControlEventArgs) Handles Button42.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim currentCell As Excel.Range = app.ActiveCell
        Dim 当前列号 As Integer = currentCell.Column
        Dim 当前行号 As Integer = currentCell.Row
        Dim 数据结尾列号 As Integer = 获取非空行数据尾列号(Sheet, 当前行号)
        Dim 用户区结尾列号 As Integer = 获取结束列号(Sheet)
        If 当前列号 = 数据结尾列号 Then
            Sheet.Cells(当前行号, 用户区结尾列号).Select()
        ElseIf 当前列号 = 用户区结尾列号 Then
            Sheet.Cells(当前行号, 数据结尾列号).Select()
        Else
            Sheet.Cells(当前行号, 数据结尾列号).Select()
        End If



    End Sub

    Private Sub Button44_Click(sender As Object, e As RibbonControlEventArgs) Handles Button44.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim selectRange As Excel.Range = app.Selection
        Dim 左上单元格 As Excel.Range = selectRange.Cells(1, 1)

        If My.Settings.是否以非空数据为边 = True Then
            Dim 最大尾列号 As Integer = Sheet.UsedRange.Row
            Dim 当前尾列号 As Integer = 最大尾列号
            For Each row As Excel.Range In selectRange.Rows
                当前尾列号 = 获取非空行数据尾列号(Sheet, row.Row)
                If 最大尾列号 < 当前尾列号 Then
                    最大尾列号 = 当前尾列号
                End If
            Next
            MyRange(Sheet, 左上单元格.Row, 左上单元格.Column, 左上单元格.Row + selectRange.Rows.Count - 1, 最大尾列号).Select()
        Else
            MyRange(Sheet, 左上单元格.Row, 左上单元格.Column, 左上单元格.Row + selectRange.Rows.Count - 1, 获取结束列号(Sheet)).Select()
        End If

    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBox1.Click
        My.Settings.是否以非空数据为边 = CheckBox1.Checked
    End Sub

    Private Sub Button45_Click(sender As Object, e As RibbonControlEventArgs) Handles Button45.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim selectRange As Excel.Range = app.Selection


        Dim 左下单元格 As Excel.Range = selectRange.Cells(selectRange.Rows.Count, 1)



        If My.Settings.是否以非空数据为边 = True Then

            Dim 最小首行号 As Integer = 获取结束单元格(Sheet).Row
            Dim 当前首行号 As Integer = 最小首行号
            For Each column As Excel.Range In selectRange.Columns
                当前首行号 = 获取非空列数据首行号(Sheet, column.Column)
                If 最小首行号 > 当前首行号 Then
                    最小首行号 = 当前首行号
                End If
            Next


            MyRange(Sheet, 左下单元格.Row, 左下单元格.Column, 最小首行号, 左下单元格.Column + selectRange.Columns.Count - 1).Select()

        Else

            MyRange(Sheet, 左下单元格.Row, 左下单元格.Column, Sheet.UsedRange.Row, 左下单元格.Column + selectRange.Columns.Count - 1).Select()

        End If
    End Sub

    Private Sub Button46_Click(sender As Object, e As RibbonControlEventArgs) Handles Button46.Click
        Dim Sheet As Excel.Worksheet = app.ActiveSheet
        Dim selectRange As Excel.Range = app.Selection

        Dim 右上角单元格 As Excel.Range = selectRange.Cells(1, selectRange.Columns.Count)

        If My.Settings.是否以非空数据为边 = True Then

            Dim 最小首列号 As Integer = 获取结束单元格(Sheet).Column
            Dim 当前首列号 As Integer = 最小首列号
            For Each row As Excel.Range In selectRange.Rows
                当前首列号 = 获取非空行数据首列号(Sheet, row.Row)
                If 最小首列号 > 当前首列号 Then
                    最小首列号 = 当前首列号
                End If
            Next




            MyRange(Sheet, 右上角单元格.Row, 右上角单元格.Column, 右上角单元格.Row + selectRange.Rows.Count - 1, 最小首列号).Select()

        Else
            MyRange(Sheet, 右上角单元格.Row, 右上角单元格.Column, 右上角单元格.Row + selectRange.Rows.Count - 1, Sheet.UsedRange.Column).Select()

        End If
    End Sub

    Private Sub Button47_Click(sender As Object, e As RibbonControlEventArgs) Handles Button47.Click
        Dim 随机抽取 As New 随机抽取控件
        添加或显示功能控件(随机抽取, "随机抽取")
    End Sub

    Private Sub Button30_Click(sender As Object, e As RibbonControlEventArgs) Handles Button30.Click
        'Me.CustomTaskPanes.Add(New 随机数控件, "任务窗格")
        taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(New 随机数控件, "任务窗格")
        taskPane.Visible = True
    End Sub

    Private Sub Button31_Click(sender As Object, e As RibbonControlEventArgs) Handles Button31.Click
        Dim 分类包含统计 As New 分类包含统计控件
        添加或显示功能控件(分类包含统计, "分类包含统计", False)
    End Sub

    Private Sub Button48_Click(sender As Object, e As RibbonControlEventArgs) Handles Button48.Click
        Dim 实验 As New 实验控件
        添加或显示功能控件(实验, "实验", False)
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBox2.Click
        My.Settings.是否启用任务窗格 = CheckBox2.Checked
    End Sub

    Private Sub Button49_Click(sender As Object, e As RibbonControlEventArgs) Handles Button49.Click
        'InitializeComponent()
        'Dim appd As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
        'Label2.Label = appd.CurrentVersion.ToString()
    End Sub

    Private Sub Button50_Click(sender As Object, e As RibbonControlEventArgs) Handles Button50.Click
        Dim 翻译 As New 翻译控件
        添加或显示功能控件(翻译, "翻译", True)

    End Sub





    Private Sub CheckBox3_Click(sender As Object, e As RibbonControlEventArgs) Handles CheckBox3.Click
        是否突出显示相同值 = CheckBox3.Checked
        If 是否突出显示相同值 = True Then
            添加关注相同值条件格式(app.ActiveSheet)
        Else
            ''''''''''删除条件格式
            ''''
            'Dim sheet As Excel.Worksheet = app.ActiveSheet
            'Dim range As Excel.Range = sheet.Cells
            'Dim n As Integer = range.FormatConditions.Count
            'Try
            '    For i = 1 To n
            '        If range.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue Then
            '            If range.FormatConditions(i).Formula1 = "=" & 当前选择信息表名 & "!$B$4" Then
            '                range.FormatConditions(i).Delete()
            '            End If
            '        End If

            '    Next
            'Catch ex As Exception

            'End Try




            '''''''''删除表格
            'app.Sheets(当前选择信息表名).Cells(1, 1).value = "~~~@@@@@#####￥￥￥￥￥￥￥￥%%%%唯一不重复的值"
            'app.DisplayAlerts = False '删除工作表不再提示
            'app.Sheets(当前选择信息表名).Delete() '删除工作表
            'app.DisplayAlerts = True  '删除工作表提示



            删除关注重复值条件格式(app.ActiveSheet)




        End If
    End Sub

    Private Sub ToggleButton2_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton2.Click
        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range = 获取工作区域(sheet)
        是否打开聚光灯效果 = ToggleButton2.Checked
        If 是否打开聚光灯效果 = True Then

            ''创建选择信息表()
            ''刷新选择信息表()
            'Dim n As Integer = sheet.Cells.FormatConditions.Count

            'Dim 是否存在探照灯效果 As Boolean = False
            'Try
            '    For i = 1 To n
            '        If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression Then
            '            If sheet.Cells.FormatConditions(i).Formula1.ToString.StartsWith("=OR(AND(ROW()>=") Then
            '                是否存在探照灯效果 = True
            '            End If
            '        End If
            '    Next

            '    If 是否存在探照灯效果 = False Then
            '        添加探照灯效果()
            '    End If

            'Catch ex As Exception

            'End Try
            开启聚光灯效果(sheet)




            刷新聚光灯效果()

        Else
            'Dim n As Integer = sheet.Cells.FormatConditions.Count
            'Try
            '    For i = 1 To n
            '        If sheet.Cells.FormatConditions(i).Type = Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression Then
            '            If sheet.Cells.FormatConditions(i).Formula1.ToString.StartsWith("=OR(AND(ROW()>=") Then
            '                sheet.Cells.FormatConditions(i).Delete()
            '            End If
            '        End If

            '    Next
            'Catch ex As Exception

            'End Try

            删除聚光灯效果条件格式(sheet)


        End If
    End Sub

    Private Sub Gallery2_ButtonClick(sender As Object, e As RibbonControlEventArgs) Handles Gallery2.ButtonClick

        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim range As Excel.Range = app.Selection

        If sender.label = "提取数字" Then
            取数字(range)
        ElseIf sender.Label = "逆序" Then
            字符串逆序(range)
        ElseIf sender.Label = "文本拆分" Then
            字符串拆分(range)
        ElseIf sender.Label = "第n次匹配位置" Then
            字符串第n次匹配的位置(range)
        End If
    End Sub

    Private Sub ToggleButton3_Click(sender As Object, e As RibbonControlEventArgs) Handles ToggleButton3.Click
        Dim 信息统计 As New 信息统计控件
        添加或显示功能控件(信息统计, "信息统计", True)
        信息显示控件 = 信息统计
    End Sub
End Class
