Public Class 循环色控件


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '设置循环色(分类区域1, 分类区域2, 分类区域3, 分类区域4, 涂色区域, NumericUpDown1.Value)
        If 区域选择控件1.区域校验(0, True, False, True) And
            区域选择控件1.区域校验(1, False, False, True) And
            区域选择控件1.区域校验(2, False, False, True) And
            区域选择控件1.区域校验(3, False, False, True) And
            区域选择控件1.区域校验(4, False, False, False) Then

            设置循环色(区域选择控件1.获取区域(0),
                  区域选择控件1.获取区域(1),
                  区域选择控件1.获取区域(2),
                  区域选择控件1.获取区域(3),
                  区域选择控件1.获取区域(4),
                  NumericUpDown1.Value)
        Else

            区域选择控件1.显示错误()
        End If


    End Sub






    Private Sub 循环色控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"分类区域1", "分类区域2", "分类区域3", "分类区域4", "涂色区域"})
        区域选择控件1.是否允许编辑区域名 = False

    End Sub

    Public Sub 设置循环色(分类区域1 As Excel.Range,
                     Optional 分类区域2 As Excel.Range = Nothing,
                     Optional 分类区域3 As Excel.Range = Nothing,
                     Optional 分类区域4 As Excel.Range = Nothing,
                     Optional 涂色区域 As Excel.Range = Nothing， Optional 颜色数 As Integer = 3)

        If 分类区域1 Is Nothing Then
            MsgBox("请选择 有数据的单元格区域！")
            Exit Sub
        ElseIf 分类区域1.Columns.Count <> 1 Then
            MsgBox("请选择 单列区域！")
            分类区域1 = Nothing
            Exit Sub
        End If
        Dim 区域数 As Integer = 1

        Dim 列号1, 列号2, 列号3, 列号4 As Integer
        If 分类区域2 Is Nothing Then
            区域数 = 1
            列号1 = 分类区域1.Cells(1, 1).Column
        ElseIf 分类区域3 Is Nothing Then
            区域数 = 2
            If 分类区域2.Columns.Count <> 1 Then
                MsgBox("请选择 单列区域！")
                分类区域2 = Nothing
                Exit Sub
            End If
            列号1 = 分类区域1.Cells(1, 1).Column
            列号2 = 分类区域2.Cells(1, 1).Column
        ElseIf 分类区域4 Is Nothing Then
            区域数 = 3
            If 分类区域3.Columns.Count <> 1 Then
                MsgBox("请选择 单列区域！")
                分类区域3 = Nothing
                Exit Sub
            End If
            列号1 = 分类区域1.Cells(1, 1).Column
            列号2 = 分类区域2.Cells(1, 1).Column
            列号3 = 分类区域3.Cells(1, 1).Column
        Else
            区域数 = 4
            If 分类区域4.Columns.Count <> 1 Then
                MsgBox("请选择 单列区域！")
                分类区域4 = Nothing
                Exit Sub
            End If
            列号1 = 分类区域1.Cells(1, 1).Column
            列号2 = 分类区域2.Cells(1, 1).Column
            列号3 = 分类区域3.Cells(1, 1).Column
            列号4 = 分类区域4.Cells(1, 1).Column
        End If




        '开始执行
        Dim sheet As Excel.Worksheet = 分类区域1.Worksheet
        If 涂色区域 Is Nothing Then
            涂色区域 = sheet.UsedRange
        End If
        Dim 当前分类标志 As String
        Dim 涂色起始行, 涂色结束行 As Integer
        Dim colorNum As Integer = 0
        Dim 是否为新标志 As Boolean = True
        Dim cell As Excel.Range
        Dim 设置区域 As Excel.Range
        涂色起始行 = 分类区域1.Cells(1, 1).Row
        涂色结束行 = 涂色起始行


        If 区域数 = 1 Then
            当前分类标志 = 分类区域1.Cells(1, 1).Value
        ElseIf 区域数 = 2 Then
            当前分类标志 = 分类区域1.Cells(1, 1).Value & 分类区域2.Cells(1, 1).Value
        ElseIf 区域数 = 3 Then
            当前分类标志 = 分类区域1.Cells(1, 1).Value & 分类区域2.Cells(1, 1).Value & 分类区域3.Cells(1, 1).Value
        Else
            当前分类标志 = 分类区域1.Cells(1, 1).Value & 分类区域2.Cells(1, 1).Value & 分类区域3.Cells(1, 1).Value & 分类区域3.Cells(1, 1).Value
        End If


        Dim tempStr As String = ""
        For Each cell In 分类区域1


            If 区域数 = 1 Then
                tempStr = ToStr(cell)
            ElseIf 区域数 = 2 Then
                tempStr = ToStr(cell) & ToStr(MyRange(sheet, cell.Row, 列号2))
            ElseIf 区域数 = 3 Then
                tempStr = ToStr(cell) & ToStr(MyRange(sheet, cell.Row, 列号2)) & ToStr(MyRange(sheet, cell.Row, 列号3))
            Else
                tempStr = ToStr(cell) & ToStr(MyRange(sheet, cell.Row, 列号2)) & ToStr(MyRange(sheet, cell.Row, 列号3)) & ToStr(MyRange(sheet, cell.Row, 列号4))
            End If




            If tempStr = 当前分类标志 Then
                涂色结束行 = cell.Row
            Else

                设置区域 = app.Intersect(涂色区域, sheet.Range(sheet.Cells(涂色起始行, 1), sheet.Cells(涂色结束行, 获取结束单元格(sheet).Column)))
                If 设置区域 IsNot Nothing Then
                    设置填充色(设置区域, My.Settings.循环色序列(colorNum Mod 颜色数))

                End If
                colorNum += 1
                当前分类标志 = tempStr
                涂色起始行 = cell.Row
                涂色结束行 = 涂色起始行
            End If
        Next
        设置区域 = app.Intersect(涂色区域, sheet.Range(sheet.Cells(涂色起始行, 1), sheet.Cells(涂色结束行, 获取结束单元格(sheet).Column)))
        If 设置区域 IsNot Nothing Then
            设置填充色(设置区域, My.Settings.循环色序列(colorNum Mod 颜色数))
        End If

    End Sub














    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        插件设置()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("这里还在施工
请绕行"， MsgBoxStyle.ApplicationModal, "功能说明")
    End Sub


End Class
