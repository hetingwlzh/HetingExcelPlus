Public Class 校验控件
    Public 正确标记 As String
    Public 错误标记 As String
    Public Enum 校验类型
        身份证号 = 0
        手机号 = 1
        电子邮箱 = 2
        汉字 = 3
        数字 = 4
        英文字母 = 5
        标点符号 = 6
        希腊字母 = 7
        长度校验 = 8
        非空校验 = 9
        异同校验 = 10
        包含校验 = 11
        不含校验 = 12
        开头校验 = 13
        结尾校验 = 14
    End Enum

    Public Function To校验类型(类型文本 As String) As 校验类型
        If 类型文本.Contains("身份证") Then
            Return 校验类型.身份证号
        ElseIf 类型文本.Contains("手机号") Then
            Return 校验类型.手机号
        ElseIf 类型文本.Contains("电子邮箱") Then
            Return 校验类型.电子邮箱
        ElseIf 类型文本.Contains("汉字") Then
            Return 校验类型.汉字
        ElseIf 类型文本.Contains("数字") Then
            Return 校验类型.数字
        ElseIf 类型文本.Contains("英文字母") Then
            Return 校验类型.英文字母
        ElseIf 类型文本.Contains("标点符号") Then
            Return 校验类型.标点符号
        ElseIf 类型文本.Contains("希腊字母") Then
            Return 校验类型.希腊字母
        ElseIf 类型文本.Contains("长度校验") Then
            Return 校验类型.长度校验
        ElseIf 类型文本.Contains("非空校验") Then
            Return 校验类型.非空校验
        ElseIf 类型文本.Contains("异同校验") Then
            Return 校验类型.异同校验
        ElseIf 类型文本.Contains("包含校验") Then
            Return 校验类型.包含校验
        ElseIf 类型文本.Contains("不含校验") Then
            Return 校验类型.不含校验
        ElseIf 类型文本.Contains("开头校验") Then
            Return 校验类型.开头校验
        ElseIf 类型文本.Contains("结尾校验") Then
            Return 校验类型.结尾校验
        Else
            Return Nothing
        End If
    End Function
    Private Sub 校验控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        区域选择控件1.设置锁定用户区(False)
        区域选择控件1.设置锁定行(False)
        区域选择控件1.设置锁定列(False)
        区域选择控件1.设置锁定表(False)
        区域选择控件1.是否允许编辑区域名 = False

        区域选择控件1.预设空区域({"校验列1"})


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim 校验 As String = ""
        If ComboBox1.SelectedItem Is Nothing Then
            MsgBox("请选择校验类型")
            Exit Sub
        Else
            校验 = ComboBox1.SelectedItem.ToString
        End If
        If 区域选择控件1.区域校验() = True Then
            If 区域选择控件1.区域记录个数 = 1 Then
                设置校验(To校验类型(校验), 区域选择控件1.获取区域(0), Nothing, NumericUpDown1.Value)
            ElseIf 区域选择控件1.区域记录个数 = 2 Then
                设置校验(To校验类型(校验), 区域选择控件1.获取区域(0), 区域选择控件1.获取区域(1), NumericUpDown1.Value)
            End If
        Else
            区域选择控件1.显示错误()
        End If




    End Sub

    Public Function 获取校验公式(校验方式 As 校验类型, 校验单元格 As Excel.Range, Optional 校验单元格2 As Excel.Range = Nothing) As String
        Dim 单元格地址 As String = 获取地址(校验单元格.Cells(1, 1), False, False, False)
        Dim 单元格地址2 As String
        If 校验单元格2 IsNot Nothing Then
            单元格地址2 = 获取地址(校验单元格2.Cells(1, 1), False, False, False)
        End If
        Dim 公式 As String = ""
        If 校验方式 = 校验类型.身份证号 Then
            '=IF(A3="","",(IF(MID("10X98765432",MOD(SUMPRODUCT(MID(A3,ROW(INDIRECT("1:17")),1)*2^(18-ROW(INDIRECT("1:17")))),11)+1,1)=MID(A3,18,18),"正确","错误")))
            公式 = "=IFERROR(IF(地址="""","""",(IF(MID(""10X98765432"",MOD(SUMPRODUCT(MID(地址,ROW(INDIRECT(""1:17"")),1)*2^(18-ROW(INDIRECT(""1:17"")))),11)+1,1)=MID(地址,18,18),""√"",""×""))),""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.手机号 Then
            '=AND(LEN(A1)=11,OR(MID(A1,1,2)="13",MID(A1,1,2)="18",MID(A1,1,2)="14",MID(A1,1,2)="15",MID(A1,1,2)="16",MID(A1,1,2)="17",MID(A1,1,2)="18",MID(A1,1,2)="19"))
            公式 = "=IF(And(Len(地址)=11,Or(Mid(地址,1,2)=""13"",Mid(地址,1,2)=""18"",Mid(地址,1,2)=""14"",Mid(地址,1,2)=""15"",Mid(地址,1,2)=""16"",Mid(地址,1,2)=""17"",Mid(地址,1,2)=""18"",Mid(地址,1,2)=""19"")),""√"",""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.电子邮箱 Then
            '=AND(IFERROR(FIND(".",A2),FALSE),IFERROR(FIND(".",A2,FIND("@",A2)),FALSE))
            公式 = "=IF(AND(IFERROR(FIND(""."",地址),FALSE),IFERROR(FIND(""."",地址,FIND(""@"",地址)),FALSE)),""√"",""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.异同校验 Then
            '=IF(A1=B1,"√","×")
            公式 = "=IF(地址1=地址2,""√"",""×"")"
            Return 公式.Replace("地址1", 单元格地址).Replace("地址2", 单元格地址2).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.长度校验 Then
            '=LEN(A1)
            Dim 位数 As Integer = NumericUpDown2.Value
            公式 = "=IF(LEN(地址)=" & 位数 & ",""√"",""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.非空校验 Then
            '=IF(ISBLANK(A1),"×","√")
            Dim 位数 As Integer = NumericUpDown2.Value
            公式 = "=IF(ISBLANK(地址),""×"",""√"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.包含校验 Then
            '=IFERROR(IF(FIND("12",A1,1)>0,"√","×"),"×")
            Dim 包含的字符串 As String = TextBox3.Text

            公式 = "=IFERROR(IF(FIND(""" & 包含的字符串 & """,地址,1)>0,""√"",""×""),""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.不含校验 Then
            '=IFERROR(IF(FIND("12",A1,1)>0,"√","×"),"×")
            Dim 包含的字符串 As String = TextBox3.Text

            公式 = "=IFERROR(IF(FIND(""" & 包含的字符串 & """,地址,1)>0,""×"",""√""),""√"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        ElseIf 校验方式 = 校验类型.开头校验 Then
            '=IF(LEFT(A1,LEN("767"))="767","√","×")
            Dim 开头的字符串 As String = TextBox3.Text
            公式 = "=IF(LEFT(地址,LEN(""" & 开头的字符串 & """))=""" & 开头的字符串 & """,""√"",""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)
        ElseIf 校验方式 = 校验类型.结尾校验 Then
            '=IF(LEFT(A1,LEN("767"))="767","√","×")
            Dim 结尾的字符串 As String = TextBox3.Text
            公式 = "=IF(RIGHT(地址,LEN(""" & 结尾的字符串 & """))=""" & 结尾的字符串 & """,""√"",""×"")"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)

        Else

            '=IF(A3="","",(IF(MID("10X98765432",MOD(SUMPRODUCT(MID(A3,ROW(INDIRECT("1:17")),1)*2^(18-ROW(INDIRECT("1:17")))),11)+1,1)=MID(A3,18,18),"正确","错误")))
            公式 = "未知校验类型"
            Return 公式.Replace("地址", 单元格地址).Replace("√", 正确标记).Replace("×", 错误标记)


        End If
    End Function



    Public Sub 设置校验(校验方式 As 校验类型, 校验区域 As Excel.Range, Optional 校验区域2 As Excel.Range = Nothing, Optional 标题列所在行 As Integer = 0)
        Dim 校验数据区 As Excel.Range
        Dim 校验数据区2 As Excel.Range

        If 校验区域 Is Nothing Then
            Exit Sub
        End If
        If 标题列所在行 >= 0 Then
            If 校验区域.Rows.Count - 标题列所在行 > 0 Then
                校验数据区 = 校验区域.Cells(标题列所在行 + 1, 1).resize(校验区域.Rows.Count - 标题列所在行, 1)
                If 校验区域2 IsNot Nothing Then
                    校验数据区2 = 校验区域2.Cells(标题列所在行 + 1, 1).resize(校验区域2.Rows.Count - 标题列所在行, 1)
                End If
            Else
                Exit Sub
            End If

        Else
            Exit Sub
        End If
        Dim 校验结果列 As Excel.Range

        If 校验区域2 Is Nothing Then
            校验结果列 = 插入列(校验区域.Worksheet, 校验区域.Column + 1)
        Else
            校验结果列 = 插入列(校验区域.Worksheet, 校验区域2.Column + 1)
        End If

        If 标题列所在行 > 0 Then
            校验结果列.Cells(校验区域(1, 1).row + 标题列所在行 - 1, 1).value = "校验(" & 校验方式.ToString & ")"
        End If
        If 校验数据区.Count > 100000 Then
            Dim 结束行号 As Integer = 获取结束行号(校验区域.Worksheet)
            If 结束行号 - 校验数据区(1, 1).row + 1 > 0 Then
                校验结果列 = 校验结果列.Cells(校验数据区(1, 1).row, 1).resize(结束行号 - 校验数据区(1, 1).row + 1, 1)
            Else
                校验结果列 = 校验结果列.Cells(校验数据区(1, 1).row, 1).resize(1, 1)
            End If

        Else
            校验结果列 = 校验结果列.Cells(校验数据区(1, 1).row, 1).resize(校验数据区.Rows.Count, 1)
        End If

        If 校验区域2 Is Nothing Then
            插入公式(获取校验公式(校验方式, 校验数据区.Cells(1, 1)), 校验结果列.Cells(1, 1))
        Else
            插入公式(获取校验公式(校验方式, 校验数据区.Cells(1, 1), 校验数据区2.Cells(1, 1)), 校验结果列.Cells(1, 1))
        End If



        拖拽填充(校验结果列)

        居中(校验结果列)


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem.ToString = "异同校验" Then
            If 区域选择控件1.区域记录个数 = 0 Then
                区域选择控件1.预设空区域({"校验区域1", "校验区域2"})
            ElseIf 区域选择控件1.区域记录个数 = 1 Then
                区域选择控件1.追加区域("校验区域2", Nothing)
            ElseIf 区域选择控件1.区域记录个数 = 1 Then
                区域选择控件1.追加区域("校验区域2", Nothing)
            ElseIf 区域选择控件1.区域记录个数 > 2 Then
                For i As Integer = 2 To 区域选择控件1.区域记录个数 - 1
                    区域选择控件1.删除区域(2)
                Next
            End If
            区域选择控件1.刷新区域列表()
            区域选择控件1.设置区域名(0, "校验区域1")

            NumericUpDown2.Visible = False
            TextBox3.Visible = False
            Label5.Visible = False

        Else
            For i As Integer = 1 To 区域选择控件1.区域记录个数 - 1
                区域选择控件1.删除区域(1)
            Next


            If ComboBox1.SelectedItem.ToString = "长度校验" Then
                NumericUpDown2.Visible = True
                TextBox3.Visible = False
                Label5.Visible = True

                Label5.Text = "校验位数:"


            ElseIf ComboBox1.SelectedItem.ToString = "包含校验" Then
                NumericUpDown2.Visible = False
                TextBox3.Visible = True
                Label5.Visible = True

                Label5.Text = "包含字符串:"

            ElseIf ComboBox1.SelectedItem.ToString = "不含校验" Then
                NumericUpDown2.Visible = False
                TextBox3.Visible = True
                Label5.Visible = True

                Label5.Text = "不含字符串:"



            ElseIf ComboBox1.SelectedItem.ToString = "开头校验" Then
                NumericUpDown2.Visible = False
                TextBox3.Visible = True
                Label5.Visible = True
                Label5.Text = "开头字符串:"


            ElseIf ComboBox1.SelectedItem.ToString = "结尾校验" Then
                NumericUpDown2.Visible = False
                TextBox3.Visible = True
                Label5.Visible = True
                Label5.Text = "结尾字符串:"
            Else
                NumericUpDown2.Visible = False
                TextBox3.Visible = False
                Label5.Visible = False

            End If







        End If






    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        正确标记 = TextBox1.Text
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        错误标记 = TextBox2.Text
    End Sub

    Private Sub 区域选择控件1_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As Windows.Forms.PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
End Class
