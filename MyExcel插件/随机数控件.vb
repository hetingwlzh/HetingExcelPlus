Public Class 随机数控件
    'Public 随机数区域 As Excel.Range
    Public START, OVER, 最小单位 As Double
    Public 位数 As Integer



    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        NumericUpDown1.DecimalPlaces = NumericUpDown3.Value
        NumericUpDown2.DecimalPlaces = NumericUpDown3.Value
        NumericUpDown1.Increment = 1 / 10 ^ NumericUpDown3.Value
        NumericUpDown2.Increment = 1 / 10 ^ NumericUpDown3.Value
        Label4.Text = 获取范围字符串()
    End Sub

    Private Sub 随机数控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '载入选区()

        区域选择控件1.设置锁定用户区(False)

        区域选择控件1.追加区域("随机数区域", app.Selection)
        'Label3.Text = 获取单元格个数(区域选择控件1.获取区域(0))
        区域选择控件1.是否允许编辑区域名 = False



        START = NumericUpDown1.Value
        OVER = NumericUpDown2.Value
        Label4.Text = 获取范围字符串()




    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged, NumericUpDown2.ValueChanged, CheckBox2.CheckedChanged, CheckBox1.CheckedChanged
        Label4.Text = 获取范围字符串()
        NumericUpDown2.Minimum = NumericUpDown1.Value
        NumericUpDown1.Maximum = NumericUpDown2.Value
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'app.Calculation = Excel.XlCalculation.xlCalculationManual '改为手动计算，避免页面刷新卡顿
        Try
            Dim 随机数区域 As Excel.Range = 区域选择控件1.获取区域(0)
            随机数区域.Worksheet.Activate()
            If 随机数区域 IsNot Nothing Then
                位数 = NumericUpDown3.Value
                最小单位 = 1 / 10 ^ 位数

                If CheckBox1.Checked = True Then
                    START = NumericUpDown1.Value
                Else
                    START = NumericUpDown1.Value + 最小单位
                End If

                If CheckBox2.Checked = True Then
                    OVER = NumericUpDown2.Value
                Else
                    OVER = NumericUpDown2.Value - 最小单位
                End If


                If RadioButton1.Checked = True Then '允许重复值


                    插入公式("=ROUND(RAND()*(" & OVER & "-" & START & ")+" & START & "," & 位数 & ")", 随机数区域.Cells(1, 1))
                    拖拽填充(随机数区域)



                    'For Each Cell As Excel.Range In 随机数区域
                    '    插入公式("=ROUND(RAND()*(" & OVER & "-" & START & ")+" & START & "," & 位数 & ")", Cell)
                    'Next
                    '插入公式("=ROUND(RAND()*(" & OVER & "-" & START & ")+" & START & "," & 位数 & ")", 随机数区域.Cells(1, 1))

                    '随机数区域.Cells(1, 1).AutoFill(Destination:=随机数区域.Cells(1, 2).resize(1, 随机数区域.Columns.Count - 1), Type:=Excel.XlAutoFillType.xlFillDefault)

                    If RadioButton3.Checked = False Then
                        随机数区域.Copy()

                        随机数区域.Select()
                        随机数区域.PasteSpecial(Excel.XlPasteType.xlPasteValues,
                                           Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                                            False, False)
                        app.CutCopyMode = False
                    End If

                Else
                    Dim NumList As New Collections.ArrayList
                    Dim 随机数 As New Random
                    Dim temp As Double
                    'OVER = OVER + 最小单位
                    For i As Integer = 1 To Math.Min(获取单元格个数(随机数区域), 1 + (OVER - START) / 最小单位)
                        temp = Math.Round(随机数.NextDouble() * (OVER - START) + START, 位数)
                        While NumList.Contains(temp)
                            temp = Math.Round(随机数.NextDouble() * (OVER - START) + START, 位数)
                        End While


                        NumList.Add(temp)
                    Next

                    Dim n As Integer = 0
                    For Each Cell As Excel.Range In 随机数区域
                        If n <= NumList.Count - 1 Then
                            Cell.Value = NumList(n)
                            n += 1
                        Else
                            Cell.Value = ""
                        End If
                    Next
                End If

                Dim tempstr As String = "0."
                For i = 1 To 位数
                    tempstr &= "0"
                Next

                If 位数 = 0 Then
                    tempstr = "0"
                End If
                随机数区域.NumberFormatLocal = tempstr
                '自动列宽(随机数区域.Worksheet)
            Else
                区域选择控件1.显示错误()
            End If
        Catch ex As Exception
            MsgBox("区域选择有问题!")
        End Try

        'app.Calculation = Excel.XlCalculation.xlCalculationAutomatic  '改回自动计算，避免页面刷新卡顿
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("这里还在施工,请绕行"， MsgBoxStyle.ApplicationModal, "功能说明")
    End Sub



    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then

            RadioButton3.Enabled = True
            RadioButton4.Enabled = True
        Else
            RadioButton4.Checked = True
            RadioButton3.Enabled = False
            RadioButton4.Enabled = True

        End If
    End Sub

    Public Function 获取范围字符串() As String
        Dim L, R As String
        If CheckBox1.Checked = True Then
            L = "["
        Else
            L = "("
        End If

        If CheckBox2.Checked = True Then
            R = "]"
        Else
            R = ")"
        End If

        Return L & NumericUpDown1.Value.ToString & "," & NumericUpDown2.Value.ToString & R
    End Function


End Class
