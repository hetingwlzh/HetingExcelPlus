Imports System.Drawing

Public Class 循环色控件
    Public PictureList As List(Of Windows.Forms.PictureBox)
    Public LabelList As List(Of Windows.Forms.Label)
    Public 上次的渐变色尾号 As Integer = 0

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
        '设置填充色(RGB()

        For i = 1 To 10
            预设循环色(i, My.Settings.循环色序列(i - 1))
        Next

        'PictureBox1.BackColor = Color.FromArgb(My.Settings.循环色序列(0))
        'PictureBox2.BackColor = Color.FromArgb(My.Settings.循环色序列(1))
        'PictureBox3.BackColor = Color.FromArgb(My.Settings.循环色序列(2))
        'PictureBox4.BackColor = Color.FromArgb(My.Settings.循环色序列(3))
        'PictureBox5.BackColor = Color.FromArgb(My.Settings.循环色序列(4))
        'PictureBox6.BackColor = Color.FromArgb(My.Settings.循环色序列(5))
        'PictureBox7.BackColor = Color.FromArgb(My.Settings.循环色序列(6))
        'PictureBox8.BackColor = Color.FromArgb(My.Settings.循环色序列(7))
        'PictureBox9.BackColor = Color.FromArgb(My.Settings.循环色序列(8))
        'PictureBox10.BackColor = Color.FromArgb(My.Settings.循环色序列(9))


        'PictureBox1.BackColor = Color.FromArgb(255, PictureBox1.BackColor.R, PictureBox1.BackColor.G, PictureBox1.BackColor.B)
        'PictureBox2.BackColor = Color.FromArgb(255, PictureBox2.BackColor.R, PictureBox2.BackColor.G, PictureBox2.BackColor.B)
        'PictureBox3.BackColor = Color.FromArgb(255, PictureBox3.BackColor.R, PictureBox3.BackColor.G, PictureBox3.BackColor.B)
        'PictureBox4.BackColor = Color.FromArgb(255, PictureBox4.BackColor.R, PictureBox4.BackColor.G, PictureBox4.BackColor.B)
        'PictureBox5.BackColor = Color.FromArgb(255, PictureBox5.BackColor.R, PictureBox5.BackColor.G, PictureBox5.BackColor.B)
        'PictureBox6.BackColor = Color.FromArgb(255, PictureBox6.BackColor.R, PictureBox6.BackColor.G, PictureBox6.BackColor.B)
        'PictureBox7.BackColor = Color.FromArgb(255, PictureBox7.BackColor.R, PictureBox7.BackColor.G, PictureBox7.BackColor.B)
        'PictureBox8.BackColor = Color.FromArgb(255, PictureBox8.BackColor.R, PictureBox8.BackColor.G, PictureBox8.BackColor.B)
        'PictureBox9.BackColor = Color.FromArgb(255, PictureBox9.BackColor.R, PictureBox9.BackColor.G, PictureBox9.BackColor.B)
        'PictureBox10.BackColor = Color.FromArgb(255, PictureBox10.BackColor.R, PictureBox10.BackColor.G, PictureBox10.BackColor.B)


        PictureList = New List(Of Windows.Forms.PictureBox) From {
        PictureBox1, PictureBox2,
        PictureBox3, PictureBox4,
        PictureBox5, PictureBox6,
        PictureBox7, PictureBox8,
        PictureBox9, PictureBox10
    }
        LabelList = New List(Of Windows.Forms.Label) From {
        Label1, Label2,
        Label3, Label4,
        Label5, Label6,
        Label7, Label8,
        Label9, Label10
    }
        刷新颜色组件()
    End Sub

    Public Function 预设循环色(颜色序号 As Integer, 颜色 As Integer)
        Dim PBX As Windows.Forms.PictureBox
        Dim R, G, B As Integer
        Dim SetColor As Color
        Select Case 颜色序号
            Case 1
                PBX = PictureBox1
            Case 2
                PBX = PictureBox2
            Case 3
                PBX = PictureBox3
            Case 4
                PBX = PictureBox4
            Case 5
                PBX = PictureBox5
            Case 6
                PBX = PictureBox6
            Case 7
                PBX = PictureBox7
            Case 8
                PBX = PictureBox8
            Case 9
                PBX = PictureBox9
            Case 10
                PBX = PictureBox10
        End Select
        SetColor = Color.FromArgb(颜色)
        R = SetColor.R
        G = SetColor.G
        B = SetColor.B
        PBX.BackColor = Color.FromArgb(255, B, G, R)
    End Function

    Public Function 预设循环色(颜色序号 As Integer, 颜色 As Color)
        Dim PBX As Windows.Forms.PictureBox
        Select Case 颜色序号
            Case 1
                PBX = PictureBox1
            Case 2
                PBX = PictureBox2
            Case 3
                PBX = PictureBox3
            Case 4
                PBX = PictureBox4
            Case 5
                PBX = PictureBox5
            Case 6
                PBX = PictureBox6
            Case 7
                PBX = PictureBox7
            Case 8
                PBX = PictureBox8
            Case 9
                PBX = PictureBox9
            Case 10
                PBX = PictureBox10
        End Select

        PBX.BackColor = 颜色
    End Function



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


        'Dim PictureList As New List(Of Windows.Forms.PictureBox) From {
        'PictureBox1, PictureBox2,
        'PictureBox3, PictureBox4,
        'PictureBox5, PictureBox6,
        'PictureBox7, PictureBox8,
        'PictureBox9, PictureBox10
        '}
        Dim ColorList As New List(Of Integer)
        For i = 0 To 9
            ColorList.Add(RGB(PictureList(i).BackColor.R, PictureList(i).BackColor.G, PictureList(i).BackColor.B))
        Next



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
                    设置填充色(设置区域, ColorList(colorNum Mod 颜色数))

                End If
                colorNum += 1
                当前分类标志 = tempStr
                涂色起始行 = cell.Row
                涂色结束行 = 涂色起始行
            End If
        Next
        设置区域 = app.Intersect(涂色区域, sheet.Range(sheet.Cells(涂色起始行, 1), sheet.Cells(涂色结束行, 获取结束单元格(sheet).Column)))
        If 设置区域 IsNot Nothing Then
            设置填充色(设置区域, ColorList(colorNum Mod 颜色数))
        End If

    End Sub














    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        插件设置()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MsgBox("这里还在施工
请绕行"， MsgBoxStyle.ApplicationModal, "功能说明")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i = 1 To 10
            预设循环色(i, My.Settings.循环色序列(i - 1))
        Next
    End Sub


    Public Function 刷新颜色组件()
        Try
            For i = 1 To 10
                If i <= NumericUpDown1.Value Then
                    PictureList(i - 1).Visible = True
                    LabelList(i - 1).Visible = True
                Else
                    PictureList(i - 1).Visible = False
                    LabelList(i - 1).Visible = False
                End If

            Next

            If CheckBox1.Checked = True And 上次的渐变色尾号 > 1 Then
                PictureList(NumericUpDown1.Value - 1).BackColor = PictureList(上次的渐变色尾号 - 1).BackColor
                设置渐变(PictureBox1.BackColor, PictureList(NumericUpDown1.Value - 1).BackColor, NumericUpDown1.Value)
            Else


            End If

            上次的渐变色尾号 = NumericUpDown1.Value

            GroupBox1.Height = 80 + NumericUpDown1.Value * 32


        Catch ex As Exception

        End Try
    End Function


    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        If PictureList IsNot Nothing Then
            刷新颜色组件()
        End If

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click, PictureBox9.Click, PictureBox8.Click, PictureBox7.Click, PictureBox6.Click, PictureBox5.Click, PictureBox4.Click, PictureBox3.Click, PictureBox2.Click, PictureBox10.Click
        If CheckBox1.Checked = False Or sender Is PictureBox1 Or sender Is PictureList(NumericUpDown1.Value - 1) Then
            Dim ColorList(9) As Integer
            For i = 0 To 9
                ColorList(i) = (RGB(PictureList(i).BackColor.R, PictureList(i).BackColor.G, PictureList(i).BackColor.B))
            Next

            ColorDialog1.CustomColors = ColorList

            If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                sender.BackColor = ColorDialog1.Color
            End If




        End If
            If CheckBox1.Checked = True Then
            设置渐变(PictureBox1.BackColor, PictureList(NumericUpDown1.Value - 1).BackColor, NumericUpDown1.Value)
        End If
    End Sub


    ''' <summary>
    ''' 让pictureBox控件实现渐变色
    ''' </summary>
    ''' <param name="color1">开始的颜色</param>
    ''' <param name="color2">结束的颜色</param>
    ''' <param name="num">结束颜色的序号，开始的为1</param>
    Public Sub 设置渐变(color1 As Color, color2 As Color, num As Integer)
        If NumericUpDown1.Value >= 3 Then

            Dim PB1, PB2 As Windows.Forms.PictureBox
            PB1 = PictureBox1
            PB2 = PictureList(num - 1)

            PB1.BackColor = color1
            PB2.BackColor = color2

            Dim R1, G1, B1, R2, G2, B2, dR, dG, dB As Integer
            R1 = PB1.BackColor.R
            G1 = PB1.BackColor.G
            B1 = PB1.BackColor.B

            R2 = PB2.BackColor.R
            G2 = PB2.BackColor.G
            B2 = PB2.BackColor.B
            dR = (R2 - R1) / (NumericUpDown1.Value - 1)
            Dim a, b, c As Integer
            a = G2 - G1
            b = NumericUpDown1.Value - 1
            dG = (G2 - G1) / (NumericUpDown1.Value - 1)
            dB = (B2 - B1) / (NumericUpDown1.Value - 1)

            Dim currentColor As Color = color1
            For i = 2 To num - 1
                currentColor = Color.FromArgb(currentColor.R + dR, currentColor.G + dG, currentColor.B + dB)
                PictureList(i - 1).BackColor = currentColor
            Next


        End If

    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        If CheckBox1.Checked = True Then
            设置渐变(PictureBox1.BackColor, PictureList(NumericUpDown1.Value - 1).BackColor, NumericUpDown1.Value)
        End If
    End Sub

    Private Sub 区域选择控件1_Load(sender As Object, e As EventArgs) Handles 区域选择控件1.Load

    End Sub
End Class
