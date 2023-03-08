Public Class 分类编号控件
    Public 分类区域1, 分类区域2, 分类区域3, 分类区域4 As Excel.Range
    Public 数据区域 As Excel.Range
    Public color1 As Drawing.Color = Drawing.Color.FromArgb(255, 210, 200)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("这里还在施工
请绕行"， MsgBoxStyle.ApplicationModal, "功能说明")
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub 分类编号控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"分类区域1", "分类区域2", "分类区域3", "分类区域4"})
        区域选择控件1.设置锁定行(False)
        区域选择控件1.设置锁定列(False)
        区域选择控件1.设置锁定表(False)
        区域选择控件1.设置锁定用户区(True)

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If 区域选择控件1.区域校验(0, True, False, True) Then

            分类计数(区域选择控件1.获取区域(0),
                 区域选择控件1.获取区域(1),
                 区域选择控件1.获取区域(2),
                 区域选择控件1.获取区域(3),
                 RadioButton1.Checked，
                 RadioButton3.Checked)
        Else
            区域选择控件1.显示错误()
        End If



        '  =IF(B3="","",COUNTIFS($B$1:B3,B3,$A$1:A3,A3))
    End Sub


End Class
