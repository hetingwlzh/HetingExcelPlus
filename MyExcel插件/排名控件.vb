Public Class 排名控件
    Public 分类区域1, 分类区域2, 分类区域3, 数据区域 As Excel.Range

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sheet As Excel.Worksheet = app.ActiveSheet


        Dim 同数据处理方式 As Integer = 1
        If RadioButton3.Checked = True Then
            同数据处理方式 = 1
        ElseIf RadioButton4.Checked = True Then
            同数据处理方式 = 0
        ElseIf RadioButton5.Checked = True Then
            同数据处理方式 = -1
        End If
        Dim 是否忽略空值 As Boolean
        If RadioButton7.Checked = True Then
            是否忽略空值 = True
        Else
            是否忽略空值 = False
        End If




        排名(区域选择控件1.获取区域(0),
           RadioButton1.Checked,
           同数据处理方式,
           是否忽略空值,
           区域选择控件1.获取区域(1),
           区域选择控件1.获取区域(2),
           区域选择控件1.获取区域(3))

        'Me.Hide()

    End Sub



    Private Sub 排名控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"数据区域", "分类区域1", "分类区域2", "分类区域3"})
    End Sub




End Class
