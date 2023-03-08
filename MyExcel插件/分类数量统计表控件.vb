Public Class 分类数量统计表控件
    Private Sub 分类数量统计表控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.预设空区域({"行分类区域", "列分类区域1", "列分类区域2", "列分类区域3"})

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim 行分类区域列号, 列分类区域列号1, 列分类区域列号2, 列分类区域列号3 As Integer

        行分类区域列号 = 区域选择控件1.获取区域(0).Column
        列分类区域列号1 = 区域选择控件1.获取区域(1).Column

        If 区域选择控件1.获取区域(2) IsNot Nothing Then
            列分类区域列号2 = 区域选择控件1.获取区域(2).Column
        Else
            列分类区域列号2 = 0
        End If

        If 区域选择控件1.获取区域(3) IsNot Nothing Then
            列分类区域列号3 = 区域选择控件1.获取区域(3).Column
        Else
            列分类区域列号3 = 0
        End If



        生成分类统计表(区域选择控件1.获取区域(0).Worksheet,
                行分类区域列号,
                列分类区域列号1,
                列分类区域列号2,
                列分类区域列号3,
                NumericUpDown1.Value,
                CheckBox1.Checked)

    End Sub

    Private Sub 区域选择控件1_点击选定当前区域按钮(index As Integer, range As Excel.Range) Handles 区域选择控件1.点击选定当前区域按钮
        Dim temp As Excel.Range = range.Worksheet.Cells(NumericUpDown1.Value + 1, 1)
        temp = temp.Resize(MAXROW - NumericUpDown1.Value, MAXCOLUMN)
        区域选择控件1.设置区域(index, app.Intersect(temp, range))
    End Sub
End Class
