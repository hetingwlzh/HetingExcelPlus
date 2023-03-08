Public Class 插件设置控件
    Private Sub 插件设置控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Checked = My.Settings.是否剔除重复值
        CheckBox2.Checked = My.Settings.是否忽略空值
        CheckBox3.Checked = My.Settings.是否忽略首尾空白字符

        TextBox1.Text = My.Settings.代表空值的字符串


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        My.Settings.是否剔除重复值 = CheckBox1.Checked
        My.Settings.是否忽略空值 = CheckBox2.Checked
        My.Settings.是否忽略首尾空白字符 = CheckBox3.Checked


        My.Settings.代表空值的字符串 = TextBox1.Text













        My.Settings.Save()

        MsgBox("设置已保存！", MsgBoxStyle.ApplicationModal, "保存设置")
    End Sub
End Class
