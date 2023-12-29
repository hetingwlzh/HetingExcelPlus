Imports MyExcel插件.编号控件

Public Class 合并工作表控件
    Private Sub 合并工作表控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        加载表()
    End Sub
    Public Function 加载表() As Integer
        CheckedListBox1.Items.Clear()
        For Each sheet As Excel.Worksheet In app.Worksheets
            CheckedListBox1.Items.Add(sheet.Name)
        Next
        刷新数据()
        Return app.Worksheets.Count
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim range, tempRange As Excel.Range
        Dim NewSheet, sheet As Excel.Worksheet



        For Each temp In CheckedListBox1.CheckedItems
            sheet = app.Sheets(temp.ToString)
            If 冗余行列检查(sheet, True) = True Then
                Exit Sub
            End If
        Next



        NewSheet = 新建工作表("合并结果", True)




        Dim row As Integer
        row = 1
        Dim n As Integer = 1
        For Each temp In CheckedListBox1.CheckedItems
            sheet = app.Sheets(temp.ToString)
            tempRange = 获取结束单元格(sheet)
            range = MyRange(sheet, 1, 1, tempRange.Row, tempRange.Column)
            range.Copy()
            MyRange(NewSheet, row, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

            '设置填充色(app.Selection, 获取循环色(n Mod 7))
            设置内部边框(app.Selection, 2, 2, RGB(0, 0, 0))
            设置外边框(app.Selection, 2, 4, RGB(0, 0, 0))
            row = 获取结束单元格(NewSheet).Row
            n += 1

            app.DisplayAlerts = False '删除工作表不再提示
            If CheckBox1.Checked = True Then
                sheet.Delete()
            End If
            app.DisplayAlerts = True  '删除工作表回复提示
        Next
        自动列宽(NewSheet)
        加载表()

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Label1.Text = "当前数据表（共" & CheckedListBox1.Items.Count & "个,选中" & CheckedListBox1.CheckedItems.Count & "个）"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        For i = 1 To CheckedListBox1.Items.Count
            CheckedListBox1.SetItemChecked(i - 1, True)
        Next
        刷新数据()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For i = 0 To CheckedListBox1.Items.Count - 1
            CheckedListBox1.SetItemChecked(i, Not CheckedListBox1.GetItemChecked(i))
        Next
        刷新数据()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        加载表()
    End Sub
    Private Sub 刷新数据()
        Label1.Text = "当前数据表（共" & CheckedListBox1.Items.Count & "个,选中" & CheckedListBox1.CheckedItems.Count & "个）"

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        CheckedListBox1.MultiColumn = CheckBox2.Checked
    End Sub
End Class
