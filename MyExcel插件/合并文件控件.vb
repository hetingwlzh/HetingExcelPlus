Imports System.IO
Imports System.Windows.Forms
Imports HetingControl
Imports Microsoft.Office.Interop.Excel

Public Class 合并文件控件
    Dim selectedFiles() As String
    ' 创建一个新的Excel应用程序实例  
    Dim excelApp As New Excel.Application
    Dim workbook As Excel.Workbook
    Dim nn As Integer = 0
    Dim 当前粘贴位置 As Excel.Range
    Dim 当前已合并文件数 As Integer
    Sub OpenAndPasteData(targetCell As Excel.Range,
                         file As String,
                         sheetName As String,
                         Optional 表头尾行 As Integer = 0,
                         Optional 是否粘贴表头 As Boolean = True
                         )
        Try
            '打开指定的Excel文件
            workbook = excelApp.Workbooks.Open(file) ' 替换为实际文件路径  
            Dim worksheet As Excel.Worksheet
            '获取第一个工作表
            If sheetName = Nothing Then
                worksheet = workbook.Sheets(1)
            Else
                worksheet = workbook.Sheets(sheetName)
            End If
            If worksheet Is Nothing Then
                MsgBox(file & vbCrLf & " 中找不到指定的工作表！")
                Exit Sub
            End If

            If 冗余行列检查(worksheet, True) = True Then

                workbook.Close(False)
                Exit Sub
            End If




            ' 获取当前活动工作表  
            'Dim activeWorksheet As Excel.Worksheet = excelApp.ActiveSheet

            '' 获取目标单元格的坐标，这里假设为当前活动工作表的当前位置  
            'Dim targetCell As Excel.Range = app.ActiveSheet.Cells(1, 1)

            ' 复制第一个工作表的所有内容  
            Dim WorkRange As Excel.Range = 获取工作区域(worksheet)
            If 是否粘贴表头 = True Then
                WorkRange.Copy()
            Else
                WorkRange(表头尾行 + 1, 1).Resize(WorkRange.Rows.Count - 表头尾行, WorkRange.Columns.Count).Copy()
            End If

            ' 粘贴到目标单元格
            'targetCell.Worksheet.Paste(targetCell)

            'targetCell.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues)

            targetCell.Select()
            targetCell.Worksheet.Paste()
            当前已合并文件数 += 1

            ListBox1.Items.Remove(Path.GetFileName(file))
            ' 保存更改并关闭工作簿  
            'workbook.Save()
            'workbook.Close()
            ' 退出Excel应用程序
            'excelApp.Quit()
            'app.Wait(1000)

            workbook.Close(False)
        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' 设置过滤器为Excel文件  
        OpenFileDialog1.Filter = "Excel Files (*.xls; *.xlsx)|*.xls;*.xlsx"

        ' 显示对话框并获取用户选择的文件  
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ' 用户选择了一个文件，获取文件路径  
            selectedFiles = OpenFileDialog1.FileNames
            For Each file As String In selectedFiles
                ListBox1.Items.Add(Path.GetFileName(file))
            Next

            ' 在这里处理选择的Excel文件  
            ' ...  


            Label1.Text = "共选择了 " & selectedFiles.Count & " 个文件"

        End If



    End Sub

    Private Sub 合并文件控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        excelApp.DisplayAlerts = False
        excelApp.ScreenUpdating = False '不更新界面，提高运行效率
        Dim sheet As Excel.Worksheet = app.ActiveSheet


        'If nn = 0 Then
        '    当前粘贴位置 = sheet.Cells(1, 1)
        'End If

        'If nn < selectedFiles.Count Then
        '    OpenAndPasteData(当前粘贴位置, selectedFiles(nn), "", NumericUpDown1.Value)
        '    当前粘贴位置 = sheet.Cells(获取结束单元格(sheet).Row + 1, 1)

        '    nn += 1
        'Else
        '    nn = 0
        'End If



        当前已合并文件数 = 0
        当前粘贴位置 = sheet.Cells(1, 1)
        Dim 是否复制表头 As Boolean = True
        For Each file As String In selectedFiles
            OpenAndPasteData(当前粘贴位置, file, "", NumericUpDown1.Value, 是否复制表头)
            当前粘贴位置 = sheet.Cells(获取结束单元格(sheet).Row + 1, 1)
            是否复制表头 = False
        Next

        ' 保存更改并关闭工作簿  
        'workbook.Save()


        ' 退出Excel应用程序  
        excelApp.Quit()

        'excelApp.ScreenUpdating = True  '更新界面，提高运行效率
        excelApp.DisplayAlerts = True
        excelApp.ScreenUpdating = True '不更新界面，提高运行效率
        MsgBox("合并文件操作完毕！共合并了 " & 当前已合并文件数 & " 个文件。")
        'Timer2.Enabled = True
        If selectedFiles.Count > 当前已合并文件数 Then
            Label1.Text = "共选择了 " & selectedFiles.Count & " 个文件，当前已合并 " & 当前已合并文件数 & " 个" & vbCrLf & "还有 " & selectedFiles.Count - 当前已合并文件数 & " 个操作失败！"
        ElseIf selectedFiles.Count = 当前已合并文件数 Then
            Label1.Text = "共选择了 " & selectedFiles.Count & " 个文件，已全部成功合并！"
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox("本功能可以汇总合并列表头相同的多个文件的数据，但是数据中的公式会在合并之后转换成值。")
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub
End Class
