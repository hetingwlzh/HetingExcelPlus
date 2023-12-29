Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class 数据拆分控件
    Dim excelApp As Excel.Application
    Dim FolderPath As String = ""
    Private Sub 数据拆分控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        区域选择控件1.设置锁定用户区(False)
        区域选择控件1.设置锁定行(False)
        区域选择控件1.设置锁定列(False)
        区域选择控件1.设置锁定表(False)
        区域选择控件1.是否允许编辑区域名 = False

        区域选择控件1.预设空区域({"拆分依据列"})

        excelApp = Globals.ThisAddIn.Application
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        数据拆分()
    End Sub

    Public Sub 数据拆分()

        Dim range As Excel.Range = 区域选择控件1.获取区域("拆分依据列")




        If range Is Nothing Then
            MsgBox("请选择 拆分依据列 所在的区域")
            Exit Sub
        End If


        If CheckBox1.Checked = False And CheckBox2.Checked = False Then
            MsgBox("请选择 拆分结果的保存形式 工作表 或 文件")
            Exit Sub
        End If





        Dim sourceSheet As Excel.Worksheet = range.Worksheet
        'Dim sheet As Excel.Worksheet = range.Worksheet

        Dim 拆分依据列号 As Integer = range(1, 1).Column
        '获取要复制的表

        If 冗余行列检查(sourceSheet, True, 20000, 1000) = True Then
            Exit Sub
        End If


        '添加新表格
        'Dim TempSheet As Excel.Worksheet = app.Worksheets.Add(After:=sourceSheet)
        '将指定工作表复制到新的工作簿中
        'sourceSheet.Copy(After:=sourceSheet)
        Dim TempSheet As Excel.Worksheet = 新建工作表("拆分数据临时表", True, False, sourceSheet)


        '复制源表格中的所有单元格
        sourceSheet.Cells.Copy()

        '还可以选择只复制 UsedRange 已使用范围
        'sourceSheet.UsedRange.Copy(newSheet.Cells(1,1))

        '设置单元格格式
        'newSheet.Cells.PasteSpecial(Excel.XlPasteType.xlPasteAll)
        TempSheet.Cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)






        '或者一次性复制格式
        'newSheet.Cells.PasteSpecial(Excel.XlPasteType.xlPasteFormats)

        '重命名新表格
        'TempSheet.Name = "拆分数据临时表"



        Dim 列表头区域 As Excel.Range = MyRange(TempSheet, 1, 1, NumericUpDown1.Value, TempSheet.UsedRange.Columns.Count)

        If TempSheet.UsedRange.Rows.Count <= NumericUpDown1.Value Then
            MsgBox("除了表头，似乎没有可拆分数据！")
            TempSheet.Delete()
            Exit Sub
        End If

        Dim 数据区域 As Excel.Range = MyRange(TempSheet, NumericUpDown1.Value + 1, 1, TempSheet.UsedRange.Rows.Count, TempSheet.UsedRange.Columns.Count)
        If 数据区域 Is Nothing Then
            MsgBox("除了表头，似乎没有可拆分数据！")
            TempSheet.Delete()
            Exit Sub
        End If


        Dim 拆分依据列 As Excel.Range = 数据区域.Columns(拆分依据列号)
        Dim 拆分依据区域 As Excel.Range = 拆分依据列.Cells

        '数据区域.Sort(数据区域.Columns(拆分依据列号))
        数据区域.Sort(Key1:=拆分依据列, Order1:=Excel.XlSortOrder.xlAscending, Orientation:=Excel.XlSortOrientation.xlSortColumns)





        Dim 当前字段 As String = 拆分依据区域(1, 1).Value
        Dim 当前表 As Excel.Worksheet = TempSheet
        Dim 当前起始行 As Integer = NumericUpDown1.Value + 1
        Dim sheetNum As Integer = 0

        excelApp.ScreenUpdating = False '不更新界面，提高运行效率
        app.DisplayAlerts = False '删除工作表不再提示

        Dim SheetNameList As New List(Of String)

        For Each cell As Excel.Range In 拆分依据区域

            If cell.Value <> 当前字段 Then

                '添加新表格


                'Dim newSheet As Excel.Worksheet = app.Worksheets.Add(After:=当前表)
                Dim newSheet As Excel.Worksheet = 新建工作表(当前字段, True, False, After:=当前表)

                sheetNum += 1
                'If 当前字段 <> Nothing Then
                '    Dim ttt As String = 格式化工作表名(当前字段)
                '    newSheet.Name = ttt
                'Else
                '    newSheet.Name = "空" & sheetNum
                'End If
                SheetNameList.Add(newSheet.Name) '记录新建的表
                当前表 = newSheet
                列表头区域.Copy()
                当前表.Cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

                TempSheet.Cells(当前起始行, 1).Resize(cell.Row - 当前起始行, 数据区域.Columns.Count).Copy
                当前表.Cells(NumericUpDown1.Value + 1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

                自动列宽(当前表)


                If CheckBox2.Checked = True Then
                    保存工作表到文件(当前表, FolderPath)
                End If


                当前起始行 = cell.Row
                当前字段 = cell.Value
            End If

        Next

        '添加最后一个新表格
        Dim lastNewSheet As Excel.Worksheet = 新建工作表(当前字段, True, False, After:=当前表)
        sheetNum += 1
        'If 当前字段 <> Nothing Then
        '    lastNewSheet.Name = 格式化工作表名(当前字段)
        'Else
        '    lastNewSheet.Name = "空" & sheetNum
        'End If
        SheetNameList.Add(lastNewSheet.Name) '记录新建的表
        当前表 = lastNewSheet
        列表头区域.Copy()
        当前表.Cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

        TempSheet.Cells(当前起始行, 1).Resize(TempSheet.UsedRange.Rows.Count - 当前起始行 + 1, 数据区域.Columns.Count).Copy
        当前表.Cells(NumericUpDown1.Value + 1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)

        If CheckBox2.Checked = True Then
            保存工作表到文件(当前表, FolderPath)
        End If
        If CheckBox1.Checked = False Then
            app.DisplayAlerts = False '删除工作表不再提示
            For Each SheetName As String In SheetNameList
                app.Sheets(SheetName).Delete()
            Next

        End If


        MsgBox("共拆分成了 " & sheetNum & " 个表格")
        app.DisplayAlerts = False '删除工作表不再提示
        TempSheet.Delete() '删除工作表
        app.DisplayAlerts = True  '恢复删除工作表提示
        excelApp.ScreenUpdating = True

    End Sub

    'Public Function 保存工作表到文件(dataSheet As Excel.Worksheet, filePath As String) As Boolean

    '    '创建不可见的Excel应用程序实例
    '    Dim excelApp As New Excel.Application
    '    excelApp.DisplayAlerts = False '关闭提示弹窗
    '    excelApp.Visible = False '设置为不可见

    '    ''打开源Excel文件
    '    'Dim workbook As Excel.Workbook = excelApp.Workbooks.Open("D:\Test.xlsx")

    '    ''获取要导出的工作表
    '    'Dim sheet As Excel.Worksheet = workbook.Sheets("Sheet1")

    '    '新建一个工作簿,作为导出的目标
    '    Dim newWorkbook As Excel.Workbook = excelApp.Workbooks.Add()

    '    '将源工作表复制到新工作簿
    '    dataSheet.Cells.Copy()
    '    newWorkbook.Sheets(1).cells(1, 1).PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll)
    '    newWorkbook.Sheets(1).name = dataSheet.Name






    '    '保存新工作簿到指定路径
    '    'newWorkbook.SaveAs("D:\NewFile.xlsx")
    '    newWorkbook.SaveAs(Path.Combine(filePath & dataSheet.Name & ".xlsx"))
    '    '关闭新工作簿
    '    newWorkbook.Close()



    '    '退出Excel应用程序
    '    excelApp.Quit()
    '    Return True
    '    Try

    '    Catch ex As Exception
    '        MsgBox("保存文件时出错！")
    '        Return False
    '    End Try
    'End Function





    ''' <summary>
    ''' 将指定工作表导出并保存到文件夹中
    ''' </summary>
    ''' <param name="theSheet">要导出的工作表</param>
    ''' <param name="folderPath">导出的目标文件夹路径</param>
    Public Sub 保存工作表到文件(ByVal theSheet As Excel.Worksheet, ByVal folderPath As String)

        '获取当前Excel应用程序对象
        'excelApp = Globals.ThisAddIn.Application






        '关闭提示弹窗,提高效率
        excelApp.DisplayAlerts = False
        'excelApp.Visible = False '设置为不可见

        '创建新的工作簿,作为导出的容器
        Dim newBook As Excel.Workbook = excelApp.Workbooks.Add()

        '将指定工作表复制到新的工作簿中
        theSheet.Copy(Before:=newBook.Sheets(1))

        '根据工作表名称生成文件名
        Dim fileName As String = theSheet.Name

        '保存工作簿到指定路径  
        newBook.SaveAs(folderPath & "\" & fileName & ".xlsx")

        '关闭工作簿,释放资源
        newBook.Close()

        '恢复提示弹窗设置
        excelApp.DisplayAlerts = True

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            If FolderBrowserDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                FolderPath = FolderBrowserDialog1.SelectedPath
            Else
                FolderPath = ""
            End If

        End If
        Label2.Text = "保存路径：" & FolderPath
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Directory.Exists(FolderPath) Then
            Process.Start("explorer.exe", FolderPath)
        End If
    End Sub

    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.ForeColor = Color.Red
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.ForeColor = Color.Green
    End Sub


End Class
