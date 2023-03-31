Imports System.Drawing
Imports System.Linq.Expressions
Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn



    Private Sub Application_SheetSelectionChange(Sh As Object, Target As Range) Handles Application.SheetSelectionChange


        If 是否突出显示相同值 = True Then

            Dim 剪切板数据 As Object = System.Windows.Forms.Clipboard.GetData(System.Windows.Forms.DataFormats.CommaSeparatedValue)

            刷新选择信息表()
            If 剪切板数据 IsNot Nothing Then
                System.Windows.Forms.Clipboard.SetData(System.Windows.Forms.DataFormats.CommaSeparatedValue, 剪切板数据)
            End If

        End If

        If 是否打开聚光灯效果 = True Then

            Dim 剪切板数据 As Object = System.Windows.Forms.Clipboard.GetData(System.Windows.Forms.DataFormats.CommaSeparatedValue)
            刷新聚光灯效果()
            If 剪切板数据 IsNot Nothing Then
                System.Windows.Forms.Clipboard.SetData(System.Windows.Forms.DataFormats.CommaSeparatedValue, 剪切板数据)
            End If



        End If


        If 是否自动统计当前选择信息 = True And 信息显示控件 IsNot Nothing Then
            Dim temp As 信息统计控件 = 信息显示控件
            temp.设置信息(统计信息())
        End If

















        '' 获取当前活动的工作表对象
        'Dim activeSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)

        '' 尝试获取名为"MyRectangle"的长方形对象，如果存在则删除它


        'Try
        '    Dim oldRectangle As Excel.Shape = activeSheet.Shapes.Item("MyRectangle")
        '    oldRectangle.Delete()
        '    Runtime.InteropServices.Marshal.ReleaseComObject(oldRectangle)
        '    oldRectangle = Nothing
        'Catch ex As Exception

        'End Try

        '' 调用CreateRectangle函数，在所选单元格位置上创建一个新的长方形对象，并显示所选单元格的行和列信息在状态栏上
        'CreateRectangle(activeSheet, Target)




        '' 获取当前活动的工作表对象
        'Dim activeSheet As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)

        '' 尝试获取名为"MyRowRectangle"和"MyColumnRectangle"的长方形对象，如果存在则删除它们
        'Try
        '    Dim oldRowRectange As Excel.Shape = activeSheet.Shapes.Item("MyRowRectangle")
        '    oldRowRectange.Delete()
        '    Runtime.InteropServices.Marshal.ReleaseComObject(oldRowRectange)
        '    oldRowRectange = Nothing

        '    Dim oldColumnRectange As Excel.Shape = activeSheet.Shapes.Item("MyColumnRectangle")
        '    oldColumnRectange.Delete()
        '    Runtime.InteropServices.Marshal.ReleaseComObject(oldColumnRectange)
        '    oldColumnRectange = Nothing

        'Catch ex As Exception

        'End Try

        '' 调用HighlightRowAndColum函数，在所选单元格所在的行和列位置上创建两个新的长方形对象，并显示所选单元格的行和列信息在状态栏上 
        'HighlightRowAndColumn(activeSheet, Target)




    End Sub



    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        'MySelectForm.加载工作表()

        If 是否突出显示相同值 = True And 新建工作表标志 = False Then
            删除关注重复值条件格式(当前表记录, False)
            添加关注相同值条件格式(app.ActiveSheet)

        End If

        If 是否打开聚光灯效果 = True Then
            删除聚光灯效果条件格式(当前表记录)
            开启聚光灯效果(app.ActiveSheet)
        End If


        新建工作表标志 = False
        当前表记录 = app.ActiveSheet
    End Sub

    Private Sub Application_WorkbookOpen(Wb As Workbook) Handles Application.WorkbookOpen
        MySelectForm.加载工作表()
        当前表记录 = app.ActiveSheet
    End Sub

    Private Sub ThisAddIn_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        'taskPane = Me.CustomTaskPanes.Add(New 随机数控件, "任务窗格")

        'taskPane.Visible = True
    End Sub



    Private Sub Application_WindowActivate(Wb As Workbook, Wn As Window) Handles Application.WindowActivate
        当前表记录 = app.ActiveSheet
    End Sub





















End Class
