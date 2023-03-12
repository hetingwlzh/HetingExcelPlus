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
