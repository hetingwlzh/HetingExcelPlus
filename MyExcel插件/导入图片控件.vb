Imports System.IO

Public Class 导入图片控件
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        InsertPictures()
    End Sub


    Sub InsertPictures()

        '1.选择文件夹
        Dim folderPath As String
        folderPath = BrowseForFolder("请选择文件夹")

        '2.获取文件夹中所有图片路径
        Dim fileNames As Object
        fileNames = GetAllImages(folderPath)
        If fileNames Is Nothing Then
            MsgBox("未找到任何图片文件，请重新选择文件夹。", vbExclamation, "提示")
            Exit Sub
        End If

        '3.插入图片到单元格

        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim i As Long
        For i = 1 To UBound(fileNames)
            'Dim pic As stdole.Picture
            sheet.Cells(i, 1).value = Path.GetFileName(fileNames(i))
            Dim picture As Excel.Shape = sheet.Shapes.AddPicture(fileNames(i), False, True,
                                                                 sheet.Cells(i, 2).Left + 1,
                                                                 sheet.Cells(i, 2).Top + 1,
                                                                 sheet.Cells(i, 2).Width + 1,
                                                                 sheet.Cells(i, 2).Height + 1)
            'pic = wks.Pictures.Insert(fileNames(i))
            'picture.ShapeRange.LockAspectRatio = True
            'picture.Width = wks.Cells(i, 2).Width ' 获取新的宽度
            'picture.Height = wks.Cells(i, 2).Height ' 获取新的高度

            picture.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize

            'app.CommandBars("Format Object").Visible = False
            'Dim newWidth As Single = wks.Cells(1, 2).Width ' 获取新的宽度
            'Dim newHeight As Single = wks.Cells(1, 2).Height ' 获取新的高度
            'picture = picture.Resize(newWidth, newHeight) ' 调整图片大小




            'pic.Top = wks.Cells(i + 1, 2).Top
            'pic.Left = wks.Cells(i + 1, 2).Left
        Next i

    End Sub

    Function BrowseForFolder(prompt As String) As String
        Dim f As Object
        f = app.FileDialog(Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker)
        f.Title = prompt
        f.ButtonName = "选择文件夹"
        If f.Show = -1 Then
            Return f.SelectedItems(1)
        End If
    End Function

    Function GetAllImages(ByVal folderPath As String) As Object
        Dim fileNames As Object
        Dim i As Long
        Dim count As Long
        With CreateObject("Scripting.FileSystemObject")
            If Not .FolderExists(folderPath) Then Exit Function
            Dim folder As Object
            folder = .GetFolder(folderPath)
            ReDim fileNames(folder.Files.count)
            For Each file In folder.Files
                Dim ss As String = LCase(Path.GetExtension(file.Name))
                If ss = ".jpg" Or ss = ".png" Then
                    count = count + 1
                    fileNames(count) = file.Path
                End If
            Next
        End With
        If count > 0 Then ReDim Preserve fileNames(count)
        Return fileNames
    End Function




End Class
