Imports System.Drawing
Imports System.IO

Public Class 导入图片控件

    Public folderPath As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        InsertPictures()
    End Sub


    Sub InsertPictures()

        Dim sheet As Excel.Worksheet = app.ActiveSheet
        Dim i As Long = 1
        For Each item As Object In ListBox1.Items
            'Dim pic As stdole.Picture
            sheet.Cells(i, 1).value = item.ToString

            Dim picture As Excel.Shape = sheet.Shapes.AddPicture(folderPath & "\" & item.ToString, False, True,
                                                                 sheet.Cells(i, 2).Left + 1,
                                                                 sheet.Cells(i, 2).Top + 1,
                                                                 sheet.Cells(i, 2).Width + 1,
                                                                 sheet.Cells(i, 2).Height + 1)
            picture.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize
            Label2.Text = "共载入 " & ListBox1.Items.Count & " 个图片"
            i += 1
        Next
        自动列宽(sheet)
        自动行高(sheet)
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

    Function GetAllImages(ByVal folderPath As String) As System.Collections.ArrayList
        Dim fileNames As New System.Collections.ArrayList
        Dim i As Long
        Dim count As Long
        With CreateObject("Scripting.FileSystemObject")
            If Not .FolderExists(folderPath) Then Exit Function
            Dim folder As Object
            folder = .GetFolder(folderPath)

            For Each file In folder.Files
                Dim ss As String = LCase(Path.GetExtension(file.Name))
                If ss = ".jpg" Or ss = ".png" Then
                    count = count + 1
                    fileNames.Add(file.Path)
                End If
            Next
        End With

        Return fileNames
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        folderPath = BrowseForFolder("请选择文件夹")
        TextBox1.Text = folderPath

        加载所有图片(folderPath)

    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        更新当前图片()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        更新当前图片()
    End Sub
    Public Sub 更新当前图片()
        If ListBox1.SelectedItem IsNot Nothing Then
            PictureBox1.ImageLocation = folderPath & "\" & ListBox1.SelectedItem.ToString
            Label2.Text = 获取图片信息(PictureBox1.ImageLocation)
        End If

    End Sub
    Public Function 加载所有图片(picturePath As String) As Integer

        Dim fileNames As New System.Collections.ArrayList
        fileNames = GetAllImages(picturePath)
        If fileNames Is Nothing Then
            MsgBox("未找到任何图片文件，请重新选择文件夹。", vbExclamation, "提示")
            PictureBox1.ImageLocation = Nothing
            ListBox1.Items.Clear()
            Exit Function
        End If
        If fileNames.Count = 0 Then
            MsgBox("未找到任何图片文件，请重新选择文件夹。", vbExclamation, "提示")
            PictureBox1.ImageLocation = Nothing
            ListBox1.Items.Clear()
            Exit Function
        End If
        '3.插入图片到单元格

        ListBox1.Items.Clear()
        Dim i As Long
        For Each file As String In fileNames
            ListBox1.Items.Add(Path.GetFileName(file))
        Next
        ListBox1.SelectedIndex = 0
    End Function
    Public Function 获取图片信息(filePath As String) As String



        '获取图片文件名
        Dim fileName As String = Path.GetFileName(filePath)
        '获取图片分辨率
        Dim image As Drawing.Image = Image.FromFile(filePath)
        Dim 分辨率 As String = image.Width & "×" & image.Height

        Dim imageFormat As Drawing.Imaging.ImageFormat = image.RawFormat

        '获取图片文件大小

        Dim 单位 As String = "KB"
        Dim fileSize As Single = New FileInfo(filePath).Length / 1024
        If fileSize >= 1024 Then
            fileSize = fileSize / 1024
            单位 = "MB"
        End If
        '获取图片文件格式

        Dim formatName As String = ""
        Select Case imageFormat.Guid.ToString().ToUpper()
            Case imageFormat.Jpeg.Guid.ToString().ToUpper()
                formatName = "JPEG"
            Case imageFormat.Bmp.Guid.ToString().ToUpper()
                formatName = "BMP"
            Case imageFormat.Gif.Guid.ToString().ToUpper()
                formatName = "GIF"
            Case imageFormat.Png.Guid.ToString().ToUpper()
                formatName = "PNG"
            Case Else
                formatName = "UNKNOWN"
        End Select
        '将获取的图片信息转换成字符串
        Dim infoString As String = "数量：" & ListBox1.Items.Count & " 个图片" & "当前 " & ListBox1.SelectedIndex + 1 & "/" & ListBox1.Items.Count & vbCrLf &
                                    "文件名：" & fileName & vbCrLf &
                                    "分辨率：" & 分辨率 & vbCrLf &
                                    "文件大小：" & Math.Round(fileSize, 2) & 单位 & vbCrLf &
                                    "文件格式：" & formatName
        '显示图片信息
        Return infoString






    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        加载所有图片(folderPath)
    End Sub
End Class
