
Imports System.ComponentModel
Imports System.Windows.Forms
Public Class 区域选择控件
    <Browsable(True), Category("heting的属性"), Description("记录已经选定的区域，格式为{{区域名1,Range1,是否锁定行，是否锁定列，是否锁定表，是否是否锁定区域}，...}的序列")>
    Property 区域列表 As New System.Collections.ArrayList
    '区域列表   数据格式：{ {区域名,区域Excel.Range,是否行锁定,是否列锁定,是否表锁定,是否用户区锁定},...}





    <Browsable(True), Category("heting的属性"), Description("作者姓名"), DefaultValue("赫挺")>
    Property heting As String

    <Browsable(True), Category("heting的属性"), Description("记录已经选定的区域，格式为{{区域名1,Range1}，{区域名2,Range2}，...}的序列")>
    Property 区域名集合 As New System.Collections.ArrayList


    <Browsable(True), Category("heting的属性"), Description("是否显示锁定行复选框"), DefaultValue("true")>
    Property 是否显示锁定行 As Boolean


    <Browsable(True), Category("heting的属性"), Description("是否显示锁定列复选框"), DefaultValue("true")>
    Property 是否显示锁定列 As Boolean

    <Browsable(True), Category("heting的属性"), Description("是否显示锁定表复选框"), DefaultValue("true")>
    Property 是否显示锁定表 As Boolean

    <Browsable(True), Category("heting的属性"), Description("是否显示添加区域按钮"), DefaultValue("true")>
    Property 是否显示添加按钮 As Boolean

    <Browsable(True), Category("heting的属性"), Description("是否显示删除区域按钮"), DefaultValue("true")>
    Property 是否显示删除按钮 As Boolean


    <Browsable(True), Category("heting的属性"), Description("是否显示锁定用户区域复选框"), DefaultValue("true")>
    Property 是否显示锁定用户区域 As Boolean


    <Browsable(True), Category("heting的属性"), Description("是否允许编辑区域名"), DefaultValue("true")>
    Property 是否允许编辑区域名 As Boolean


    <Browsable(True), Category("heting的属性"), Description("是否显示错误信息"), DefaultValue("true")>
    Property 是否显示错误信息 As Boolean
    <Browsable(True), Category("heting的属性"), Description("错误信息序列")>
    Property 错误信息序列 As New System.Collections.ArrayList










    Public 区域校验提示信息 As String = ""



    Public Event 区域变化(index As Integer, range As Excel.Range, 操作类型 As String)

    Public Event 点击选定当前区域按钮(index As Integer, range As Excel.Range)



    '<Browsable(True), Category("heting的属性"), Description("是否显示添加区域按钮"), DefaultValue("true")>
    'Property 是否显示添加按钮 As Boolean


    ''' <summary>XML注释</summary>
    ''' 

    <Browsable(False), Category("heting的属性"), Description("已选定区域的个数")>
    Public ReadOnly Property 区域记录个数() As String
        Get
            Return 区域列表.Count
        End Get
        'Set(ByVal value As String)
        '    heting = value
        'End Set
    End Property

    <Browsable(True), Category("heting的属性"), Description("验证是否是单列区域")>
    Property 单列校验 As Boolean

    <Browsable(True), Category("heting的属性"), Description("验证是否是单行区域")>
    Property 单行校验 As Boolean

    <Browsable(True), Category("heting的属性"), Description("验证是否是空区域")>
    Property 空区域校验 As Boolean

    Public Function 记录错误(错误信息 As String) As Integer
        错误信息序列.Add(错误信息)
        Return 错误信息序列.Count
    End Function
    Public Function 错误个数() As Integer
        Return 错误信息序列.Count
    End Function
    Public Sub 清除错误()
        错误信息序列.Clear()
    End Sub
    Public Function 显示错误(Optional index As Integer = -1) As String
        If 错误个数() = 0 Then
            Label1.Text = 生成区域校验信息()
            Return ""
        Else

            Dim ErrorText As String = ""
            If index >= 0 Then
                If index < 错误信息序列.Count Then
                    ErrorText = 错误信息序列(index)
                Else
                    Return ""
                End If

            Else
                ErrorText = 获取错误文本(True)
            End If
            MsgBox(ErrorText, MsgBoxStyle.Exclamation, "错误提醒")
            清除错误()
            Return ErrorText
        End If


    End Function
    Public Function 获取错误文本(Optional 是否分行 As Boolean = False) As String
        Dim n As Integer = 1
        Dim ErrorText As String = ""
        Dim 分行 As String = ""
        If 是否分行 = True Then
            分行 = vbCrLf
        Else
            分行 = ""
        End If
        For Each err As String In 错误信息序列
            ErrorText &= n & "、" & err & " " & 分行
            n += 1
        Next
        Return ErrorText
    End Function
    'Public Class 区域记录
    '    Public 区域名 As String
    '    Public Range As Excel.Range
    '    Public Sub New(_区域名 As String, _Range As Excel.Range)
    '        区域名 = _区域名
    '        Range = _Range
    '    End Sub
    'End Class
    Public Sub 追加区域(区域名 As String,
                    Range As Excel.Range,
                    Optional 行锁定 As Boolean = Nothing,
                    Optional 列锁定 As Boolean = Nothing,
                    Optional 表锁定 As Boolean = Nothing,
                    Optional 用户区锁定 As Boolean = Nothing,
                    Optional 是否刷新区域列表 As Boolean = True)
        If 行锁定 = Nothing Then
            行锁定 = lockRow.Checked
        End If
        If 列锁定 = Nothing Then
            列锁定 = lockColumn.Checked
        End If
        If 表锁定 = Nothing Then
            表锁定 = lockSheet.Checked
        End If
        If 用户区锁定 = Nothing Then
            用户区锁定 = lockUserRange.Checked
        End If



        Dim 记录 As New System.Collections.ArrayList From {区域名, Range, 行锁定, 列锁定, 表锁定, 用户区锁定}
        Dim index As Integer
        index = 区域列表.Add(记录)
        If 是否刷新区域列表 = True Then
            刷新区域列表()
        End If
        RaiseEvent 区域变化(index, 记录, "添加") '引发事件

    End Sub


    Public Sub 预设空区域(区域名序列 As String())
        For Each name As String In 区域名序列
            追加区域(name, Nothing,,,,, False)
        Next
        刷新区域列表()
    End Sub

    Public Sub 插入区域(index As Integer,
                    区域名 As String,
                    Range As Excel.Range,
                    Optional 行锁定 As Boolean = Nothing,
                    Optional 列锁定 As Boolean = Nothing,
                    Optional 表锁定 As Boolean = Nothing,
                    Optional 用户区锁定 As Boolean = Nothing)

        If 行锁定 = Nothing Then
            行锁定 = lockRow.Checked
        End If
        If 列锁定 = Nothing Then
            列锁定 = lockColumn.Checked
        End If
        If 表锁定 = Nothing Then
            表锁定 = lockSheet.Checked
        End If
        If 用户区锁定 = Nothing Then
            用户区锁定 = lockUserRange.Checked
        End If





        Dim 记录 As New System.Collections.ArrayList From {区域名, Range, 行锁定, 列锁定, 表锁定, 用户区锁定}
        区域列表.Insert(index, 记录)
        刷新区域列表()
        RaiseEvent 区域变化(index, 记录, "添加") '引发事件
    End Sub

    Public Sub 删除区域(index As Integer)
        Dim range As Excel.Range = 区域列表(index)(1)
        区域列表.RemoveAt(index)
        刷新区域列表()
        RaiseEvent 区域变化(index, range, "删除") '引发事件
    End Sub


    Public Sub 设置区域(index As Integer, 区域名 As String, Range As Excel.Range)
        If index > 区域记录个数 - 1 Then
            MsgBox("设置区域 错误!超出最大索引！",, "区域选择控件错误")
            Exit Sub
        Else
            区域列表(index)(0) = 区域名
            区域列表(index)(1) = Range
            DataGridView1.Rows(index).Cells(2).Value = 获取区域地址(index)
            RaiseEvent 区域变化(index, Range, "更改") '引发事件
        End If

    End Sub


    Public Sub 设置区域名(index As Integer, 区域名 As String)
        If index > 区域记录个数 - 1 Then
            MsgBox("设置区域 错误!超出最大索引！",, "区域选择控件错误")
            Exit Sub
        Else
            区域列表(index)(0) = 区域名
            DataGridView1.Rows(index).Cells(1).Value = 区域名
        End If

    End Sub
    Public Sub 设置区域(index As Integer, Range As Excel.Range)
        If index > 区域记录个数 - 1 Then
            MsgBox("设置区域 错误!超出最大索引！",, "区域选择控件错误")
            Exit Sub
        Else
            区域列表(index)(1) = Range
            DataGridView1.Rows(index).Cells(2).Value = 获取区域地址(index)
            RaiseEvent 区域变化(index, Range, "更改") '引发事件
        End If

    End Sub
    Public Function 获取区域记录(index As Integer) As System.Collections.ArrayList
        If index > 区域列表.Count - 1 Then
            Return Nothing
        Else
            Return 区域列表(index)
        End If


    End Function
    Public Function 获取区域名(index As Integer) As String
        If index > 区域列表.Count - 1 Then
            Return Nothing
        Else
            Return 区域列表(index)(0)
        End If

    End Function
    Public Function 获取区域(index As Integer) As Excel.Range
        If index > 区域列表.Count - 1 Then
            Return Nothing
        Else
            Return 区域列表(index)(1)
        End If

    End Function



    Public Function 获取区域记录(区域名 As String) As System.Collections.ArrayList
        Return 区域列表(获取区域索引(区域名))
    End Function
    Public Function 获取区域名(区域名 As String) As String
        Return 区域列表(获取区域索引(区域名))(0)
    End Function
    Public Function 获取区域(区域名 As String) As Excel.Range
        Return 区域列表(获取区域索引(区域名))(1)
    End Function


    Public Function 获取区域地址(index As Integer) As String
        Dim 记录 As System.Collections.ArrayList = 区域列表(index)
        Dim range As Excel.Range = 记录(1)


        If range Is Nothing Then
            Return ""
        Else
            Dim sheetName As String = ""
            If 记录(4) = True Then
                sheetName = "'" & range.Worksheet.Name & "'!"
            End If
            Return sheetName & range.Address(记录(2), 记录(3))
        End If

    End Function

    Public Function 获取区域地址(区域名 As String) As String
        Return 获取区域地址(获取区域索引(区域名))
    End Function












    Public Function 获取区域索引(区域名 As String) As Integer
        Dim 记录 As System.Collections.ArrayList
        For i As Integer = 0 To 区域列表.Count - 1
            记录 = 区域列表(i)
            If 记录(0) = 区域名 Then
                Return i
            End If
        Next
        Return -1
    End Function

    Public Function 获取选中区域() As Excel.Range
        If DataGridView1.SelectedCells.Count > 0 Then
            Dim row As Integer = DataGridView1.SelectedCells(0).RowIndex
            Return 区域列表(row)(1)
        Else
            Return Nothing
        End If

    End Function
    Public Function 获取选中区域名() As String
        If DataGridView1.SelectedCells.Count > 0 Then
            Dim row As Integer = DataGridView1.SelectedCells(0).RowIndex
            Return 区域列表(row)(0)
        Else
            Return Nothing
        End If

    End Function
    Public Function 获取选中区索引() As Integer
        If DataGridView1.SelectedCells.Count > 0 Then
            Return DataGridView1.SelectedCells.Item(0).RowIndex
        Else
            Return -1
        End If

    End Function

    Public Function 设置选中区域(range As Excel.Range) As Integer
        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                If 区域列表(index)(5) = True Then '用户区域锁定
                    设置区域(index, app.Intersect(range, range.Worksheet.UsedRange))
                Else
                    设置区域(index, range)
                End If

                If 区域列表(index)(1) IsNot Nothing Then
                    CType(区域列表(index)(1), Excel.Range).Select()
                End If
                Return index
            End If
        Else
            Return -1
        End If

    End Function
    Public Sub 设置选中区域名(区域名 As String)

        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                设置区域名(index, 区域名)
            End If
        End If
    End Sub

    Public Sub 设置选中区域行锁定(Optional 是否锁定行 As Boolean = True)

        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                区域列表(index)(2) = 是否锁定行
                刷新区域列表()
            End If
        End If
    End Sub
    Public Sub 设置选中区域列锁定(Optional 是否锁定列 As Boolean = True)

        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                区域列表(index)(3) = 是否锁定列
                刷新区域列表()
            End If
        End If
    End Sub

    Public Sub 设置选中区域表锁定(Optional 是否锁定表 As Boolean = True)

        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                区域列表(index)(4) = 是否锁定表
                刷新区域列表()
            End If
        End If
    End Sub
    Public Sub 设置选中区域用户区锁定(Optional 是否锁定用户区 As Boolean = True)

        If 区域记录个数 > 0 Then
            Dim index As Integer = 获取选中区索引()
            If index >= 0 Then
                区域列表(index)(5) = 是否锁定用户区
                刷新区域列表()
            End If
        End If
    End Sub



    Public Sub 设置选中行(index As Integer)
        If index >= 0 And index <= DataGridView1.Rows.Count - 1 Then
            DataGridView1.Rows(index).Selected = True
        End If
    End Sub

    Public Sub 刷新区域列表()
        Dim selectRow, selectColumn As Integer
        selectRow = -1
        selectColumn = -1
        If DataGridView1.SelectedCells.Count > 0 Then
            selectRow = DataGridView1.SelectedCells(0).RowIndex
            selectColumn = DataGridView1.SelectedCells(0).ColumnIndex
        Else
            selectRow = 0
            selectColumn = 0
        End If

        DataGridView1.Rows.Clear()
        For Each 记录 As System.Collections.ArrayList In 区域列表
            Dim row As Integer = DataGridView1.Rows.Add()
            'DataGridView1.Rows(row).DefaultCellStyle.BackColor = System.Drawing.Color.Brown
            'DataGridView1.Rows(row).DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Green
            'DataGridView1.Rows(row).DefaultCellStyle.SelectionBackColor = System.Drawing.Color.Yellow

            DataGridView1.Rows(row).Cells(0).Value = row + 1
            DataGridView1.Rows(row).Cells(1).Value = 记录(0)
            DataGridView1.Rows(row).Cells(2).Value = 获取区域地址(row)



            'If 记录(2) = True Then
            '    If 记录(3) = True Then
            '        DataGridView1.Rows(row).Cells(3).Value = "锁定行列"
            '    Else
            '        DataGridView1.Rows(row).Cells(3).Value = "锁定行"
            '    End If
            'Else
            '    If 记录(3) = True Then
            '        DataGridView1.Rows(row).Cells(3).Value = "锁定列"
            '    Else
            '        DataGridView1.Rows(row).Cells(3).Value = "不锁定"
            '    End If
            'End If
            DataGridView1.Rows(row).Cells(3).Value = 记录(2)
            DataGridView1.Rows(row).Cells(4).Value = 记录(3)

            DataGridView1.Rows(row).Cells(5).Value = 记录(4)
            DataGridView1.Rows(row).Cells(6).Value = 记录(5)

        Next



        If selectRow >= 0 And selectRow < DataGridView1.RowCount And selectColumn >= 0 And selectColumn < DataGridView1.ColumnCount Then
            DataGridView1.ClearSelection()
            DataGridView1.Rows(selectRow).Cells(selectColumn).Selected = True
        End If
        DataGridView1.AutoResizeColumns()

    End Sub




    Private Sub 区域选择控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("序号", "序号")
        DataGridView1.Columns.Add("区域名", "区域名")
        DataGridView1.Columns.Add("地址", "地址")

        Dim temp As New DataGridViewCheckBoxColumn()
        temp.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        temp.HeaderText = "锁行"
        temp.Name = "锁行"
        temp.TrueValue = True
        temp.FalseValue = False
        temp.DataPropertyName = "锁行"
        DataGridView1.Columns.Add(temp) '锁定行


        temp = New DataGridViewCheckBoxColumn()
        temp.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        temp.HeaderText = "锁列"
        temp.Name = "锁列"
        temp.TrueValue = True
        temp.FalseValue = False
        temp.DataPropertyName = "锁列"
        DataGridView1.Columns.Add(temp) '锁定列

        temp = New DataGridViewCheckBoxColumn()
        temp.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        temp.HeaderText = "锁表"
        temp.Name = "锁表"
        temp.TrueValue = True
        temp.FalseValue = False
        temp.DataPropertyName = "锁表"
        DataGridView1.Columns.Add(temp) '锁定表

        temp = New DataGridViewCheckBoxColumn()
        temp.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        temp.HeaderText = "锁用户区"
        temp.Name = "锁用户区"
        temp.TrueValue = True
        temp.FalseValue = False
        temp.DataPropertyName = "锁用户区"
        DataGridView1.Columns.Add(temp) '锁定用户区



        DataGridView1.Columns.Item(0).ReadOnly = True
        DataGridView1.Columns.Item(1).ReadOnly = Not 是否允许编辑区域名
        DataGridView1.Columns.Item(2).ReadOnly = False
        DataGridView1.Columns.Item(3).ReadOnly = True
        DataGridView1.Columns.Item(4).ReadOnly = True
        DataGridView1.Columns.Item(5).ReadOnly = True
        DataGridView1.Columns.Item(6).ReadOnly = True



        DataGridView1.RowHeadersWidth = 20
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        DataGridView1.Columns.Item(0).Width = 100
        DataGridView1.Columns.Item(1).Width = 300
        DataGridView1.Columns.Item(2).Width = 700
        DataGridView1.Columns.Item(3).Width = 300
        DataGridView1.Columns.Item(4).Width = 300
        DataGridView1.Columns.Item(5).Width = 300
        DataGridView1.Columns.Item(6).Width = 300

        DataGridView1.Columns.Item(0).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(1).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(2).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(3).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(4).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(5).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序
        DataGridView1.Columns.Item(6).SortMode = Windows.Forms.DataGridViewColumnSortMode.NotSortable '禁止排序



        DataGridView1.DefaultCellStyle.Font = New Drawing.Font("微软雅黑", 10)
        DataGridView1.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        DataGridView1.DefaultCellStyle.BackColor = System.Drawing.Color.White
        DataGridView1.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White
        DataGridView1.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(100, 200, 250) 'System.Drawing.Color.Blue







        '是否显示锁定行 = True
        '是否显示锁定列 = True
        '是否显示锁定表 = True
        '是否显示添加按钮 = True

        lockRow.Visible = 是否显示锁定行
        lockColumn.Visible = 是否显示锁定列
        lockSheet.Visible = 是否显示锁定表
        addButton.Visible = 是否显示添加按钮
        lockUserRange.Visible = 是否显示锁定用户区域
        delButton.Visible = 是否显示删除按钮

        Label1.Text = 生成区域校验信息()



    End Sub
    Public Sub 显示锁定行(Optional 是否显示 As Boolean = True)
        是否显示锁定行 = 是否显示
        界面刷新()
    End Sub

    Public Sub 显示锁定列(Optional 是否显示 As Boolean = True)
        是否显示锁定列 = 是否显示
        界面刷新()
    End Sub
    Public Sub 显示锁定表(Optional 是否显示 As Boolean = True)
        是否显示锁定表 = 是否显示
        界面刷新()
    End Sub
    Public Sub 显示添加按钮(Optional 是否显示 As Boolean = True)
        是否显示添加按钮 = 是否显示
        界面刷新()
    End Sub
    Public Sub 显示锁定用户区(Optional 是否显示 As Boolean = True)
        是否显示锁定用户区域 = 是否显示
        界面刷新()
    End Sub

    Public Sub 设置锁定行(是否锁定行 As Boolean)
        lockRow.Checked = 是否锁定行
    End Sub
    Public Sub 设置锁定列(是否锁定列 As Boolean)
        lockColumn.Checked = 是否锁定列
    End Sub
    Public Sub 设置锁定表(是否锁定表 As Boolean)
        lockSheet.Checked = 是否锁定表
    End Sub
    Public Sub 设置锁定用户区(是否锁定用户区 As Boolean)
        lockUserRange.Checked = 是否锁定用户区
    End Sub

    Public Sub 界面刷新()
        lockRow.Visible = 是否显示锁定行
        lockColumn.Visible = 是否显示锁定列
        lockSheet.Visible = 是否显示锁定表
        addButton.Visible = 是否显示添加按钮
        lockUserRange.Visible = 是否显示锁定用户区域
    End Sub
    Private Sub SelectButton_Click(sender As Object, e As EventArgs) Handles SelectButton.Click
        If DataGridView1.SelectedCells.Count > 0 Then
            Dim index As Integer = 设置选中区域(app.Selection)

            If DataGridView1.SelectedCells(0).RowIndex < DataGridView1.Rows.Count - 1 Then
                'DataGridView1.Rows(DataGridView1.SelectedCells(0).RowIndex + 1).Cells(2).Selected = True
                DataGridView1.CurrentCell = DataGridView1.Rows(DataGridView1.SelectedCells(0).RowIndex + 1).Cells(2)
            End If
            区域校验()
            RaiseEvent 点击选定当前区域按钮(index, app.Selection)
        End If

    End Sub
    Public Function 获取当前表中激活区域() As Excel.Range
        Return app.Selection

    End Function

    Public Function 区域校验(index As Integer,
                         Optional 是否空区域校验 As Boolean = Nothing,
                         Optional 是否单行校验 As Boolean = Nothing,
                         Optional 是否单列校验 As Boolean = Nothing) As Boolean
        Try



            Dim range As Excel.Range = 获取区域(index)


            If 是否空区域校验 = Nothing Then
                是否空区域校验 = 空区域校验
            End If

            If 是否单行校验 = Nothing Then
                是否单行校验 = 单行校验
            End If

            If 是否单列校验 = Nothing Then
                是否单列校验 = 单列校验
            End If



            If range Is Nothing Then
                If 是否空区域校验 = True Then
                    设置警告色(index, True)
                    记录错误("第" & index + 1 & "个区域:" & 区域列表(index)(0) & " 不能为空！")
                    Return False
                Else
                    Return True
                End If
            Else
                If 是否单列校验 = True And range.Columns.Count <> 1 Then
                    设置警告色(index, True)
                    记录错误("第" & index + 1 & "个区域:" & 区域列表(index)(0) & " 必须为单列！")
                    Return False
                ElseIf 是否单行校验 = True And range.Rows.Count <> 1 Then
                    设置警告色(index, True)
                    记录错误("第" & index + 1 & "个区域:" & 区域列表(index)(0) & " 必须为单行！")
                    Return False
                Else
                    设置警告色(index, False)
                    Return True
                End If
            End If


        Catch ex As Exception
            设置警告色(index, True)
            记录错误("区域校验：区域设置时出现意外错误！")
            Return False
        End Try

    End Function
    Public Function 生成区域校验信息()
        区域校验提示信息 = "区域验证："
        If 单行校验 = True Then
            区域校验提示信息 &= " [单行区域] "
        End If
        If 单列校验 = True Then
            区域校验提示信息 &= " [单列区域] "
        End If
        If 空区域校验 = True Then
            区域校验提示信息 &= " [非空区域] "
        End If
        Return 区域校验提示信息
    End Function
    Public Function 区域校验() As Boolean
        清除错误()
        Dim result As Boolean = True
        For i As Integer = 0 To 区域记录个数 - 1
            If 区域校验(i) = False Then
                result = False
            End If
        Next
        Label1.Text = 获取错误文本(False)
        If Label1.Text = "" Then
            Label1.Text = 生成区域校验信息()
        End If
        If 是否显示错误信息 = True Then
            显示错误()
        End If

        Return result
    End Function

    Public Sub 设置警告色(index As Integer, 是否警告 As Boolean, Optional 列 As Integer = 7)
        If 列 >= 0 And 列 < 7 Then
            If 是否警告 = True Then
                DataGridView1.Rows(index).Cells(列).Style.Font = New Drawing.Font("微软雅黑", 10)
                DataGridView1.Rows(index).Cells(列).Style.ForeColor = System.Drawing.Color.Red
                DataGridView1.Rows(index).Cells(列).Style.BackColor = System.Drawing.Color.Beige
                DataGridView1.Rows(index).Cells(列).Style.SelectionForeColor = System.Drawing.Color.Green
                DataGridView1.Rows(index).Cells(列).Style.SelectionBackColor = System.Drawing.Color.Yellow
            Else
                DataGridView1.Rows(index).Cells(列).Style.Font = New Drawing.Font("微软雅黑", 10)
                DataGridView1.Rows(index).Cells(列).Style.ForeColor = System.Drawing.Color.Black
                DataGridView1.Rows(index).Cells(列).Style.BackColor = System.Drawing.Color.White
                DataGridView1.Rows(index).Cells(列).Style.SelectionForeColor = System.Drawing.Color.White
                DataGridView1.Rows(index).Cells(列).Style.SelectionBackColor = System.Drawing.Color.FromArgb(100, 200, 250) 'System.Drawing.Color.Blue
            End If

        ElseIf 列 >= 7 Then
            For i As Integer = 0 To 6
                设置警告色(index, 是否警告, i)
            Next

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles addButton.Click
        Dim range As Excel.Range
        If lockUserRange.Checked = True Then
            range = app.Intersect(app.Selection, app.ActiveSheet.UsedRange)
        Else
            range = app.Selection
        End If
        追加区域("区域" & 区域记录个数 + 1, range)
        If range IsNot Nothing Then
            range.Select()
        End If

        区域校验()
    End Sub



    'Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
    '    Dim row As Integer = DataGridView1.SelectedCells(0).RowIndex
    '    Dim column As Integer = DataGridView1.SelectedCells(0).ColumnIndex
    '    Dim text As String = DataGridView1.SelectedCells(0).Value

    '    If column = 3 Then '行列锁定列
    '        区域列表(row)(2) = Not 区域列表(row)(2)
    '        DataGridView1.SelectedCells(0).Value = 区域列表(row)(2)
    '        'DataGridView1.SelectedCells(0).Value = 行列锁定集合((行列锁定集合.IndexOf(DataGridView1.SelectedCells(0).Value) + 1) Mod 4)

    '        'If DataGridView1.SelectedCells(0).Value = 行列锁定集合(0) Then
    '        '    区域列表(row)(2) = True
    '        '    区域列表(row)(3) = False
    '        'ElseIf DataGridView1.SelectedCells(0).Value = 行列锁定集合(1) Then
    '        '    区域列表(row)(2) = False
    '        '    区域列表(row)(3) = True
    '        'ElseIf DataGridView1.SelectedCells(0).Value = 行列锁定集合(2) Then
    '        '    区域列表(row)(2) = True
    '        '    区域列表(row)(3) = True
    '        'ElseIf DataGridView1.SelectedCells(0).Value = 行列锁定集合(3) Then
    '        '    区域列表(row)(2) = False
    '        '    区域列表(row)(3) = False
    '        'End If

    '    ElseIf column = 3 Then '表锁定列
    '        区域列表(row)(3) = Not 区域列表(row)(4)
    '        DataGridView1.SelectedCells(0).Value = 区域列表(row)(3)

    '    ElseIf column = 4 Then '表锁定列
    '        区域列表(row)(4) = Not 区域列表(row)(4)
    '        DataGridView1.SelectedCells(0).Value = 区域列表(row)(4)

    '    ElseIf column = 6 Then '用户区锁定列
    '        区域列表(row)(5) = Not 区域列表(row)(5)
    '        DataGridView1.SelectedCells(0).Value = 区域列表(row)(5)
    '    End If
    'End Sub





    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick


        '//不是序号列和标题列时才执行
        If 区域列表.Count > 0 And e.RowIndex < 区域列表.Count Then
            If e.RowIndex <> -1 And e.ColumnIndex <> -1 Then
                Dim dgv As DataGridView = CType(sender, DataGridView)
                Dim value As Boolean




                If DataGridView1.Columns(e.ColumnIndex).GetType.ToString.EndsWith("DataGridViewCheckBoxColumn") Then
                    value = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Not value
                    'DataGridView1.CurrentCell = Nothing '取消编辑状态
                    区域列表(e.RowIndex)(e.ColumnIndex - 1) = Not value
                    'MsgBox(value)
                    If e.ColumnIndex = 6 Then '锁定用户区的列
                        Dim range As Excel.Range = 区域列表(e.RowIndex)(1)
                        If range Is Nothing Then
                            设置区域(e.RowIndex, Nothing)
                        Else
                            设置区域(e.RowIndex, app.Intersect(range, range.Worksheet.UsedRange))
                        End If

                    End If


                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = 获取区域地址(e.RowIndex)
                    'DataGridView1.Rows(row).Cells(3).Value
                End If
                If e.Button = MouseButtons.Right Then
                    If e.ColumnIndex <= 7 Then '序号列，区域名列，区域列
                        Dim temp As Excel.Range = 获取区域(e.RowIndex)
                        If temp IsNot Nothing Then
                            temp.Worksheet.Activate()
                            temp.Select()
                        End If
                    End If
                    DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                End If




            End If

            'Label1.Text = e.RowIndex & "  " & e.ColumnIndex & "  " & e.Button & e.X & "  " & e.Y
        End If



        'MsgBox(DataGridView1.SelectedCells(0).RowIndex & DataGridView1.SelectedCells(0).ColumnIndex & DataGridView1.SelectedCells(0).Value)
    End Sub



    Private Sub DataGridView1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit


        Try
            If 区域列表.Count > 0 And e.RowIndex < 区域列表.Count Then
                If e.ColumnIndex = 2 Then '区域列
                    Dim range As Excel.Range
                    If DataGridView1.CurrentCell.Value = "" Then
                        range = Nothing
                    Else
                        range = app.Range(DataGridView1.CurrentCell.Value)
                    End If

                    'range = app.Intersect(range, app.ActiveSheet.UsedRange)

                    设置区域(e.RowIndex, range)

                    设置警告色(DataGridView1.CurrentCell.RowIndex, False)
                    区域校验(DataGridView1.CurrentCell.RowIndex)
                    If range IsNot Nothing Then
                        range.Select()
                    End If


                ElseIf e.ColumnIndex = 1 Then '区域名列
                    'range = app.Intersect(range, app.ActiveSheet.UsedRange)
                    设置区域名(DataGridView1.CurrentCell.RowIndex, DataGridView1.CurrentCell.Value)

                End If

            End If

        Catch ex As Exception
            设置警告色(DataGridView1.CurrentCell.RowIndex, True)
            '设置区域(DataGridView1.CurrentCell.RowIndex, range)
        End Try
    End Sub

    Private Sub delButton_Click(sender As Object, e As EventArgs) Handles delButton.Click
        Dim index As Integer = 获取选中区索引()
        If index >= 0 Then
            删除区域(index)
        End If
        刷新区域列表()
    End Sub

    ''' <summary>
    ''' 选取当前行所记录的区域
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        If e.ColumnIndex <= 7 Then '序号列，区域名列，区域列
            Dim temp As Excel.Range = 获取区域(e.RowIndex)
            If temp IsNot Nothing Then
                temp.Worksheet.Activate()
                temp.Select()
            End If
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        If DataGridView1.CurrentCell.RowIndex < 区域列表.Count Then
            DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(2).Value = ""
            设置区域(DataGridView1.CurrentCell.RowIndex, Nothing)
        End If


    End Sub
End Class
