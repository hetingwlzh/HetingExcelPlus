Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.CheckBox1 = Me.Factory.CreateRibbonCheckBox
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.Box3 = Me.Factory.CreateRibbonBox
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.Group8 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.CheckBox3 = Me.Factory.CreateRibbonCheckBox
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.SplitButton2 = Me.Factory.CreateRibbonSplitButton
        Me.CheckBox2 = Me.Factory.CreateRibbonCheckBox
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group7 = Me.Factory.CreateRibbonGroup
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button47 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button38 = Me.Factory.CreateRibbonButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button34 = Me.Factory.CreateRibbonButton
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.Button48 = Me.Factory.CreateRibbonButton
        Me.Button31 = Me.Factory.CreateRibbonButton
        Me.ToggleButton1 = Me.Factory.CreateRibbonToggleButton
        Me.Button36 = Me.Factory.CreateRibbonButton
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Button49 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Button41 = Me.Factory.CreateRibbonButton
        Me.Button42 = Me.Factory.CreateRibbonButton
        Me.Button40 = Me.Factory.CreateRibbonButton
        Me.Button39 = Me.Factory.CreateRibbonButton
        Me.Button46 = Me.Factory.CreateRibbonButton
        Me.Button44 = Me.Factory.CreateRibbonButton
        Me.Button45 = Me.Factory.CreateRibbonButton
        Me.Button43 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button32 = Me.Factory.CreateRibbonButton
        Me.Button33 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button50 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Gallery1 = Me.Factory.CreateRibbonGallery
        Me.我的设置 = Me.Factory.CreateRibbonButton
        Me.颜色 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button35 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.ToggleButton3 = Me.Factory.CreateRibbonToggleButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button37 = Me.Factory.CreateRibbonButton
        Me.Gallery2 = Me.Factory.CreateRibbonGallery
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.ToggleButton2 = Me.Factory.CreateRibbonToggleButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Box2.SuspendLayout()
        Me.Box3.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.Group8.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.Group7.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Groups.Add(Me.Group8)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "heting"
        Me.Tab1.Name = "Tab1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Box1)
        Me.Group2.Items.Add(Me.Box2)
        Me.Group2.Items.Add(Me.Box3)
        Me.Group2.Items.Add(Me.Button4)
        Me.Group2.Items.Add(Me.Button5)
        Me.Group2.Items.Add(Me.Button13)
        Me.Group2.Label = "定位选择"
        Me.Group2.Name = "Group2"
        '
        'Box1
        '
        Me.Box1.Items.Add(Me.Button1)
        Me.Box1.Items.Add(Me.Button2)
        Me.Box1.Items.Add(Me.Button3)
        Me.Box1.Items.Add(Me.CheckBox1)
        Me.Box1.Name = "Box1"
        '
        'CheckBox1
        '
        Me.CheckBox1.Label = "数据边"
        Me.CheckBox1.Name = "CheckBox1"
        '
        'Box2
        '
        Me.Box2.Items.Add(Me.Button41)
        Me.Box2.Items.Add(Me.Button42)
        Me.Box2.Items.Add(Me.Button40)
        Me.Box2.Items.Add(Me.Button39)
        Me.Box2.Name = "Box2"
        '
        'Box3
        '
        Me.Box3.Items.Add(Me.Button46)
        Me.Box3.Items.Add(Me.Button44)
        Me.Box3.Items.Add(Me.Button45)
        Me.Box3.Items.Add(Me.Button43)
        Me.Box3.Name = "Box3"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.Button15)
        Me.Group6.Items.Add(Me.Button32)
        Me.Group6.Items.Add(Me.Button33)
        Me.Group6.Items.Add(Me.Button14)
        Me.Group6.Items.Add(Me.Button22)
        Me.Group6.Items.Add(Me.Button50)
        Me.Group6.Items.Add(Me.Button20)
        Me.Group6.Items.Add(Me.Button21)
        Me.Group6.Items.Add(Me.Button16)
        Me.Group6.Items.Add(Me.Gallery1)
        Me.Group6.Label = "综合操作"
        Me.Group6.Name = "Group6"
        '
        'Group8
        '
        Me.Group8.Items.Add(Me.Button25)
        Me.Group8.Items.Add(Me.Button35)
        Me.Group8.Items.Add(Me.Button11)
        Me.Group8.Items.Add(Me.Button17)
        Me.Group8.Label = "统计"
        Me.Group8.Name = "Group8"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button23)
        Me.Group4.Items.Add(Me.ToggleButton3)
        Me.Group4.Items.Add(Me.Button19)
        Me.Group4.Items.Add(Me.Button29)
        Me.Group4.Items.Add(Me.Button37)
        Me.Group4.Items.Add(Me.Gallery2)
        Me.Group4.Items.Add(Me.ToggleButton2)
        Me.Group4.Items.Add(Me.CheckBox3)
        Me.Group4.Label = "校验 生成"
        Me.Group4.Name = "Group4"
        '
        'CheckBox3
        '
        Me.CheckBox3.Label = "相同值"
        Me.CheckBox3.Name = "CheckBox3"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button26)
        Me.Group5.Items.Add(Me.Button12)
        Me.Group5.Items.Add(Me.SplitButton2)
        Me.Group5.Label = "设置"
        Me.Group5.Name = "Group5"
        '
        'SplitButton2
        '
        Me.SplitButton2.Items.Add(Me.CheckBox2)
        Me.SplitButton2.Label = "快速设置"
        Me.SplitButton2.Name = "SplitButton2"
        '
        'CheckBox2
        '
        Me.CheckBox2.Checked = True
        Me.CheckBox2.Label = "启用任务窗格"
        Me.CheckBox2.Name = "CheckBox2"
        '
        'Tab2
        '
        Me.Tab2.Groups.Add(Me.Group7)
        Me.Tab2.Label = "TEST"
        Me.Tab2.Name = "Tab2"
        '
        'Group7
        '
        Me.Group7.Items.Add(Me.Button10)
        Me.Group7.Items.Add(Me.Button47)
        Me.Group7.Items.Add(Me.Button6)
        Me.Group7.Items.Add(Me.Button38)
        Me.Group7.Items.Add(Me.Button24)
        Me.Group7.Items.Add(Me.Button34)
        Me.Group7.Items.Add(Me.SplitButton1)
        Me.Group7.Items.Add(Me.Button48)
        Me.Group7.Items.Add(Me.Button31)
        Me.Group7.Items.Add(Me.ToggleButton1)
        Me.Group7.Items.Add(Me.Button36)
        Me.Group7.Items.Add(Me.Label2)
        Me.Group7.Items.Add(Me.Button49)
        Me.Group7.Label = "Group7"
        Me.Group7.Name = "Group7"
        '
        'Button10
        '
        Me.Button10.Label = "实验控件"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button47
        '
        Me.Button47.Label = "随机抽取"
        Me.Button47.Name = "Button47"
        '
        'Button6
        '
        Me.Button6.Label = "显示当前选区"
        Me.Button6.Name = "Button6"
        '
        'Button38
        '
        Me.Button38.Label = "Button38"
        Me.Button38.Name = "Button38"
        '
        'Button24
        '
        Me.Button24.Label = "Button24"
        Me.Button24.Name = "Button24"
        '
        'Button34
        '
        Me.Button34.Label = "Button34"
        Me.Button34.Name = "Button34"
        '
        'SplitButton1
        '
        Me.SplitButton1.Items.Add(Me.Button30)
        Me.SplitButton1.Label = "SplitButton1"
        Me.SplitButton1.Name = "SplitButton1"
        '
        'Button30
        '
        Me.Button30.Label = "任务窗格"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        '
        'Button48
        '
        Me.Button48.Label = "Button48"
        Me.Button48.Name = "Button48"
        '
        'Button31
        '
        Me.Button31.Label = "Button31"
        Me.Button31.Name = "Button31"
        '
        'ToggleButton1
        '
        Me.ToggleButton1.Label = "编辑选区"
        Me.ToggleButton1.Name = "ToggleButton1"
        '
        'Button36
        '
        Me.Button36.Label = "拖拽填充实验"
        Me.Button36.Name = "Button36"
        '
        'Label2
        '
        Me.Label2.Label = "Label2"
        Me.Label2.Name = "Label2"
        '
        'Button49
        '
        Me.Button49.Label = "获取版本号"
        Me.Button49.Name = "Button49"
        Me.Button49.ShowImage = True
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "数据"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "左上"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "右下"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Button41
        '
        Me.Button41.Image = CType(resources.GetObject("Button41.Image"), System.Drawing.Image)
        Me.Button41.Label = "最左"
        Me.Button41.Name = "Button41"
        Me.Button41.ShowImage = True
        '
        'Button42
        '
        Me.Button42.Image = CType(resources.GetObject("Button42.Image"), System.Drawing.Image)
        Me.Button42.Label = "最右"
        Me.Button42.Name = "Button42"
        Me.Button42.ShowImage = True
        '
        'Button40
        '
        Me.Button40.Image = CType(resources.GetObject("Button40.Image"), System.Drawing.Image)
        Me.Button40.Label = "最上"
        Me.Button40.Name = "Button40"
        Me.Button40.ShowImage = True
        '
        'Button39
        '
        Me.Button39.Image = CType(resources.GetObject("Button39.Image"), System.Drawing.Image)
        Me.Button39.Label = "最下"
        Me.Button39.Name = "Button39"
        Me.Button39.ShowImage = True
        Me.Button39.Tag = ""
        '
        'Button46
        '
        Me.Button46.Image = CType(resources.GetObject("Button46.Image"), System.Drawing.Image)
        Me.Button46.Label = "左边"
        Me.Button46.Name = "Button46"
        Me.Button46.ShowImage = True
        '
        'Button44
        '
        Me.Button44.Image = CType(resources.GetObject("Button44.Image"), System.Drawing.Image)
        Me.Button44.Label = "右边"
        Me.Button44.Name = "Button44"
        Me.Button44.ShowImage = True
        '
        'Button45
        '
        Me.Button45.Image = CType(resources.GetObject("Button45.Image"), System.Drawing.Image)
        Me.Button45.Label = "上边"
        Me.Button45.Name = "Button45"
        Me.Button45.ShowImage = True
        '
        'Button43
        '
        Me.Button43.Image = CType(resources.GetObject("Button43.Image"), System.Drawing.Image)
        Me.Button43.Label = "下边"
        Me.Button43.Name = "Button43"
        Me.Button43.ShowImage = True
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "列宽"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Button5
        '
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.Label = "行高"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button13
        '
        Me.Button13.Image = CType(resources.GetObject("Button13.Image"), System.Drawing.Image)
        Me.Button13.Label = "文本"
        Me.Button13.Name = "Button13"
        Me.Button13.ShowImage = True
        '
        'Button15
        '
        Me.Button15.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button15.Image = CType(resources.GetObject("Button15.Image"), System.Drawing.Image)
        Me.Button15.Label = "枚举"
        Me.Button15.Name = "Button15"
        Me.Button15.ShowImage = True
        '
        'Button32
        '
        Me.Button32.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button32.Image = CType(resources.GetObject("Button32.Image"), System.Drawing.Image)
        Me.Button32.Label = "行匹配"
        Me.Button32.Name = "Button32"
        Me.Button32.ShowImage = True
        '
        'Button33
        '
        Me.Button33.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button33.Image = CType(resources.GetObject("Button33.Image"), System.Drawing.Image)
        Me.Button33.Label = "列汇总"
        Me.Button33.Name = "Button33"
        Me.Button33.ShowImage = True
        '
        'Button14
        '
        Me.Button14.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button14.Image = CType(resources.GetObject("Button14.Image"), System.Drawing.Image)
        Me.Button14.Label = "合并表"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Button22
        '
        Me.Button22.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button22.Image = CType(resources.GetObject("Button22.Image"), System.Drawing.Image)
        Me.Button22.Label = "条件编号"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        '
        'Button50
        '
        Me.Button50.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button50.Image = CType(resources.GetObject("Button50.Image"), System.Drawing.Image)
        Me.Button50.Label = "翻译"
        Me.Button50.Name = "Button50"
        Me.Button50.ShowImage = True
        '
        'Button20
        '
        Me.Button20.Image = CType(resources.GetObject("Button20.Image"), System.Drawing.Image)
        Me.Button20.Label = "编考号"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Button21
        '
        Me.Button21.Image = CType(resources.GetObject("Button21.Image"), System.Drawing.Image)
        Me.Button21.Label = "匹配表头"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button16
        '
        Me.Button16.Image = CType(resources.GetObject("Button16.Image"), System.Drawing.Image)
        Me.Button16.Label = "表间匹配"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'Gallery1
        '
        Me.Gallery1.Buttons.Add(Me.我的设置)
        Me.Gallery1.Buttons.Add(Me.颜色)
        Me.Gallery1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Gallery1.Image = CType(resources.GetObject("Gallery1.Image"), System.Drawing.Image)
        Me.Gallery1.Label = "命令集"
        Me.Gallery1.Name = "Gallery1"
        Me.Gallery1.ShowImage = True
        '
        '我的设置
        '
        Me.我的设置.Label = "设置"
        Me.我的设置.Name = "我的设置"
        '
        '颜色
        '
        Me.颜色.Label = "颜色代码"
        Me.颜色.Name = "颜色"
        '
        'Button25
        '
        Me.Button25.Image = CType(resources.GetObject("Button25.Image"), System.Drawing.Image)
        Me.Button25.Label = "分类包含统计"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Button35
        '
        Me.Button35.Image = CType(resources.GetObject("Button35.Image"), System.Drawing.Image)
        Me.Button35.Label = "分类统计表"
        Me.Button35.Name = "Button35"
        Me.Button35.ShowImage = True
        '
        'Button11
        '
        Me.Button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.Label = "分类计数"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Button17
        '
        Me.Button17.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.Label = "排名"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'ToggleButton3
        '
        Me.ToggleButton3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ToggleButton3.Image = CType(resources.GetObject("ToggleButton3.Image"), System.Drawing.Image)
        Me.ToggleButton3.Label = "信息统计"
        Me.ToggleButton3.Name = "ToggleButton3"
        Me.ToggleButton3.ShowImage = True
        '
        'Button19
        '
        Me.Button19.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button19.Image = CType(resources.GetObject("Button19.Image"), System.Drawing.Image)
        Me.Button19.Label = "循环色"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Button29
        '
        Me.Button29.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button29.Image = CType(resources.GetObject("Button29.Image"), System.Drawing.Image)
        Me.Button29.Label = "随机数"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        '
        'Button37
        '
        Me.Button37.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button37.Image = CType(resources.GetObject("Button37.Image"), System.Drawing.Image)
        Me.Button37.Label = "校验"
        Me.Button37.Name = "Button37"
        Me.Button37.ShowImage = True
        '
        'Gallery2
        '
        Me.Gallery2.Buttons.Add(Me.Button7)
        Me.Gallery2.Buttons.Add(Me.Button8)
        Me.Gallery2.Buttons.Add(Me.Button9)
        Me.Gallery2.Buttons.Add(Me.Button18)
        Me.Gallery2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Gallery2.Image = CType(resources.GetObject("Gallery2.Image"), System.Drawing.Image)
        Me.Gallery2.Label = "文本"
        Me.Gallery2.Name = "Gallery2"
        Me.Gallery2.ShowImage = True
        '
        'Button7
        '
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.Label = "提取数字"
        Me.Button7.Name = "Button7"
        Me.Button7.ShowImage = True
        '
        'Button8
        '
        Me.Button8.Image = CType(resources.GetObject("Button8.Image"), System.Drawing.Image)
        Me.Button8.Label = "逆序"
        Me.Button8.Name = "Button8"
        Me.Button8.ShowImage = True
        '
        'Button9
        '
        Me.Button9.Image = CType(resources.GetObject("Button9.Image"), System.Drawing.Image)
        Me.Button9.Label = "文本拆分"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Image = CType(resources.GetObject("Button18.Image"), System.Drawing.Image)
        Me.Button18.Label = "第n次匹配位置"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'ToggleButton2
        '
        Me.ToggleButton2.Image = CType(resources.GetObject("ToggleButton2.Image"), System.Drawing.Image)
        Me.ToggleButton2.Label = "聚光灯"
        Me.ToggleButton2.Name = "ToggleButton2"
        Me.ToggleButton2.ShowImage = True
        '
        'Button26
        '
        Me.Button26.Image = CType(resources.GetObject("Button26.Image"), System.Drawing.Image)
        Me.Button26.Label = "插件设置"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button12
        '
        Me.Button12.Image = CType(resources.GetObject("Button12.Image"), System.Drawing.Image)
        Me.Button12.Label = "初始化"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'Button23
        '
        Me.Button23.Label = "Button23"
        Me.Button23.Name = "Button23"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.Box3.ResumeLayout(False)
        Me.Box3.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.Group8.ResumeLayout(False)
        Me.Group8.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.Group7.ResumeLayout(False)
        Me.Group7.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group7 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ToggleButton1 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group8 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Gallery1 As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents 我的设置 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents 颜色 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button32 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button33 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button34 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button35 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button36 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button37 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button38 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button43 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button44 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button39 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button40 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button41 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button42 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Button45 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button46 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box3 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents CheckBox1 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Button47 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button31 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button48 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton2 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents CheckBox2 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Button49 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button50 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CheckBox3 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents ToggleButton2 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Gallery2 As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ToggleButton3 As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
