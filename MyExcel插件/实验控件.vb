
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Threading '导入命名空间
Imports System.Speech.Synthesis
Public Class 实验控件


    Public t As Thread '定义全局线程变量

    Public words As String = ""
    Public 语速 As Integer = 0




    Public Function 检索字符串(str As String) As String()
        '定义一个字符串，包含中英文混合


        '定义一个用于存储英文单词的数组
        Dim words As String()

        '使用正则表达式，识别出字符串中的所有英文单词
        words = Regex.Matches(str, "\b[a-zA-Z]+\b").Cast(Of Match).Select(Function(m) m.Value).ToArray()

        '输出英文单词
        Return words
    End Function

    Public Function 检索字符串2(str As String) As String()
        '定义一个字符串，包含中英文混合


        '定义一个用于存储英文单词的数组
        Dim words As String()

        '使用正则表达式，识别出字符串中的所有英文单词
        words = Regex.Matches(str, "\b[a-zA-Z]+\b").Cast(Of Match).Select(Function(m) m.Value).ToArray()

        '输出英文单词
        Return words
    End Function
    Public Function SayWord(ByVal word As String,
                            Optional 语速 As Integer = 0,
                            Optional 是否女生 As Boolean = True,
                            Optional 年龄 As Boolean = True) As String
        Dim synth As New SpeechSynthesizer
        synth.Rate = 语速
        Dim age As VoiceAge
        age = VoiceAge.Senior

        If 是否女生 = True Then

            synth.SelectVoiceByHints(VoiceGender.Female, age)
        Else
            synth.SelectVoiceByHints(VoiceGender.Male, age)
        End If
        'synth.SelectVoice(ComboBox1.Text)

        synth.Speak(word)

        Return synth.ToString
    End Function





    Private Sub 实验控件_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim synth As New SpeechSynthesizer
        语速 = NumericUpDown1.Value
        For Each voice As System.Speech.Synthesis.InstalledVoice In synth.GetInstalledVoices
            ComboBox1.Items.Add(voice.VoiceInfo.Name)
        Next

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        朗读(TextBox1.Text, NumericUpDown1.Value)
    End Sub
    Public Sub 朗读(text As String, Optional 速度 As Integer = 0)
        words = text
        语速 = 速度
        t = New Thread(AddressOf SayWord2) '创建线程，使它指向test过程，注意该过程不能带有参数
        t.Start() '启动线程
        'TextBox2.Text = SayWord(TextBox1.Text, NumericUpDown1.Value, False, NumericUpDown2.Value)

    End Sub
    Public Function SayWord2() As String
        Dim synth As New SpeechSynthesizer
        synth.Rate = 语速

        synth.Speak(words)

        Return synth.ToString
    End Function



    Private Sub DataCell1_点击清除按钮()
        MsgBox("88")
    End Sub

    Private Sub Timer2_改变定数间隔(间隔 As Integer)
        MsgBox(间隔)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim aa As New HetingControl.Timer
        TabPage1.Controls.Add(aa)
        aa.Top = 100
        aa.Left = 100
    End Sub





























End Class
