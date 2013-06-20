Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strHeight, strWidge As String
        strHeight = Button1.Size.Height.ToString
        strWidge = Button1.Size.Width.ToString
        TextBox1.Text = "幅" & strHeight & ",高さ" & strWidge & "です"
        Label1.Text = TextBox1.Text
        ' これはプロパティｊにsetがないのでできない
        'Button1.Size. = 80
        'Button1.Size = (80, 80)
        Button1.Size = New Size(80, 80)
    End Sub

    Private Sub Form1_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseClick
        '　このメソッドの引数にはイベントeを渡します。
        'Dim strX, strY As String
        EvSend(e)
         '
        '        strX = e.X.ToString
        '        strY = e.Y.ToString
        '        TextBox1.Text = "X:" & strX & ",Y:" & strY & "です"
    End Sub

    Private Sub EvSend(ByRef e As System.Windows.Forms.MouseEventArgs)
        Dim strX, strY As String

        strX = e.X.ToString
        strY = e.Y.ToString
        TextBox1.Text = "X:" & strX & ",Y:" & strY & "です"
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim Arg As Long
        Dim str As String
        'str &= "ByRef:"
        str &= "Arg = " & Arg & vbCrLf

        str &= "Call MyProcByVal(Arg)"
        'Call MyProcByRef(Arg)
        Call MyProcByVal(Arg)
        str &= "Arg = " & Arg
        Label1.Text = str
    End Sub
    Private Sub MyProcByRef(ByRef Argument As Long)

        Argument = 100

    End Sub
    Private Sub MyProcByVal(ByVal Argument As Long)

        Argument = 100
        TextBox1.Text = "MyProcByVal_args:" & Argument

    End Sub
End Class
