Public Class Form1
    'プロパティの再作成　ボタンをクリックするとボタンのサイズが変更できる============***
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strHeight, strWidge As String
        strHeight = Button1.Size.Height.ToString
        strWidge = Button1.Size.Width.ToString
        TextBox1.Text = "幅" & strHeight & ",高さ" & strWidge & "です"
        Label1.Text = TextBox1.Text
        ' これはプロパティｊなのでできない
        'Button1.Size. = 80
        'Button1.Size = (80, 80)
        Button1.Size = New Size(80, 80)
    End Sub
    'マウスクリックイベントeをメソッドEvSendに渡している====================***
    Private Sub Form1_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseClick
        '　このメソッドの引数にはイベントeを渡します。
        EvSend(e)
    End Sub
    'イベントを渡されたメソッド ======================================***
    Private Sub EvSend(ByRef e As System.Windows.Forms.MouseEventArgs)
        Dim strX, strY As String

        strX = e.X.ToString
        strY = e.Y.ToString
        TextBox1.Text = "X:" & strX & ",Y:" & strY & "です"
    End Sub
    'ByValとByRefの違いをリストボックスで選び、確認する =========================***
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim Arg As Long
        Dim str As String

        str &= "Arg = " & Arg & vbCrLf

        str &= "Call MyProcByVal(Arg)"
        If ComboBox1.SelectedIndex = 0 Then
            Call MyProcByRef(Arg)
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Call MyProcByVal(Arg)
        End If
        str &= "Arg = " & Arg
        Label1.Text = str
    End Sub
    Private Sub MyProcByRef(ByRef Argument As Long)

        Argument = 100
        TextBox1.Text = "MyProcByRef_args:" & Argument

    End Sub
    Private Sub MyProcByVal(ByVal Argument As Long)

        Argument = 100
        TextBox1.Text = "MyProcByVal_args:" & Argument

    End Sub
End Class
