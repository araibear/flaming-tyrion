Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strHeight, strWidge As String
        strHeight = Button1.Size.Height.ToString
        strWidge = Button1.Size.Width.ToString
        TextBox1.Text = "幅" & strHeight & ",高さ" & strWidge & "です"
        ' これはプロパティｊにsetがないのでできない
        'Button1.Size. = 80
        'Button1.Size.Width = 80
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
End Class
