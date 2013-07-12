Public Class Form1
    'ボタンクリックイベントをダブルクリックで作ったもの
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim strHello As String = "こんにちは"    '変数strHelloを宣言
        Label1.Text = strHello  'ラベルに変数strHelloの中身を表示
    End Sub
End Class
