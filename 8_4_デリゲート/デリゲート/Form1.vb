Delegate Sub SumCallMethpd(sender As System.Object, e As System.EventArgs)

Public Class Form1
    'デリゲートはメソッドの外に作成する
    Delegate Sub CallMethod(ByRef str As String)



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim box As TestClass = New TestClass
        Dim dele As CallMethod = New CallMethod(AddressOf box.Method1)
        dele.Invoke("Invokeで出力")
        dele("直接出力")
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        TextBox1.Text = "これは" & e.ToString & "によって押されました。"
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ListBox1.SelectedIndex = 2
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim sum, delA, delB, delC As SumCallMethpd
        delA = New SumCallMethpd(AddressOf Me.Button1_Click)
        delB = New SumCallMethpd(AddressOf Me.Button2_Click)
        delC = New SumCallMethpd(AddressOf Me.Button3_Click)
        sum = delA
        sum.Invoke(sender, e)
        sum = delB
        sum.Invoke(sender, e)
        sum = delC
        sum.Invoke(sender, e)
    End Sub
End Class
'デリゲート用のメソッドを作成したクラス
<Obsolete()>
Public Class TestClass
    '引数はStringのMethodメソッドを作成
    Public Sub Method1(ByRef str As String)
        MsgBox("value=" & str)
    End Sub
End Class
