Public Class Form1

    Delegate Sub CallMethod(ByRef str As String)


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim box As TestClass = New TestClass
        Dim dele As CallMethod = New CallMethod(AddressOf box.Method)
        dele.Invoke("あなたの魂に安らぎアレ")
        dele("帝王の殻")
    End Sub
End Class

<Obsolete()>
Public Class TestClass
    Public Sub Method(ByRef str As String)
        MsgBox("value=" & str)
    End Sub
End Class
