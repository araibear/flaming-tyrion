Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim attr As A = New A()
        attr.MethA()
        Label1.Text = attr.ToString()
    End Sub
End Class

Public Class TestAttribute
    Inherits System.Attribute
    Public Overrides Function ToString() As System.String
        Return "属性ですよ！属性！"
    End Function
End Class
<Test()>
Public Class A
    Public Sub MethA()

    End Sub
End Class
