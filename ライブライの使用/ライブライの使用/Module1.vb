Module Module1

    Sub Main()
        Dim Cl1 As New ClassLibrary1.Class1
        Dim str As String

        str = Cl1.TestLib("aa")
        System.Console.WriteLine(str)
    End Sub

End Module
