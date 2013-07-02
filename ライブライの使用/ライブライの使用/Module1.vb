Module Module1

    Sub Main()
        Dim Cl1 As New ClassLibrary1.Class1
        Dim o As New Shell32.Shell
        Dim str As String

        str = Cl1.TestMeth("a")
        System.Console.WriteLine(str)
        o.ToggleDesktop()
        o.Open("C:\windows\")


    End Sub

End Module
