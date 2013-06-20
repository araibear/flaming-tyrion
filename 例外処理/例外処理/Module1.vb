Module Module1

    Sub Main()
        Dim Answer As String = "初期値"
        Try

            Answer = CDec(Console.ReadLine()) / CDec(Console.ReadLine())

        Catch ex As Exception

            Answer = "!Error:"
            Answer += ex.Message
            Throw ex

        Finally

            Console.WriteLine(Answer)

        End Try

    End Sub

End Module
