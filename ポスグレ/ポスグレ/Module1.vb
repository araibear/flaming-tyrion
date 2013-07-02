Imports Npgsql
Imports NpgsqlTypes


Module Module1

    Sub Main()
        Dim cn As New NpgsqlConnection
        cn.ConnectionString = "Server = localhost;" & _
                         "Database = hogehoge;" & _
                         "User ID = postgres;" & _
                         "password = xxxx;" & _
                         "Encoding = Unicode"

        Dim cmd As New NpgsqlCommand
        Dim sql As String
        sql = "INSERT INTO Item ("
        sql &= "ItemCode, "
        sql &= "ItemName, "
        sql &= "ItemPrice"
        sql &= ") VALUES ("
        sql &= ":Code, "
        sql &= ":Name, "
        sql &= ":Price"
        sql &= ");"
        Try
            cn.Open()
            Dim trn As NpgsqlTransaction = cn.BeginTransaction
            With cmd
                .Connection = cn
                .CommandText = sql
                .Parameters.Add("Code", NpgsqlDbType.Bigint)
                .Parameters.Add("Name", NpgsqlDbType.Varchar)
                .Parameters.Add("Price", NpgsqlDbType.Numeric)
                .Parameters.Item("Code").Value = 1
                .Parameters.Item("Name").Value = "ぱぱさんのチョコレート"
                .Parameters.Item("Price").Value = 150.55
                .ExecuteNonQuery()
                If .Parameters.Item("Name").Value = "ぱぱさんのチョコレート" Then
                    trn.Commit()
                Else
                    trn.Rollback()
                End If
            End With
        Catch ex As NpgsqlException
            Console.WriteLine(ex.Message)
        Finally
            cn.Close()
            cmd = Nothing
            cn = Nothing
        End Try



    End Sub

End Module
