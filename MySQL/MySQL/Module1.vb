Imports MySql.Data.MySqlClient

Module Module1

    Sub Main()
        Dim Connection As New MySqlConnection
        Dim Command As MySqlCommand
        Dim DataReader As MySqlDataReader
        '接続文字列を設定  

        Connection.ConnectionString = "Database=データベース名;Data Source=サーバー;User Id=ログインユーザー;Password=パスワード"
        'オープン  
        Connection.Open()
        'コマンド作成  
        Command = Connection.CreateCommand


        'SQL作成  
        Command.CommandText = "SELECT * FROM hoge"

        'データリーダーにデータ取得  
        DataReader = Command.ExecuteReader
        'データを全件出力  
        Do Until Not DataReader.Read
            Debug.Print(DataReader.Item("field01").ToString)

        Loop
        '破棄  
        DataReader.Close()
        Command.Dispose()
        Connection.Close()
        Connection.Dispose()
    End Sub

End Module
