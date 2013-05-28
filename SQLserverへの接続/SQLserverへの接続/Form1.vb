Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        ' Windows認証のロジック  ここから*********************************************
        ' 接続に必要な定数を設定=============================================
        'Dim St As String
        'Dim Cn As New System.Data.SqlClient.SqlConnection
        'Dim SQL As System.Data.SqlClient.SqlCommand
        'Dim ServerName As String = "192.168.5.64" 'サーバー名(またはIPアドレス)
        'Dim DatabaseName As String = "Look.PPS" 'データベース

        'St = "Server=""192.168.5.64"";"
        'St &= "integrated security=SSPI;"
        'St &= "initial catalog = Look.PPS"
        'Cn.ConnectionString = St
        'SQL = Cn.CreateCommand
        ' ここまではきめうちで設定しておく

        ' SQLserver認証のロジック　　ここから*******************************************

        ' 接続に必要な定数を設定=============================================
        Dim St As String
        Dim Cn As New System.Data.SqlClient.SqlConnection
        Dim SQL As System.Data.SqlClient.SqlCommand
        Dim ServerName As String = "192.168.5.64" 'サーバー名(またはIPアドレス)
        Dim UserID As String = "looktest" 'ユーザーID
        Dim Password As String = "looktest" 'パスワード
        Dim DatabaseName As String = "Look.PPS" 'データベース
        ' 定数終了==========================================================
        ' 接続文字列を連結する
        St = "Server=" & ServerName & ";"
        St &= "User ID=" & UserID & ";"
        St &= "Password=" & Password & ";"
        St &= "Initial Catalog=" & DatabaseName
        ' 接続文字列をConnectionStringにセットする
        Cn.ConnectionString = St
        SQL = Cn.CreateCommand
        ' ここまではきめうちで設定しておく

        ' SQLの値の部分を置き換えるメソッド　TextBox2に入力された数字で検索に行く
        Dim creater As New SQLcreater
        SQL.CommandText = creater.SearchHinban(TextBox2.Text.ToString)
        'SQL.CommandText = "SELECT 品番 FROM 品番マスタ WHERE 品番ID = 1"
        '　DBアクセスはプログラム以外でのエラー発生の可能性があるので、try句でくるむ。
        Try
            Cn.Open()

            TextBox1.Text = SQL.ExecuteScalar
            ' MsgBox(SQL.ExecuteScalar)
        Catch ex As Exception
            '出力表示のTextBox1にエラーを出力　→余力があれば、log4netを用いてエラーログ出力に変えてみよう
            TextBox1.Text = ex.ToString
        End Try


        Cn.Close()
        SQL.Dispose()
        Cn.Dispose()

    End Sub
    Public Class SQLcreater
        ' SQL基本構文フィールド?には数字（品番ID）が入る
        Dim preStr As String = "SELECT 品番 FROM 品番マスタ WHERE 品番ID = ?"

        Public Function SearchHinban(ByRef strNumber As String)
            Dim number As Integer
            Try
                number = Integer.Parse(strNumber)
            Catch ex As Exception
                number = 1
            End Try
            strNumber = number.ToString()
            preStr = Replace(preStr, "?", strNumber)
            Return preStr
        End Function
    End Class

    
End Class
