Public Class Form1
    'ＳＱＬ接続用定数
    Public ServerName As String = "192.168.5.64" 'サーバー名(またはIPアドレス)
    Public UserID As String = "looktest" 'ユーザーID
    Public Password As String = "looktest" 'パスワード
    Public DatabaseName As String = "Look.PPS" 'データベース
    ' 定数終了==========================================================

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
        '検索用ＳＱＬ作成
        SQL.CommandText = creater.SearchKeyID(TextBox2.Text.ToString)
        '更新用ＡＳＱＬ作成
        'SQL.CommandText = creater.UpdateBCard(TextBox2.Lines)
        '追加用ＳＱＬ作成
        SQL.CommandText = creater.InsertBCard(TextBox2.Lines)
        '削除用ＳＱＬ作成
        'SQL.CommandText = creater.DeleteBCard(TextBox2.Text.ToString)
        If SQL.CommandText = "-1" Then
            Exit Sub
        End If
        '更新用ＳＱＬ
        'SQL.CommandText = "UPDATE バースデーカード SET 氏名 = 'プリンタ太郎吉', 敬称フラグ = '9' WHERE ID = 4"
        'データセット取得のＳＱＬ
        'SQL.CommandText = "SELECT * FROM バースデーカード WHERE 性別 = 1"
        '　DBアクセスはプログラム以外でのエラー発生の可能性があるので、try句でくるむ。
        Try
            Cn.Open()
            'TextBox1.Text = SQL.ExecuteScalar
            TextBox1.Text = SQL.ExecuteNonQuery
            'ここからreader()
            'Dim cReader As System.Data.SqlClient.SqlDataReader = SQL.ExecuteReader()
            'While cReader.Read()
            '    TextBox1.Text &= CStr(cReader("氏名"))
            'End While

        Catch ex As Exception
            '出力表示のTextBox1にエラーを出力　→余力があれば、log4netを用いてエラーログ出力に変えてみよう
            TextBox1.Text = ex.ToString
        End Try

        Cn.Close()
        SQL.Dispose()
        Cn.Dispose()

    End Sub

    '*=========================================================================**
    ' クラス名：SQLcreater関数
    ' 概要：バースデーカードの追加・検索・更新・削除用処理クラス
    '*==========================================================================*
    Public Class SQLcreater
        ' SQL基本構文フィールド?には数字（品番ID）が入る
        Dim readPreStr As String = "SELECT ID FROM バースデーカード WHERE ID = ?"
        Dim UpdatePreStr As String = "UPDATE バースデーカード SET" & _
                                        " 氏名 = ?," & _
                                            " 誕生日 = ?," & _
                                            " 敬称フラグ = ?," & _
                                            " 性別 = ?," & _
                                            " 選択リスト区分 = ?," & _
                                            " 本文 = ?," & _
                                            " 更新日付 = ?" & _
                                            " WHERE ID = ?"
        Dim InsertPreStr As String = "INSERT INTO バースデーカード (" & _
                                        "  氏名," & _
                                        "  誕生日," & _
                                        "  敬称フラグ," & _
                                        "  性別," & _
                                        "  選択リスト区分," & _
                                        "  本文," & _
                                        "  更新日付)" & _
                                        "VALUES" & _
                                        "(  ?," & _
                                        "   ?," & _
                                        "   ?," & _
                                        "   ?," & _
                                        "   ?," & _
                                        "   ?," & _
                                        "   ?)"

        Dim DeletePreStr As String = "DELETE FROM バースデーカード WHERE ID = ?"



        '***************************************************************
        ' 機能：バースデーカードの追加
        ' 引数：キー項目氏名
        ' 戻値：ＳＱＬ文字列
        '******1*********2*********3*********4*********5**********6*****
        Public Function InsertBCard(ByRef card() As String)

            Dim strSQL, strBuf1, strBuf2 As String
            Dim sLen As Integer
            strSQL = InsertPreStr
            '入力項目回？の置き換えを行う
            For i = 0 To card.Length - 1
                sLen = InStr(strSQL, "?")
                '？がある場合、手前から置き換えし、なくなったら入力項目数オーバーでエラー
                If (sLen > 0) Then
                    strBuf1 = strSQL.Substring(0, sLen)
                    strBuf2 = strSQL.Substring(sLen)
                    strBuf1 = Replace(strBuf1, "?", "'" & card(i) & "'")
                    strSQL = strBuf1 + strBuf2
                Else
                    MsgBox("入力項目が多すぎます")
                    strSQL = "-1"
                    Exit For
                End If
            Next
            '再度？を確認し、あれば、置き換えする入力項目不足でエラー
            If InStr(strSQL, "?") > 0 Then
                MsgBox("入力項目が足りません")
                strSQL = "-1"
            End If
            Return strSQL
        End Function

        '***************************************************************
        ' 機能：三角形の斜辺を求める関数
        ' 引数：レコード全項目
        ' 戻値：ＳＱＬ文字列
        '******1*********2*********3*********4*********5**********6*****
        Public Function SearchKeyID(ByRef strNumber As String)
            Dim strSQL As String
            Dim number As Integer
            strSQL = readPreStr
            Try
                number = Integer.Parse(strNumber)
            Catch ex As Exception
                number = 1
            End Try
            strNumber = number.ToString()
            strSQL = Replace(strSQL, "?", strNumber)
            Return strSQL
        End Function

        '***************************************************************
        ' 機能：三角形の斜辺を求める関数
        ' 引数：レコード全項目
        ' 戻値：ＳＱＬ文字列
        '******1*********2*********3*********4*********5**********6*****
        Public Function UpdateBCard(ByRef card() As String)
            Dim strSQL, strBuf1, strBuf2 As String
            Dim sLen As Integer
            strSQL = UpdatePreStr
            '入力項目回？の置き換えを行う
            For i = 0 To card.Length - 1
                sLen = InStr(strSQL, "?")
                '？がある場合、手前から置き換えし、なくなったら入力項目数オーバーでエラー
                If (sLen > 0) Then
                    strBuf1 = strSQL.Substring(0, sLen)
                    strBuf2 = strSQL.Substring(sLen)
                    strBuf1 = Replace(strBuf1, "?", "'" & card(i) & "'")
                    strSQL = strBuf1 + strBuf2
                Else
                    MsgBox("入力項目が多すぎます")
                    strSQL = "-1"
                    Exit For
                End If
            Next
            '再度？を確認し、あれば、置き換えする入力項目不足でエラー
            If InStr(strSQL, "?") > 0 Then
                MsgBox("入力項目が足りません")
                strSQL = "-1"
            End If
            Return strSQL
        End Function

        '***************************************************************
        ' 機能：三角形の斜辺を求める関数
        ' 引数：キー項目氏名
        ' 戻値：ＳＱＬ文字列
        '******1*********2*********3*********4*********5**********6*****
        Public Function DeleteBCard(ByRef strNumber As String)
            Dim strSQL As String
            Dim number As Integer
            strSQL = DeletePreStr
            Try
                number = Integer.Parse(strNumber)
            Catch ex As Exception
                number = 1
            End Try
            strNumber = number.ToString()
            strSQL = Replace(strSQL, "?", strNumber)
            Return strSQL

        End Function
    End Class



End Class
