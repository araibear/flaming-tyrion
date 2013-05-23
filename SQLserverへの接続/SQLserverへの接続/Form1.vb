Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        '接続用変数の定義
        'ADO.NETの仕組みを使います。
        Dim strConnectSQL As String
        Dim strSQL As String
        Dim SQLDA As SqlClient.SqlDataAdapter
        Dim SQLDS As New DataSet()

        strConnectSQL = _
          "Server=TEST;" & _
          "Initial Catalog=test;" & _
          "User ID=test_user;" & _
          "Password=test_user;" & _
          "Integrated Security=false"

        'SQL構文を指定します。
        strSQL = "select * from testTable"

        Try
            '接続の設定を行います。
            SQLDA = New SqlClient.SqlDataAdapter(strSQL, strConnectSQL)

            'データセットに格納します。
            SQLDA.Fill(SQLDS, "TEST")

            'データグリッドビューのデータソースを設定
            Me.DataGridView1.DataSource = SQLDS.Tables("TEST")

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try




    End Sub
End Class
