Public Class Form1
    ReadOnly log As log4net.ILog = _
  log4net.LogManager.GetLogger( _
  System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    'ＤＢ接続情報
    Public ServerName As String = "192.168.5.64" 'サーバー名(またはIPアドレス)
    Public UserID As String = "looktest" 'ユーザーID
    Public Password As String = "looktest" 'パスワード
    Public DatabaseName As String = "Look.PPS" 'データベース
    'ボタンの名称
    Public BTNTITLE As String = "バースデーカードを作る"
    'アプリケーションタイトル
    Public FRMTITLE As String = "バースデーカードエディタ"
    'リストのあるディレクトリパス
    Public LISTPATHE As String = "C:\test"
    '出力ファイルディレクトリパス
    Public OFILEPATHE As String = "C:\letter"
    '出力ファイルディレクトリパス
    Public LOGFILE As String = "C:\letter\log\logfile.txt"
    'IDを大域変数
    Dim statId As Integer = 0
    '出力用エリア
    Dim strDisp As String = ""

    '文字コード
    Public ENC As String = "Shift-JIS"
    '大域ファイルパス
    Public filesBack As String()
    'バースデーカード構造体
    Private Structure BithdayCard
        Public Name As String
        Public BithDay As Date
        Public nTitle As Integer
        Public sex As Integer
        Public txtList As Integer
        Public mainText As String
        Public up_Date As Date
        Public id As Integer
    End Structure
    'ログ構造体
    Private Structure tLog
        Public sysTime As DateTime
        Public sysCall As Integer
        Public level As String
        Public appName As String
        Public mes As String
    End Structure
    '初期化用デリゲート
    Delegate Sub InitCall(sender As System.Object, e As System.EventArgs)

    '***************************************************************
    ' 機能：初期処理
    ' 引数：イベントe、sender
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Text = FRMTITLE
        Button1.Text = BTNTITLE
        ListBox1.Items.Clear()
        Dim files As String()
        Dim i As Integer
        Dim strx As String = e.ToString
        Try
            files = System.IO.Directory.GetFiles( _
    LISTPATHE, "*.txt", System.IO.SearchOption.AllDirectories)
            'ファイルパスをコピーしておく()
            filesBack = files.Clone
            'ファイルのパスと拡張子を取り除く
            Dim strl As Integer = 0
            For i = 0 To files.Length - 1
                strl = InStr(files(i), LISTPATHE) + LISTPATHE.Length
                files(i) = files(i).Substring(strl)
                strl = InStr(files(i), ".txt")
                files(i) = files(i).Substring(0, strl)


            Next
            'ListBox1に結果を表示する
            ListBox1.Items.AddRange(files)
            ListBox1.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(LISTPATHE & "フォルダがありません。")
            ListBox1.Items.Clear()
        End Try

        Button1.Text = BTNTITLE
        TextBoxDay.Text = "//"
        MonthCalendar1.SetDate(Date.Now)
        CheckBox1.Checked = False
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        Label2.Text = "テンプレート"
        Label3.Text = "本文"
        Label4.Text = "誕生日"
        Label5.Text = "初期表示しました"
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.ControlBox = True
        Me.WindowState = FormWindowState.Normal

    End Sub
#Region "ボタンホバーのリージョン==========================================================="
    Private Sub Button1_MouseHover(sender As System.Object, e As System.EventArgs) Handles Button1.MouseHover
        If TextBox1.Text.Equals("") And RichTextBox1.Text.Equals("") Then
            Button1.Text = "データを入力してください。"
        End If
    End Sub

    Private Sub Button1_MouseLeave(sender As System.Object, e As System.EventArgs) Handles Button1.MouseLeave
        Button1.Text = BTNTITLE
    End Sub
#End Region
    '***************************************************************
    ' 機能：入力内容を出力する
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim card As BithdayCard

        '入力項目が満たされているかチェックする
        If (TextBox1.Text.Equals("") = True) Or (TextBox1.Text.Length > 50) Then
            MsgBox("氏名の入力が不正です。")
            Exit Sub
        End If
        Try
            Dim sDate As Date = TextBoxDay.Text.ToString
            MonthCalendar1.SetDate(sDate)
        Catch ex As Exception
            MsgBox("誕生日を正確に入力してください。")
        End Try
        If MonthCalendar1.SelectionRange.End > Date.Now Then
            MsgBox("誕生日が不正です。")
            Exit Sub
        End If
        If (RadioButton1.Checked = False) And (RadioButton2.Checked = False) Then
            GroupBox1.Focus()
            MsgBox("性別を選択してください。")
            Exit Sub
        End If
        If (RichTextBox1.Text.Equals("") = True) Or (RichTextBox1.Text.Length > 1000) Then
            MsgBox("本文の入力は1000文字までにしてください。")
            Exit Sub
        End If

        card = CtrlSet()

        If Rich_to_FileT(OFILEPATHE, card) = 0 Then

            log.Info(OFILEPATHE & "に出力しました。")
        End If

    End Sub

    '***************************************************************
    ' 機能：氏名からバースデーカードテーブルを検索し表示
    ' 引数：ログレベル(level)、ログメッセージ(mes)
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' 接続に必要な定数を設定=============================================
        Dim card As BithdayCard = Nothing
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
        Dim creater As New SQLcreater
        Dim findStr As String = TextBox1.Text
        SQL.CommandText = creater.SearchName(findStr)
        Try
            Cn.Open()
            If SQL.ExecuteScalar > 0 Then
                'ここからreader()
                Dim cReader As System.Data.SqlClient.SqlDataReader = SQL.ExecuteReader()
                While cReader.Read()
                    card.Name = CStr(cReader("氏名"))
                    card.BithDay = cReader("誕生日")
                    card.id = cReader("ID")
                    card.mainText = CStr(cReader("本文"))
                    card.nTitle = CStr(cReader("敬称フラグ"))
                    card.sex = CStr(cReader("性別"))
                    card.txtList = CStr(cReader("選択リスト区分"))
                    card.up_Date = cReader("更新日付")
                End While

            Else
                MsgBox("対象のお名前が見つかりません")
            End If
            strDisp = "ID:" & card.id & "をロードしました。"
        Catch ex As Exception
            '出力表示のTextBox1にエラーを出力　→余力があれば、log4netを用いてエラーログ出力に変えてみよう
            log.Error(ex.ToString)
        End Try
        CtrlGet(card)
        Label5.Text = strDisp
    End Sub
    '***************************************************************
    ' 機能：IDからバースデーカードを1件削除し、初期化
    ' 引数：
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If statId = 0 Then
            MsgBox("このカードは登録されていません。")
            Exit Sub
        End If
        ' 接続に必要な定数を設定=============================================
        Dim card As BithdayCard = Nothing
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
        Dim creater As New SQLcreater
        SQL.CommandText = creater.SearchKeyID(statId)
        Try
            Cn.Open()
            If SQL.ExecuteScalar > 0 Then
                'ここからreader()
                SQL.CommandText = ""
                SQL.CommandText = creater.DeleteBCard(statId)
                strDisp = "ID:"
                strDisp &= SQL.ExecuteScalar()
                strDisp &= statId
                strDisp &= "を正常に削除しました。"
            Else
                MsgBox("対象のお名前が見つかりません")
            End If

        Catch ex As Exception
            '出力表示のTextBox1にエラーを出力　→余力があれば、log4netを用いてエラーログ出力に変えてみよう
            log.Error(ex.ToString)
        End Try
        Label5.Text = strDisp
        statId = 0
        Dim initC As InitCall = New InitCall(AddressOf Me.Form1_Load)
        initC(sender, e)

    End Sub

    '***************************************************************
    ' 機能：カレンダーテキスト入力
    ' 引数：イベントe
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim i As Integer = ListBox1.SelectedIndex
        File_to_RichT(filesBack(i))
    End Sub

    '***************************************************************
    ' 機能：カレンダーテキスト入力
    ' 引数：イベントe
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub TextBoxDay_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBoxDay.TextChanged
        If TextBoxDay.Text.Length > 9 Then
            Try
                Dim sDate As Date = TextBoxDay.Text.ToString
                MonthCalendar1.SetDate(sDate)
            Catch ex As Exception
                MonthCalendar1.SetDate(Date.Now)
            End Try
        End If
    End Sub
    '***************************************************************
    ' 機能：カレンダー日付選択
    ' 引数：カレンダーイベントe
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        TextBoxDay.Text = e.End.ToString("yyyy/MM/dd")
    End Sub
    '***************************************************************
    ' 機能：ログの出力
    ' 引数：ログレベル(level)、ログメッセージ(mes)
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub CtrlGet(ByRef card As BithdayCard)
        Dim i As Integer
        TextBox1.Text = card.Name
        MonthCalendar1.SetDate(card.BithDay)
        If card.nTitle = "0" Then
            CheckBox1.Checked = False
        Else
            CheckBox1.Checked = True
        End If
        If card.sex = "1" Then
            RadioButton1.Checked = True
        ElseIf card.sex = "2" Then
            RadioButton2.Checked = True
        Else
            RadioButton1.Checked = False
            RadioButton2.Checked = False
        End If
        Integer.TryParse(card.txtList, i)
        ListBox1.SelectedIndex = i
        RichTextBox1.Text = card.mainText
        statId = card.id
    End Sub

    '***************************************************************
    ' 機能：各コントロールの情報をカード構造体にセットする
    ' 引数：なし
    ' 戻値：カード構造体BithdayCard
    '******1*********2*********3*********4*********5**********6*****
    Private Function CtrlSet()
        Dim card As BithdayCard
        Dim strMessage() As String
        Dim mes As String
        card.id = statId
        card.Name = TextBox1.Text.ToString
        card.BithDay = MonthCalendar1.SelectionRange.End
        If CheckBox1.Checked Then
            card.nTitle = 1
        Else
            card.nTitle = 0
        End If
        If RadioButton1.Checked Then
            card.sex = 1
        ElseIf RadioButton2.Checked Then
            card.sex = 2
        End If


        card.txtList = ListBox1.SelectedIndex
        card.mainText = RichTextBox1.Text.ToString
        card.up_Date = Date.Now
        If statId > 0 Then
            ReDim strMessage(7)
            strMessage(7) = statId
        Else
            ReDim strMessage(6)
        End If
        strMessage(0) = card.Name.ToString
        strMessage(1) = card.BithDay.ToString("yyyy-MM-dd")
        strMessage(2) = card.nTitle.ToString
        strMessage(3) = card.sex.ToString
        strMessage(4) = card.txtList.ToString
        strMessage(5) = card.mainText.ToString
        strMessage(6) = card.up_Date.ToString("yyyy-MM-dd")


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
        '追加用ＳＱＬ作成
        If statId > 0 Then
            mes = "[更新]"
            SQL.CommandText = creater.UpdateBCard(strMessage)
        Else
            mes = "[登録]"
            SQL.CommandText = creater.InsertBCard(strMessage)
        End If
        If SQL.CommandText = "-1" Then
            Return card
            Exit Function
        End If
        Try
            Cn.Open()
            mes &= SQL.ExecuteNonQuery
            mes &= "件のデータを処理しました。"
            Dim findStr As String = card.Name.ToString
            SQL.CommandText = ""
            SQL.CommandText = creater.SearchName(findStr)
            If SQL.ExecuteScalar > 0 Then
                'ここからreader()
                Dim cReader As System.Data.SqlClient.SqlDataReader = SQL.ExecuteReader()
                While cReader.Read()
                    statId = cReader("ID")
                End While

            Else
                MsgBox("対象のお名前が見つかりません")
            End If
        Catch ex As Exception
            '出力表示のTextBox1にエラーを出力　→余力があれば、log4netを用いてエラーログ出力に変えてみよう
            mes = ex.ToString
        End Try
        strDisp = mes
        Label5.Text = strDisp
        LogCreater("info", mes)

        Return card
    End Function
    '***************************************************************
    ' 機能：文面雛形をRichTextに展開する
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Sub File_to_RichT(ByRef fPath As String)
        Dim Reader As New IO.StreamReader(fPath, System.Text.Encoding.GetEncoding(ENC))
        Dim Line As String = Reader.ReadLine
        RichTextBox1.Text = ""
        Do Until IsNothing(Line)
            RichTextBox1.Text &= Line & vbCrLf
            Line = Reader.ReadLine

        Loop

    End Sub

    '***************************************************************
    ' 機能：バースデーカードの出力
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Function Rich_to_FileT(ByRef fPath As String, ByRef card As BithdayCard)
        Dim Line As String = ""
        '年齢の計算
        Dim sDate, eDate As Date
        Dim age As Integer
        sDate = card.BithDay
        eDate = Date.Now
        age = eDate.Year - sDate.Year



        Dim Writer As New IO.StreamWriter(fPath & "\VBTest.txt", False, System.Text.Encoding.UTF8)
        'お名前
        If card.nTitle = 1 Then
            Line = card.Name & "ちゃん" & "へ"
        Else
            Line = card.Name & "へ"
        End If

        Writer.WriteLine(Line)
        '年齢
        Line = age.ToString & "才のお誕生日おめでとう！"
        Writer.WriteLine(Line)
        Line = ""
        For i = 0 To RichTextBox1.Lines.Length - 1

            Line = RichTextBox1.Lines(i)
            If InStr(Line, "{0}") > 0 Then
                Line = Replace(Line, "{0}", age.ToString)
            End If
            If InStr(Line, "{n}") > 0 Then
                Line = Replace(Line, "{n}", card.Name)
            End If
            Writer.WriteLine(Line)

        Next

        Writer.WriteLine(Date.Now.ToString("yyyy年MM月dd日") & "にて")
        Writer.WriteLine(vbCrLf)
        Writer.WriteLine(vbCrLf)
        Writer.Close()
        Return 0
    End Function


    '***************************************************************
    ' 機能：ログの出力
    ' 引数：ログレベル(level)、ログメッセージ(mes)
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub LogCreater(ByRef level As String, ByRef mes As String)
        Dim Line As String = ""
        Dim log As tLog
        log.appName = My.Application.Info.Title
        log.sysTime = Date.Now
        log.sysCall = 1
        log.level = level
        log.mes = mes

        Dim Writer As New IO.StreamWriter(LOGFILE, True, System.Text.Encoding.UTF8)
        Line &= log.sysTime.ToString & " "
        Line &= "[" & log.sysCall & "] "
        Line &= log.appName & " "
        Line &= log.level & " "
        Line &= log.mes
        Try
            Writer.WriteLine(Line)
        Catch ex As Exception
        End Try
        Writer.Close()
        Debug.WriteLine(Line)
    End Sub

End Class

'*=========================================================================**
' クラス名：SQLcreater関数
' 概要：バースデーカードの追加・検索・更新・削除用処理クラス
'*==========================================================================*
Public Class SQLcreater
    ReadOnly log As log4net.ILog = _
log4net.LogManager.GetLogger( _
System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    ' SQL基本構文フィールド?には数字（品番ID）が入る
    Dim readPreStr As String = "SELECT ID FROM バースデーカード WHERE ID = ?"
    Dim FidnPreStr As String = "SELECT * FROM バースデーカード WHERE 氏名 = ?"
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
        Dim sLen, scLen1, scLen2 As Integer
        strSQL = ""
        strBuf1 = ""
        strBuf2 = InsertPreStr
        '入力項目回？の置き換えを行う
        For i = 0 To card.Length - 1
            sLen = InStr(strBuf2, "?")
            '？がある場合、手前から置き換えし、なくなったら入力項目数オーバーでエラー
            If (sLen > 0) Then
                strBuf1 = strBuf2.Substring(0, sLen)
                strBuf2 = strBuf2.Substring(sLen)
                strBuf1 = Replace(strBuf1, "?", "'" & card(i) & "'")
                strSQL &= strBuf1
            Else
                If card(i) <> "In" Then
                    MsgBox("入力項目が多すぎます")
                    strSQL = "-1"
                    Exit For
                End If

            End If
        Next
        strSQL &= strBuf2
        '再度？を確認し、あれば、置き換えする入力項目不足でエラー
        sLen = 0
        scLen1 = 0
        scLen2 = 0
        If InStr(strSQL, "?") > 0 Then
            sLen = InStr(strSQL, "?")
            strBuf1 = strSQL.Substring(0, sLen)
            strBuf2 = strSQL.Substring(sLen)
            scLen1 = InStr(strBuf1, "'")
            scLen2 = InStr(strBuf2, "'")
            If (scLen1 <= 0) And (scLen2 <= 0) Then
                MsgBox("入力項目が足りません")
                strSQL = "-1"
            End If
        End If
        Return strSQL
    End Function

    '***************************************************************
    ' 機能：ＩＤより検索用ＳＱＬ作成
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
    ' 機能：行進用ＳＱＬ作成
    ' 引数：レコード全項目
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Public Function UpdateBCard(ByRef card() As String)
        Dim strSQL, strBuf1, strBuf2 As String
        Dim sLen, scLen1, scLen2 As Integer
        strSQL = ""
        strBuf1 = ""
        strBuf2 = UpdatePreStr
        '入力項目回？の置き換えを行う
        For i = 0 To card.Length - 1
            sLen = InStr(strBuf2, "?")
            '？がある場合、手前から置き換えし、なくなったら入力項目数オーバーでエラー
            If (sLen > 0) Then
                strBuf1 = strBuf2.Substring(0, sLen)
                strBuf2 = strBuf2.Substring(sLen)
                strBuf1 = Replace(strBuf1, "?", "'" & card(i) & "'")
                strSQL &= strBuf1
            Else
                MsgBox("入力項目が多すぎます")
                strSQL = "-1"
                Exit For
            End If
        Next
        strSQL &= strBuf2
        '再度？を確認し、あれば、置き換えする入力項目不足でエラー
        sLen = 0
        scLen1 = 0
        scLen2 = 0
        If InStr(strSQL, "?") > 0 Then
            sLen = InStr(strSQL, "?")
            strBuf1 = strSQL.Substring(0, sLen)
            strBuf2 = strSQL.Substring(sLen)
            scLen1 = InStr(strBuf1, "'")
            scLen2 = InStr(strBuf2, "'")
            If (scLen1 <= 0) And (scLen2 <= 0) Then
                MsgBox("入力項目が足りません")
                strSQL = "-1"
            End If
        End If
        Return strSQL
    End Function

    '***************************************************************
    ' 機能：ＩＤより削除ＳＱＬ作成
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
    '***************************************************************
    ' 機能：氏名検索ＳＱＬ作成
    ' 引数：レコード全項目
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Public Function SearchName(ByRef strName As String)
        Dim strSQL As String
        strSQL = FidnPreStr
        If strName.Equals("") Then
            log.Error("氏名=" & strName & "が入力されていません。")
        Else
            strName = "'" & strName & "'"
        End If
        strSQL = Replace(strSQL, "?", strName)
        Return strSQL
    End Function

End Class

