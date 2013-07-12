Public Class Form1
    Public strMessage() As String = {"１", "２", "３", "４", "５", "６", "７", "８"}
    'ボタンの名称
    Public BTNTITLE As String = "バースデーカードを作る"
    'アプリケーションタイトル
    Public FRMTITLE As String = "バースデーカードエディタ"
    'リストのあるディレクトリパス
    Public LISTPATHE As String = "C:\test"
    '出力ファイルディレクトリパス
    Public OFILEPATHE As String = "C:\letter"

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

    End Structure

    '***************************************************************
    ' 機能：初期処理
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Text = FRMTITLE
        Button1.Text = BTNTITLE
        ListBox1.Items.Clear()
        Dim files As String() = System.IO.Directory.GetFiles( _
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
        Button1.Text = BTNTITLE
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.ControlBox = True
        Me.WindowState = FormWindowState.Normal
        ComboBox1.Text = "はじめ"
    End Sub

    Private Sub Button1_MouseHover(sender As System.Object, e As System.EventArgs) Handles Button1.MouseHover
        If TextBox1.Text.Equals("") And RichTextBox1.Text.Equals("") Then
            Button1.Text = "データを入力してください。"
        End If
    End Sub

    Private Sub Button1_MouseLeave(sender As System.Object, e As System.EventArgs) Handles Button1.MouseLeave
        Button1.Text = BTNTITLE
    End Sub
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
        Rich_to_FileT(OFILEPATHE)

    End Sub
    '***************************************************************
    ' 機能：各コントロールの情報をカード構造体にセットする
    ' 引数：なし
    ' 戻値：カード構造体BithdayCard
    '******1*********2*********3*********4*********5**********6*****
    Private Function CtrlSet()
        Dim card As BithdayCard
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

        strMessage(0) = card.Name.ToString
        strMessage(1) = card.BithDay.ToString("yyyy-MM-dd")
        strMessage(2) = card.nTitle.ToString
        strMessage(3) = card.sex.ToString
        strMessage(4) = card.txtList.ToString
        strMessage(5) = card.mainText.ToString
        strMessage(6) = card.up_Date.ToString("yyyy-MM-dd")
        'strMessage(7) = card.id.ToString

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

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim i As Integer = ListBox1.SelectedIndex
        File_to_RichT(filesBack(i))
    End Sub
    '***************************************************************
    ' 機能：バースデーカードの出力
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Rich_to_FileT(ByRef fPath As String)
        Dim Line As String = ""
        Dim Writer As New IO.StreamWriter(fPath & "\VBTest.txt", True, System.Text.Encoding.UTF8)
        For i = 0 To RichTextBox1.Lines.Length - 1
            Line = RichTextBox1.Lines(i)
            Writer.WriteLine(Line)

        Next

        Writer.Close()

    End Sub

    Private Sub ComboBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.TextChanged
        Dim str As String = ComboBox1.Text
        If str.Length <> 0 Then
            For i = 0 To ComboBox1.Items.Count - 1
                If (InStr(ComboBox1.Items(i).ToString, str) > 0) Then
                    ComboBox1.SelectedIndex = i
                    Exit For
                Else
                    ComboBox1.Text = str
                End If
            Next
        End If




    End Sub
End Class

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

