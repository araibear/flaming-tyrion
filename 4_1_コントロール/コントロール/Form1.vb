Public Class Form1

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

    '***************************************************************
    ' 機能：初期処理
    ' 引数：キー項目氏名
    ' 戻値：ＳＱＬ文字列
    '******1*********2*********3*********4*********5**********6*****
    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'リストボックスにファイルを表示する
        ListBox1.Items.Clear()
        Dim files As String() = System.IO.Directory.GetFiles( _
           "C:\test", "*.txt", System.IO.SearchOption.AllDirectories)
        filesBack = files.Clone
        ListBox1.Items.AddRange(files)
        ListBox1.SelectedIndex = 0
        '初期設定
        Me.Text = FRMTITLE
        Button1.Text = BTNTITLE
        TextBox1.Focus()
        Button1.Text = BTNTITLE
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.ControlBox = True
        Me.WindowState = FormWindowState.Normal
        ComboBox1.Text = "はじめ"
    End Sub
    Private Sub Form1_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseClick, RichTextBox1.MouseClick
        Dim x, y As Integer
        Dim mType As String
        x = e.X
        y = e.Y
        mType = e.Button.ToString
        Label1.Text = "Ｘ座標：" & x & "Ｙ座標：" & y & "マウス：" & mType
    End Sub

    Private Sub Button1_MouseHover(sender As System.Object, e As System.EventArgs) Handles Button1.MouseHover
        Button1.Text = "データを入力してください。"
    End Sub

    Private Sub Button1_MouseLeave(sender As System.Object, e As System.EventArgs) Handles Button1.MouseLeave
        Button1.Text = BTNTITLE
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Dim str As String = Nothing
        '4-1②Senderを利用した例 ==============***
        Dim ss As String
        ss = sender.ToString
        RichTextBox1.Text = ss

        '4-2③Linesプロパティの使用例
        For i = 0 To TextBox1.Lines.Length - 1
            str &= TextBox1.Lines(i)
        Next
        RichTextBox1.Text &= str

        'グループボックスの活用例()
        'GroupBox1.Enabled = False
        'TextBox1.Text = "現在" & GroupBox1.Text & "は" & GroupBox1.Enabled & "です。"


        '入力項目が満たされているかチェックする
        '例文チェックボックスの例
        If (RadioButton1.Checked = False) And (RadioButton2.Checked = False) Then
            GroupBox1.Focus()
            MsgBox("性別を選択してください。")
            Exit Sub
        End If

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            CheckBox1.Text = "チェックをいれないでください"
        Else
            CheckBox1.Text = "チェックをいれてください"
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim i As Integer = ListBox1.SelectedIndex
        File_to_RichT(filesBack(i))
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

    Private Sub MonthCalendar1_DateChanged(sender As System.Object, e As System.Windows.Forms.DateRangeEventArgs) Handles MonthCalendar1.DateChanged
        Dim birthDay As Date
        Dim bstr As String
        birthDay = e.Start

        bstr = birthDay.ToString("yyyy年MM月dd日")
        birthDay = e.End
        bstr &= birthDay.ToString("yyyy年MM月dd日")
        TextBox1.Text = bstr

    End Sub
    '***************************************************************
    ' 機能：（自作クラス）文面雛形をRichTextに展開する
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


End Class


