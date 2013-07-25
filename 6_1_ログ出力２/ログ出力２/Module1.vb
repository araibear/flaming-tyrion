Module Module1
    ' ライブラリlog4netを参照したら、下記のようにコンストラクタ生成する
    ' もし、log4netが参照できない場合、コンパイル→詳細コンパイルオプションボタン
    ' →対象のフレームワーク（すべての構成）で.net framework 4.0を選択する。
    ' Client Profileになっている場合がある。
    ' ソリューション→
    ReadOnly log As log4net.ILog = _
      log4net.LogManager.GetLogger( _
      System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    'ログファイルパス
    Public LOGFILE As String = "C:\log\logfile.txt"

    '自作ログ構造体
    Private Structure tLog
        Public sysTime As DateTime
        Public sysCall As Integer
        Public level As String
        Public appName As String
        Public mes As String
    End Structure


    Sub Main()
        Dim mes As String
        mes = "program started."
        Console.WriteLine(mes)
        ' これはデバッグレベルのログが出力される
        log.Debug(mes)
        ' これはフェイタルレベル（最も高い）のログが出力される
        log.Fatal("Program  ended!")

        ' コンソールとログ双方に同じメッセージを出力する
        mes = "program continued...."
        Dim myobj As New someclass
        myobj.mymethod(mes)

        'こっちは自作ログ出力
        LogCreater("info", mes)

    End Sub
    '***************************************************************
    ' 機能：ログ作成
    ' 第1引数：レベル
    ' 第２引数：出したいメッセージ
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    '自作ログ出力メソッド
    Private Sub LogCreater(ByRef level As String, ByRef mes As String)

        Dim Line As String = ""
        Dim log As tLog
        'ログに必要な情報を格納
        log.appName = My.Application.Info.Title
        log.sysTime = Date.Now
        log.sysCall = 1
        log.level = level
        log.mes = mes
        '出力ストリーム
        Dim Writer As New IO.StreamWriter(LOGFILE, True, System.Text.Encoding.UTF8)
        '1行分のログを作成
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


End Module
'***************************************************************
' 機能：Log4netによるログ出力
' 概要：Log4netを使って、ログを出してくれる
'******1*********2*********3*********4*********5**********6*****
Public Class someclass
    ' 個別のクラスに設定する場合も以下のようにコンストラクタ生成する
    Private Shared ReadOnly log As log4net.ILog = _
      log4net.LogManager.GetLogger( _
      System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    '***************************************************************
    ' 機能：ログ出力、コンソール出力
    ' 引数：メッセージ
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Public Sub mymethod(ByRef mes As String)
        '以下のようにメッセージを参照渡しして、コンソールとログ両方に出力するものを作るのもよい。
        log.InfoFormat(mes)
        Console.WriteLine(mes)
        log.InfoFormat("mymethod:今日の日付 Date:{0}", DateTime.Now)
    End Sub



End Class


