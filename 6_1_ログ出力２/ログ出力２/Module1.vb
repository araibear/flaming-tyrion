Module Module1
    ' ライブラリlog4netを参照したら、下記のようにコンストラクタ生成する
    ' もし、log4netが参照できない場合、コンパイル→詳細コンパイルオプションボタン
    ' →対象のフレームワーク（すべての構成）で.net framework 4.0を選択する。
    ' Client Profileになっている場合がある。
    ' ソリューション→
    ReadOnly log As log4net.ILog = _
      log4net.LogManager.GetLogger( _
      System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)


    Sub Main()
        Dim mes As String
        mes = "program started."
        Console.WriteLine(mes)
        ' これはデバッグレベルのログが出力される
        log.Debug(mes)
        ' コンソールとログ双方に同じメッセージを出力する
        mes = "program continued...."
        Dim myobj As New someclass
        myobj.mymethod(mes)
        ' これはフェイタルレベル（最も高い）のログが出力される
        log.Fatal("Program  ended!")
    End Sub

End Module

Public Class someclass
    ' 個別のクラスに設定する場合も以下のようにコンストラクタ生成する
    Private Shared ReadOnly log As log4net.ILog = _
      log4net.LogManager.GetLogger( _
      System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub mymethod(ByRef mes As String)
        '以下のようにメッセージを参照渡しして、コンソールとログ両方に出力するものを作るのもよい。
        log.InfoFormat(mes)
        Console.WriteLine(mes)
        log.InfoFormat("mymethod:今日の日付 Date:{0}", DateTime.Now)
    End Sub

End Class


