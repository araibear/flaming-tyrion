Module Module1
    '数字でない文字列を入力した場合に変換する数値：デフォルト：0
    Public ERRNUM As Integer = 1

    Sub Main()
        ' 関数を使わないバージョンです。
        ' 変数の宣言
        Dim inStr As String
        Dim zeikomi As Double
        ' キーボードから3商品の税込み価格を入力
        Console.WriteLine("商品の税込み価格を入力してEnterキーを押してください")
        inStr = Console.ReadLine()
        ' 文字列を整数に変換して配列に保存
        zeikomi = ParseInt_s(inStr, ERRNUM)

        Dim hontai As Integer
        Dim zei As Integer
        ' 消費税額を求めて表示する
        hontai = zeikomi / 1.05    ' 本体価格は四捨五入
        zei = zeikomi - hontai
        Console.WriteLine("消費税は" & zei & "円です")

        ' Function関数を使うバージョンです。=================================

        ' 変数の宣言
        inStr = ""
        zeikomi = 0
        hontai = 0
        Do
            ' キーボードから3商品の税込み価格を入力
            Console.WriteLine("商品の本体価格を入力してEnterキーを押してください")
            inStr = Console.ReadLine()

            If inStr = "end" Then
                ' endを入力したら、Exit Doで、Loopを抜ける
                Exit Do
            End If
            ' 文字列を整数に変換して配列に保存
            hontai = ParseInt_s(inStr, ERRNUM)
            Dim seireki As Integer
            seireki = Year(Now)
            ' 消費税率を求めて返します
            Dim zeiritu As Double
            zeiritu = CheckDutyRate(seireki)
            Console.WriteLine("消費税率は" & zeiritu & "％です")
            CalucZeikomi(hontai, zeiritu)
            Console.WriteLine("プログラムを終える場合はendと入力してenterを押してください")
        Loop
    End Sub

    '***************************************************************
    ' 機能：消費税額を計算する関数
    ' 引数：商品価格
    ' 戻値：消費税額
    '******1*********2*********3*********4*********5**********6*****
    Private Function CheckDutyRate(ByRef seireki As Integer) As Integer
        Dim zeiritu As Integer = 5    ' リターン値
        ' 消費税率を求める
        If seireki < 2014 Then
            zeiritu = 5
        Else
            zeiritu = 8
        End If
        ' ◆戻値を設定してリターン
        Return zeiritu
    End Function

    '***************************************************************
    ' 機能：税込価格を計算して表示する関数
    ' 引数：商品価格
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Private Sub CalucZeikomi(ByRef hontai As Integer, zeiritu As Double)
        Dim zeikomi As Integer = 0
        ' 税込価格を計算
        zeikomi = hontai * (1 + zeiritu / 100)
        ' 税込価格を表示
        Console.WriteLine()              ' 改行
        Console.WriteLine("税込価格は")
        Console.WriteLine(CDec(zeikomi) & "円です")
    End Sub
    '***************************************************************
    ' 機能：入力文字列を安全に数字変換する関数
    ' 第１引数：変換する文字列
    ' 第２引数：エラーだった場合に置き換える数字
    ' 戻値：置き換わった結果
    '******1*********2*********3*********4*********5**********6*****
    Private Function ParseInt_s(ByRef strNum As String, ByRef errNum As Integer)
        Dim res As Integer
        If (Integer.TryParse(strNum, res)) Then
        Else
            res = errNum
            Console.WriteLine("********** ERROR :{0} 数字ではありません。 **********", strNum)
        End If
        Return res
    End Function


End Module
