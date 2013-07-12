Module Module1

    Sub Main()
        'for文による繰り返し　３回まで入力された数字を合計する。=====***
        ' 変数の宣言
        Dim inStr As String
        Dim inNum As Integer
        Dim sum As Integer = 0
        Console.WriteLine("*** 合計値を求める ***")
        Console.WriteLine()        ' 改行
        ' キーボードから3回数値を入力して、その合計を求める
        For i As Integer = 1 To 3
            Console.WriteLine(i & "番目の数字を入力して")
            Console.WriteLine("Enterキーを押してください")
            ' キーボードから入力
            inStr = Console.ReadLine()
            ' 文字列を整数に変換
            inNum = Integer.Parse(inStr)
            ' 入力値を前回までの合計値に加える
            sum = sum + inNum
        Next
        ' 合計を表示
        Console.WriteLine()        ' 改行
        Console.WriteLine("入力した数の合計は")
        Console.WriteLine(sum)
        Console.WriteLine("です")

        'for文のネスト（入れ子）による九九表の作成　====================***
        ' 九九表のタイトル表示
        Console.WriteLine("*** 九九表 ***")
        Console.WriteLine()    ' 改行
        Console.Write("            ")    ' 調整用空白
        For i As Integer = 1 To 9
            Console.Write(" X" & i)
        Next
        Console.WriteLine()    ' 改行

        ' 九九を二重ループを使って計算
        For j As Integer = 1 To 9
            Console.Write(j & "のだん --> ")
            For k As Integer = 1 To 9
                ' 九九計算と表示
                Console.Write("{0,3:#0}", j * k)
            Next
            ' 各段終了時に改行
            Console.WriteLine()
        Next


        '配列にFor each文を使うと便利　======================================***
        ' 配列変数の宣言
        Dim arryTensuu() As Integer = {85, 80, 90, 75, 60}
        sum = 0
        ' For Each Next を使って配列の合計を求める
        For Each tensuu As Integer In arryTensuu
            Console.WriteLine(tensuu & "を加算")
            ' 前回までの合計値に加える
            sum = sum + tensuu
        Next
        ' 合計を表示
        Console.WriteLine()        ' 改行
        Console.Write("合計は、")
        Console.Write(sum)
        Console.Write(" です")
        Console.WriteLine()        ' 改行

        'Do While文の例　入力値の合計が100になるまで繰り返す ==========***
        ' 変数の宣言
        Dim inStr3 As String
        Dim inNum3 As Integer = 0
        sum = 0

        ' 入力値の和が100を超えるまで入力を続ける
        Do While sum <= 100
            Console.WriteLine("整数を入力してください")
            ' キーボードから入力
            inStr3 = Console.ReadLine()
            ' 入力した文字列を整数に変換
            inNum3 = Integer.Parse(inStr3)
            ' 入力値を加算
            sum = sum + inNum3
        Loop

        ' 合計値を表示
        Console.WriteLine()    ' 改行
        Console.WriteLine("お疲れ様でした")
        Console.WriteLine("入力した数値の和は")
        Console.WriteLine(sum)
        Console.WriteLine("です")

        'Do Until文の例　入力値の合計が100以下で繰り返す ==========***
        ' 変数の宣言
        Dim inStr4 As String
        Dim inNum4 As Integer = 0
        sum = 0
        ' 入力値の和が100を超えなければ入力を続ける
        Do Until sum > 100
            Console.WriteLine("整数を入力してください")
            ' キーボードから入力
            inStr4 = Console.ReadLine()
            ' 入力した文字列を整数に変換
            inNum4 = Integer.Parse(inStr4)
            ' 入力値を加算
            sum = sum + inNum4
        Loop

        ' 合計値を表示
        Console.WriteLine()    ' 改行
        Console.WriteLine("お疲れ様でした")
        Console.WriteLine("入力した数値の和は")
        Console.WriteLine(sum)
        Console.WriteLine("です")

        'Loop Whileの文例　100を超えるまで繰り返す ==========================***
        ' 変数の宣言
        Dim inStr5 As String
        Dim inNum5 As Integer = 0
        sum = 0

        ' Do Loop While を使って
        ' 入力値の和が100を超えるまで入力を続ける
        Do
            Console.WriteLine("整数を入力してください")
            ' キーボードから入力
            inStr5 = Console.ReadLine()
            ' 入力した文字列を整数に変換
            inNum5 = Integer.Parse(inStr5)
            ' 入力値を加算
            sum = sum + inNum5
        Loop While sum <= 100

        ' 合計値を表示
        Console.WriteLine()    ' 改行
        Console.WriteLine("=======================")
        Console.WriteLine("入力した数値の合計は")
        Console.WriteLine(sum)
        Console.WriteLine("です")

    End Sub

End Module
