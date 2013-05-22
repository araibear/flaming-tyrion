Module Module1

    Sub Main()
        ' コンソールに文字を表示させる	Console.Write("")
        Console.WriteLine("最初のプログラム")


        ' 文字列変数の宣言	Dim (変数名) As String
        Dim str1 As String
        Dim str2, str3 As String
        ' 数値変数Integerの宣言	Dim （数値変数名） As Integer
        Dim kakaku As Integer




        ' 一度変数の宣言をしてから、改めて変数に値を代入する
        Dim str As String
        Dim kingaku As Integer

        ' 文字列変数に代入
        str = "文字列です"
        kingaku = 1250

        ' 文字列変数の宣言と初期化
        Dim strInit As String = "文字列です"
        Dim kingakuInit As Integer = 1250

        ' 文字列をコンソールに表示
        Console.WriteLine(strInit)
        Console.WriteLine(kingakuInit)


        ' 文字列変数の宣言
        Dim after1 As String
        ' 文字列変数の宣言と代入
        Dim str1gyo As String = "1行目"
        Dim common As String = "の文字列"
        ' 文字列を連結して画面に表示
        after1 = str1gyo & common
        Console.WriteLine(after1)


        ' 文字列変数の宣言
        Dim after As String
        ' 文字列変数を混在させて連結
        Dim commonStr As String = "の文字列"
        after = "1" & "行目" & common
        Console.WriteLine(after)

        Dim gyosu As Integer = 6
        ' このような数値型と文字列の連結はエラーとなる
        ' after = "1" & gyosu & common
        Console.WriteLine(after)

        ' 4つ文字列を連結して表示




        ' 文字列変数の宣言
        Dim tenki As String
        ' 文字列変数の宣言と代入
        Dim tenki1 As String = "今日の天気は"
        Dim tenki2 As String = "晴れのち曇り"
        Dim tenki3 As String = "所によっては"
        Dim tenki4 As String = "一時雨でしょう"

        ' 文字列変数を連結して4行で画面に表示
        tenki = tenki1 & vbCrLf & tenki2 & vbCrLf & tenki3 & vbCrLf & tenki4
        Console.WriteLine(tenki)
    End Sub

End Module
