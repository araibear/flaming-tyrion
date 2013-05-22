Module Module1

    Sub Main()
        '文字列配列の宣言　変数名() As String
        Dim str(20) As String

        For i As Integer = 0 To 20
            ' 文字列の連結
            str(i) = (i + 1) & "行目"

            ' 文字列の表示
            Console.WriteLine(str(i))
        Next



        ' １次元整数配列変数の初期化
        Dim price() As Integer = {120, 80, 55, 300, 168}
        ' 要素の値変更
        price(1) = price(1) * -1
        price(3) = price(3) * -1
        ' 配列要素の表示
        For i As Integer = 0 To 4
            Console.WriteLine(price(i))
        Next




        ' 生徒数
        Dim cntStudent As Integer = 5
        ' 科目数
        Dim cntSubject As Integer = 3
        ' 成績の2次元配列変数の初期化{国語, 算数, 理科}
        Dim score(,) As Integer = {{75, 85, 90}, _
                                    {100, 80, 80}, _
                                    {90, 90, 90}, _
                                    {65, 90, 80}, _
                                    {80, 85, 100}}
        ' 合計値の2次元配列変数の宣言
        Dim sum(0, 2) As Integer
        ' 平均値の2次元配列変数の宣言（データ型に注意）
        Dim avgScore(0, 2) As Double
        ' 各科目合計を求める
        For i As Integer = 0 To cntStudent - 1
            sum(0, 0) = sum(0, 0) + score(i, 0)
            sum(0, 1) = sum(0, 1) + score(i, 1)
            sum(0, 2) = sum(0, 2) + score(i, 2)
        Next
        ' 平均値の計算と表示
        For j As Integer = 0 To cntSubject - 1
            avgScore(0, j) = sum(0, j) / cntStudent
            Console.Write("avgScore(0, " & j & ") = ")
            Console.WriteLine(avgScore(0, j))
        Next


        '動的配列
    End Sub

End Module
