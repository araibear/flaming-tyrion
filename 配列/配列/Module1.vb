Module Module1
    '数字でない文字列を入力した場合に変換する数値：デフォルト：0
    Public ERRNUM As Integer = 1


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
        Dim price() As Integer = {89, 105, 525, 480, 1980}
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
        ' 成績の2次元配列変数の初期化{国語,数学, 英語}
        Dim score(,) As Integer = {{75, 84, 90}, _
                                    {100, 80, 80}, _
                                    {90, 89, 90}, _
                                    {65, 90, 80}, _
                                    {69, 85, 100}}
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


        '動的配列ReDimで再定義可能
        Dim aryDynamic() As Integer
        Dim num As Integer
        '入力した数の配列として再定義する
        num = ParseInt_s(Console.ReadLine(), ERRNUM)
        ReDim aryDynamic(num)
        '配列は0から始まるため、-1する
        For i As Integer = 0 To num - 1
            aryDynamic(i) = ParseInt_s(Console.ReadLine(), ERRNUM)
        Next
        '実は、配列数外の値を表示しようとしてもこけない
        For i As Integer = 0 To num
            Console.WriteLine(aryDynamic(i).ToString)
        Next
        Console.WriteLine("最後には0が表示されましたか？")

        '配列の代入とクローン

        Dim ary1() As String = {"あ", "い", "う", "え", "お"}
        Dim ary2() As String = {"か", "き", "く", "け", "こ"}
        'まずは配列を見てみます。
        For i As Integer = 0 To 4
            Console.WriteLine("ary1(" & i & "):" & ary1(i))
        Next
        For i As Integer = 0 To 4
            Console.WriteLine("ary2(" & i & "):" & ary2(i))
        Next
        '配列を代入
        ary1 = ary2
        '配列１の４つめに「せ」をセット
        ary1(3) = "せ"
        Console.WriteLine("ary1(3):" & ary1(3))
        Console.WriteLine("ary2(3):" & ary2(3))
        ary2 = {"か", "き", "く", "け", "こ"}
        '配列の型だけ継承したい場合はcloneを使いましょう
        ary1 = ary2.Clone
        ary1(3) = "せ"
        Console.WriteLine("ary1(3):" & ary1(3))
        Console.WriteLine("ary2(3):" & ary2(3))
    End Sub
    '***************************************************************
    ' 機能：入力文字列を安全に数字変換する関数
    ' 第１引数：変換する文字列
    ' 第２引数：エラーだった場合に置き換える数字
    ' 戻値：置き換わった結果
    '******1*********2*********3*********4*********5**********6*****
    Public Function ParseInt_s(ByRef strNum As String, ByRef errNum As Integer)
        Dim res As Integer
        If (Integer.TryParse(strNum, res)) Then
        Else
            res = errNum
            Console.WriteLine("********** ERROR :{0} 数字ではありません。 **********", strNum)
        End If
        Return res
    End Function
End Module
