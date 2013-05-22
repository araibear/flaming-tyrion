Module Module1

    Sub Main()

        ' 整数値変数Integer -2,147,483,648 ～ 2,147,483,647までの整数
        Dim ans As Integer = 0
        ' Long型 
        ' -9,223,372,036,854,775,808 ～ 9,223,372,036,854,775,807 までの整数
        Dim kingaku As Long
        ' Short 型 -32,768 ～ 32,767 までの整数
        Dim shtAns As Short
        ' Double型 (倍精度浮動小数点数型)
        Dim dblAns As Double

        ' 加算（足し算）
        ans = 5 + 2
        Console.WriteLine(ans)   ' ansは 7

        ' 整数値変数の宣言と代入
        Dim a As Integer = 3
        Dim b As Integer = 12
        Dim c As Integer = 12
        Dim d As Integer = 12

        ' 加算
        ans = a + b
        Console.WriteLine(ans)   ' ansは 15


        ' 減算
        ans = a - b
        Console.WriteLine(ans)   ' ansは -5
        ' このような書き方もできる
        a = a - 1


        ' 乗算（掛け算）
        ans = 5 * 2
        Console.WriteLine(ans)   ' ansは 10
        ' 整数値変数の宣言と代入
        a = 100000000
        b = 1000
        ans = 0
        ' 演算結果は、ans変数には入りきらない大きい整数値（オーバーフロー）
        'ans = a * b
        Console.WriteLine(ans)


        ' Long型変数に揃えて宣言
        Dim longC As Long = -10
        Dim longD As Long = 9
        Dim ansL As Long = 0
        ' 整数除算
        ansL = longC \ longD
        Console.WriteLine(ansL)    ' ans2は -1
        ' ans2変数を使わないで変数cに再度代入
        c = c \ d
        Console.WriteLine(c)       ' cは -1


        ' Integer型変数の宣言と代入
        a = -3
        b = -2
        ans = 0
        ' 剰余計算
        ans = a Mod b
        Console.WriteLine(ans)    ' ansは -1



        ' Long型変数の宣言と代入
        Dim zero As Long = 0
        Dim ansZeroDiv As Long = 0
        ' 整数除算
        ansZeroDiv = zero Mod 10
        Console.WriteLine(ans)    ' 0
        'ansZeroDiv = 10 Mod zero           System.DivideByZeroException
        Console.WriteLine(ans)



        ' Double型変数の宣言と代入
        Dim ans1 As Double = 0
        Dim ans2 As Double = 0
        Dim ans3 As Double = 0
        Dim ans4 As Double = 0
        ' べき乗計算
        ans1 = 2 ^ 2
        Console.WriteLine(ans1)   ' ans1は 4
        ans2 = -2 ^ 2
        Console.WriteLine(ans2)   ' ans2は -4
        ans3 = (-2) ^ 2
        Console.WriteLine(ans3)   ' ans3は 4
        ans4 = 2 ^ -2             ' 2 / (2 * 2)
        Console.WriteLine(ans4)   ' ans4は 0.5
        ' Double型変数の宣言と代入
        Dim ans5 As Double = 0
        Dim dA As Double = 2
        ' べき乗計算
        ans5 = 1.1 ^ dA
        Console.WriteLine(ans5)   ' ans5は 1.21


        Dim numL As Integer = 5
        Dim numR As Integer = 3

        numL += numR              ' numLは  8  numL = numL + numR と同じ
        numL += 10                ' numLは 18  numL = numL + 10 と同じ

        numL = 5
        numR = 3

        numL -= numR              ' numLは  2  numL = numL - numR と同じ
        numL -= 10                ' numLは -8  numL = numL - 10 と同じ

        numL = 5
        numR = 3

        numL *= numR              ' numLは  15  numL = numL * numR と同じ
        numL *= 10                ' numLは 150  numL = numL * 10 と同じ

        numL = 5
        numR = 3

        numL /= numR    ' numLは 1.6666666666666667  numL = numL / numR と同じ
        numL /= 10      ' numLは 0.16666666666666669  numL = numL / 10 と同じ



        numL = 2
        numR = 3

        numL ^= numR             ' numLは  8.0  numL = numL ^ numR と同じ
        numL ^= 2                ' numLは 64.0  numL = numL ^ 2 と同じ


        Dim str As String = "Visual "
        Dim str1 As String = "Basic "
        Dim str2 As String = "2010"

        str &= str1      ' strは Visual Basic       str = str & str1 と同じ
        str &= str2      ' strは Visual Basic 2010  str = str & str2 と同じ

    End Sub

End Module
