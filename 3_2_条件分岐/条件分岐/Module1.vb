Module Module1

    Sub Main()
        'If文の例　入力された数を正、負、ゼロの３つに場合わけする =================***
        ' 変数の宣言
        Dim inStr As String
        Dim inNum As Integer
        Dim msg As String
        ' キーボードから入力
        Console.WriteLine("整数を入力してください")
        inStr = Console.ReadLine()
        ' 文字列を整数に変換
        inNum = Integer.Parse(inStr)
        ' 入力値を判定して処理を振分
        If inNum > 0 Then
            ' 正の整数用のメッセージを作成
            msg = "正の整数が入力されました"
        ElseIf inNum = 0 Then
            ' ゼロ用のメッセージを作成
            msg = "ゼロが入力されました"
        Else
            ' 負の整数用のメッセージを作成
            msg = "負の整数が入力されました"
        End If
        ' 結果を表示
        Console.WriteLine(msg)

        'CASE文の例　入力された年齢で、コメント（つまらないが・・・）を分ける =====***
        ' 変数の宣言
        Dim inStr2 As String
        Dim inNum2 As Integer
        ' キーボードから入力
        Console.WriteLine("年齢を入力してください")
        inStr2 = Console.ReadLine()
        ' 文字列を整数に変換
        inNum2 = Integer.Parse(inStr2)
        ' 年齢を判定して処理を振分
        Select Case inNum2
            Case Is < 0, Is >= 150
                ' メッセージを作成
                msg = "年齢に誤りがあります"
            Case 0 To 19
                msg = "まだ選挙権はありません"
            Case 20
                msg = "大人の仲間入りです"
            Case 30, 40, 50
                msg = inStr & "歳です。飲みすぎに注意"
            Case 21 To 59
                msg = "運動していますか"
            Case Else
                msg = "健康に気をつけましょう"
        End Select
        ' 結果を表示
        Console.WriteLine(msg)


    End Sub

End Module
