Module Module1

    Sub Main()
        ' 名前空間UserCSVFileを指定して実行。
        UserCSVFile.ReadCSVFile.file()
        UserCSVFile.WriteCSVFile.file()
    End Sub

End Module
#Region "ＣＳＶファイル入出力ライブラリ"
Namespace UserCSVFile
    '***************************************************************
    ' クラス：ＣＳＶの読み込みに関するクラス
    ' メソッド：file()
    ' メソッド：split_s(String,String)
    '******1*********2*********3*********4*********5**********6*****
    Public Class ReadCSVFile
        '***************************************************************
        ' 機能：都道府県ＣＳＶを読み込み、コンソールに表示する
        ' 引数：なし
        ' 戻値：なし
        '******1*********2*********3*********4*********5**********6*****
        Public Shared Sub file()
            ' ファイル読み込み処理==================================================================*
            Dim Reader As New IO.StreamReader("D:\佐々木の資料\flaming-tyrion\名前空間\19yamana.csv",
                                              System.Text.Encoding.GetEncoding("Shift-JIS"))
            Dim Array() As String                'CSVの各項目を表す配列
            Dim Line As String = Reader.ReadLine 'CSVの一行
            Dim PostalCode As String             '郵便番号
            Dim Address As String                '住所
            Do Until IsNothing(Line)

                Array = Line.Split(",")                   '一行を, (カンマ)で区切って項目ごとに分解

                PostalCode = Array(2)                     '郵便番号取得
                PostalCode = PostalCode.Replace("""", "") '郵便番号から" を省く

                Address = Array(6) & Array(7) & Array(8)  '都道府県名と市区町村名と町域名を結合
                Address = Address.Replace("""", "")       '上記から" を省く

                Console.WriteLine(PostalCode & " - " & Address)

                Line = Reader.ReadLine                    '次の行を読み込む。

            Loop

            Reader.Close()

        End Sub
        '***************************************************************
        ' 機能：CSVの安全なスプリットメソッド（作りましょう）
        ' 引数：レコード（１行）
        ' 戻値：配列（済み）
        '******1*********2*********3*********4*********5**********6*****
        Private Function split_s(ByRef col As String(), ByRef rec As String)
            col(0) = 0
            Return col
        End Function


    End Class
    '***************************************************************
    ' クラス：ＣＳＶの書き込みに関するクラス
    ' メソッド：file()
    ' メソッド：outXML(String,String)
    '******1*********2*********3*********4*********5**********6*****
    Public Class WriteCSVFile
        '***************************************************************
        ' 機能：入力された回数、カラムを作り、ＣＳＶファイルとして出力
        ' 引数：なし
        ' 戻値：なし
        '******1*********2*********3*********4*********5**********6*****
        Public Shared Sub file()
            ' ファイル書き込み処理==================================================================*
            Dim Writer As New IO.StreamWriter("D:\佐々木の資料\flaming-tyrion\名前空間\VBTest.txt", True, System.Text.Encoding.UTF8)


            Dim i As Integer = 0
            Dim InStr, OutStr As String
            Console.WriteLine("ひとつめのカラムを入力してください")
            OutStr = Console.ReadLine()
            i += 1
            Do
                i += 1
                ' キーボードからカラムの中身を入力
                Console.WriteLine(i & "つめのカラム入力。入力を終えるときは、endを入力してください")
                InStr = Console.ReadLine()

                If InStr = "end" Then
                    ' endを入力したら、Exit Doで、Loopを抜ける
                    Exit Do
                End If
                OutStr = OutStr & "," & InStr

            Loop
            Writer.WriteLine(OutStr)
            Writer.Close()

        End Sub
        '***************************************************************
        ' 機能：配列をＸＭＬファイルにして出力する（作りましょう）
        ' 引数：配列
        ' 戻値：ＸＭＬファイル
        '******1*********2*********3*********4*********5**********6*****
        Private Function outXML(ByRef col As String()) As Integer
            col(0) = 0
            Return 0
        End Function

    End Class

End Namespace
#End Region



