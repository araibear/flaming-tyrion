Module Module1

    Sub Main()
        '自作クラスライブラリシャープの宣言とインスタンス
        Dim Cl1 As New ClassLibrary1.Class1
        'シェルの宣言とインスタンス
        Dim o As New Shell32.Shell
        Dim str As String

        str = Cl1.TestMeth("a")
        '自作ライブラリのTestMethの戻り値6を表示する
        System.Console.WriteLine(str)
        'すべてのウィンドウを最小化する
        o.ToggleDesktop()
        'エクスプローラで指定のフォルダを開く
        o.Open("C:\windows\")


    End Sub

End Module
