Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '"C:\My Documents"以下のファイルをすべて取得
        'ワイルドカード"*"は、すべてのファイルを意味する
        Dim files As String() = System.IO.Directory.GetFiles( _
            "D:\佐々木の資料\科警研", "酒*.bmp", System.IO.SearchOption.AllDirectories)

        'ListBox1に結果を表示する
        ListBox1.Items.AddRange(files)




        '    'Visual Basic で特定のパターンに一致するファイルを検索する
        '    For Each foundFile As String In My.Computer.FileSystem.GetFiles( _
        'My.Computer.FileSystem.SpecialDirectories.MyDocuments, _
        'FileIO.SearchOption.SearchAllSubDirectories, "*.xlsx")

        'ListBox1.Items.Add(foundFile)
        'Next



    End Sub
End Class
