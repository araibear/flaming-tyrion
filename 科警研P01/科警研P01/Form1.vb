Option Explicit On
Option Compare Binary

Imports System.Management


Public Class Form1
    'ファイルキー文字列（連結）
    Dim keyFileName As String
    'ファイルパス情報
    Dim Inpath As String
    Dim Outpath As String
    Dim Imagepath As String
    Dim inFile As String
    Dim OutFile As String
    Dim ImageFile As String
    'キー項目
    Dim monitID As String
    Dim KANA As String
    Dim paperCnt As Integer = 0
    Dim TotalCnt As Integer = 0
    '2次元配列
    Dim csvDat(,) As String = Nothing
    Dim kanaDat(,) As String = Nothing
    Dim bcsvDat(,) As String = Nothing

    Dim mes As String
    Private Const csvPtFile As String = "D:\佐々木の資料\科警研\座標ファイル\記載枠座標_本番用変換済み.csv"            'CSV座標ファイルパス
    Private Const tmpFilePath As String = "D:\佐々木の資料\科警研\記入文字ファイル\"            '記入文字ファイルフォルダパス
    Private Const bmpPtfile As String = "D:\佐々木の資料\科警研\BMP座標ファイル\BMP座標_本番用変換済み.csv"            'BMP座標ファイルパス
    Private Const InOutfile As String = "D:\佐々木の資料\科警研\入出力ファイル\ファイル.txt"            '入出力ファイルパス
    Private Const PROGRAM_ID As String = "画像切出しツール"            'CSV座標ファイルパス
    Dim WriterLog As System.IO.StreamWriter
    Dim WriterInOut As System.IO.StreamWriter
    Dim inCSV As System.IO.StreamReader
    Dim ptXY As Point
    Dim shtMap As String()
    Dim rctSize As Integer = 360
    'コントロールの値を変更するためのデリゲート
    Private Delegate Sub SetProgressValueDelegate(ByVal num As Integer)
    'バックグラウンド処理が終わった時にコントロールの値を変更するためのデリゲート
    Private Delegate Sub ThreadCompletedDelegate()
    '処理がキャンセルされた時にコントロールの値を変更するためのデリゲート
    Private Delegate Sub ThreadCanceledDelegate()
    'キャンセルボタンがクリックされたかを示すフラッグ
    Private _canceled As Boolean = False
    '別処理をするためのスレッド
    Private workerThread As System.Threading.Thread

    Private ReadOnly canceledSyncObject As Object = New Object
    '*******************************************************************************************
    'キャンセルフラグプロパティ
    Private Property canceled() As Boolean
        Get
            SyncLock (canceledSyncObject)
                Return Me._canceled
            End SyncLock
        End Get
        Set(ByVal Value As Boolean)
            SyncLock (canceledSyncObject)
                Me._canceled = Value
            End SyncLock
        End Set
    End Property
    '*******************************************************************************************
    'イニシャル処理（画面起動時）
    '内容：設定ファイルを読み込む。
    '引数、戻り値・なし
    '*******************************************************************************************

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '設定ファイル読み込み
        tmpFileRead()
        '用紙対応配列の設定（コンスト）
        ReDim shtMap(45)
        shtMap(0) = "21"
        shtMap(1) = "31"
        shtMap(2) = "41"
        shtMap(3) = "51"
        shtMap(4) = "61"
        shtMap(5) = "71"
        shtMap(10) = "52"
        shtMap(11) = "62"
        shtMap(12) = "32"
        shtMap(13) = "42"
        shtMap(14) = "72"
        shtMap(15) = "22"
        shtMap(20) = "73"
        shtMap(21) = "53"
        shtMap(22) = "23"
        shtMap(23) = "43"
        shtMap(24) = "63"
        shtMap(25) = "33"
        shtMap(30) = "64"
        shtMap(31) = "34"
        shtMap(32) = "74"
        shtMap(33) = "54"
        shtMap(34) = "24"
        shtMap(35) = "44"
        shtMap(40) = "45"
        shtMap(41) = "75"
        shtMap(42) = "65"
        shtMap(43) = "25"
        shtMap(44) = "35"
        shtMap(45) = "55"

    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        inCSV = Nothing
        Dim Result As Boolean
        '//コントロールを初期化する
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 78
        ProgressBar1.Value = 0
        WriterLog = New System.IO.StreamWriter("D:\佐々木の資料\科警研\log\log.log", True, System.Text.Encoding.UTF8)

        mes = Date.Now.ToString & "処理開始します。"
        WriterLog.WriteLine(mes)

        Dim Source As System.Drawing.Bitmap
        '１．入力チェック   =======================================================***
        initCheck()

        '処理実行   ===========================================================***
        'ボタン類を日活性にする
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False

        '記入文字ファイルを読み込む
        paperCnt = 1
        TotalCnt = 1
        'キー項目セット
        monitID = TextBox4.Text
        '主処理　====================================================================***
        putImage()
        'For TotalCnt = 1 To 5
        '    '記入文字ファイルの読み込み
        '    kanaFileRead(TotalCnt)

        '    For paperCnt = 1 To 6
        '        '１次元まどぇ
        '        Dim strsht As String = shtMap(Integer.Parse(TotalCnt - 1.ToString & paperCnt - 1.ToString))
        '        'Imagepath = TextBox3.Text & "\" & _
        '        '            monitID & (paperCnt + 1).ToString & TotalCnt.ToString & ".bmp"
        '        ImageFile = Imagepath & "\" & _
        '                    monitID & strsht & ".bmp"
        '        Try
        '            'Source = Nothing
        '            GC.Collect()       ' System.GCクラス
        '            Source = New System.Drawing.Bitmap(ImageFile)
        '            '縦軸
        '            For i As Integer = 0 To 12
        '                '横軸
        '                For j As Integer = 0 To 8
        '                    KANA = kanaDat((paperCnt - 1) * 13 + i, j)
        '                    If Char.IsLower(KANA) Then
        '                        KANA &= KANA
        '                    End If
        '                    '画像切出し===========================================***
        '                    '3次元まで
        '                    keyFileName = KANA & "_" & monitID & "_" & TotalCnt
        '                    OutFile = Outpath & "\" & _
        '                                keyFileName & ".bmp"

        '                    '座標ファイル確認
        '                    inFile = Inpath & "\" & keyFileName & ".csv"
        '                    Try
        '                        inCSV = New System.IO.StreamReader(inFile, System.Text.Encoding.Default)
        '                        Dim temp() As String
        '                        Dim str As String
        '                        '一行目読み飛ばし
        '                        inCSV.ReadLine()
        '                        temp = Split(inCSV.ReadLine(), ",")
        '                        '座標チェック
        '                        ptXY.X = i
        '                        ptXY.Y = j
        '                        Result = ptCheck(CType(Double.Parse(temp(4)), Integer), CType(Double.Parse(temp(5)), Integer), i + 1, j + 1)
        '                        If Not Result Then
        '                            WriterLog.WriteLine(mes)
        '                        End If
        '                        '書き出し
        '                        ProgressBar1.Value = ((paperCnt - 1) * 13 + i) / 5 * TotalCnt
        '                        TrimImageBMP(Source, ptXY.X, ptXY.Y)
        '                        '画像切出し=終了==========================================***

        '                    Catch ex As Exception
        '                        mes = inFile & ex.ToString
        '                        WriterLog.WriteLine(mes)
        '                        'MsgBox(keyFileName & ".csvファイルが作成されていません。")
        '                    End Try

        '                Next
        '            Next

        '        Catch ex As Exception
        '            mes = Imagepath & ex.ToString
        '            WriterLog.WriteLine(mes)
        '            'MsgBox(mes)

        '        End Try


        '    Next
        'Next


 
        '終了処理 ====================================================================***
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True

        mes = Date.Now.ToString & "処理終了します。"
        WriterLog.WriteLine(mes)
        WriterLog.Close()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim fbd As New FolderBrowserDialog

        '上部に表示する説明テキストを指定する
        fbd.Description = "CSVフォルダの指定"
        fbd.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        fbd.SelectedPath = "D:\佐々木の資料\科警研\切り出しファイル結果ＣＳＶ"
        'ユーザーが新しいフォルダを作成できるようにする
        fbd.ShowNewFolderButton = True

        If fbd.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Inpath = fbd.SelectedPath
            TextBox1.Text = Inpath
        End If



    End Sub
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim fbd2 As New FolderBrowserDialog

        '上部に表示する説明テキストを指定する
        fbd2.Description = "画像保存フォルダの指定"
        fbd2.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        fbd2.SelectedPath = "D:\佐々木の資料\科警研\OUTPUT"
        'ユーザーが新しいフォルダを作成できるようにする
        fbd2.ShowNewFolderButton = True

        If fbd2.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Outpath = fbd2.SelectedPath
            TextBox2.Text = Outpath
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim fbd3 As New FolderBrowserDialog
        '上部に表示する説明テキストを指定する
        fbd3.Description = "画像保存フォルダの指定"
        fbd3.RootFolder = Environment.SpecialFolder.Desktop
        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        fbd3.SelectedPath = "D:\佐々木の資料\科警研\98-2-6"
        'ユーザーが新しいフォルダを作成できるようにする
        fbd3.ShowNewFolderButton = True

        If fbd3.ShowDialog(Me) = DialogResult.OK Then
            '選択されたフォルダを表示する
            Imagepath = fbd3.SelectedPath
            TextBox3.Text = Imagepath
        End If


    End Sub
    '***************************************************************
    ' 機能：座標ファイルと文字座標ファイルを読み込む
    ' 引数：なし
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Public Sub tmpFileRead()
        ' CSV座標ファイル読み込み処理==================================================================*
        Using sr As New System.IO.StreamReader(csvPtFile, System.Text.Encoding.Default)
            Dim i, j As Integer
            ReDim Preserve csvDat(156, 6)
            Do Until sr.Peek() = -1
                Dim temp() As String
                temp = Split(sr.ReadLine(), ",")

                For j = 0 To temp.GetUpperBound(0)
                    csvDat(i, j) = temp(j)
                Next
                i += 1
            Loop
        End Using


        ' BMP座標ファイル読み込み処理==================================================================*
        Using sr As New System.IO.StreamReader(bmpPtfile, System.Text.Encoding.Default)
            Dim i, j As Integer
            ReDim Preserve bcsvDat(156, 6)
            Do Until sr.Peek() = -1
                Dim temp() As String
                temp = Split(sr.ReadLine(), ",")

                For j = 0 To temp.GetUpperBound(0)
                    bcsvDat(i, j) = temp(j)
                Next
                i += 1
            Loop
        End Using


    End Sub
    '***************************************************************
    '記入文字ファイルを読み込む
    ' 引数：回数（何回目のシートか）
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Public Sub kanaFileRead(ByRef paperCnt As Integer)
        ' 記入文字ファイル読み込み処理==================================================================*
        Using sr As New System.IO.StreamReader(tmpFilePath & paperCnt & "回目.txt", System.Text.Encoding.Default)
            Dim i, j As Integer
            ReDim Preserve kanaDat(78, 9)
            Do Until sr.Peek() = -1
                Dim temp() As String
                Dim str As String
                str = sr.ReadLine()
                If InStr(str, paperCnt.ToString & "回目") > 0 Then
                    str = sr.ReadLine()
                End If
                temp = Split(str, vbTab)
                For j = 0 To 8
                    kanaDat(i, j) = temp(j)
                Next
                i += 1
            Loop
        End Using

    End Sub
    '***************************************************************
    ' 機能：座標ファイルと文字座標ファイルを読み込む
    ' 引数：なし
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****
    Public Sub TrimImageBMP(ByRef Source As System.Drawing.Bitmap, ByRef i As Integer, ByRef j As Integer)
        Dim temp As System.Drawing.Bitmap
        Dim Destination As System.Drawing.Bitmap
        Dim SourceRectangleF As System.Drawing.RectangleF
        Dim DestinationRectangleF As System.Drawing.RectangleF
        ' 画像解像度を取得する
        Dim hRes As Single = Source.HorizontalResolution
        Dim vRes As Single = Source.VerticalResolution

        Dim hColor As System.Drawing.Color = Nothing
        Dim PTFLG As Boolean = False
        ' ある整数と浮動小数の絶対値を取得する
        Dim rgAbs As Integer = 0
        Dim gbAbs As Integer = 0
        Dim brAbs As Integer = 0
        Dim clrCnt As Integer = 0

        '元イメージの開始座標
        SourceRectangleF.X = 395 + j * 475
        SourceRectangleF.Y = 750 + i * 630
        If j > 8 Then
            SourceRectangleF.X += 135
        End If
        '元イメージの範囲（サイズ）
        SourceRectangleF.Width = rctSize
        SourceRectangleF.Height = rctSize

        Dim ClrPt As Integer = 0
        For Ct As Integer = SourceRectangleF.Y - 50 To SourceRectangleF.Y + 50
            For Cs As Integer = SourceRectangleF.X To SourceRectangleF.X + 359
                hColor = Source.GetPixel(Cs, Ct)
                rgAbs = System.Math.Abs(Convert.ToInt32(hColor.R) - Convert.ToInt32(hColor.G))
                gbAbs = System.Math.Abs(Convert.ToInt32(hColor.G) - Convert.ToInt32(hColor.B))
                brAbs = System.Math.Abs(Convert.ToInt32(hColor.B) - Convert.ToInt32(hColor.R))
                If (rgAbs > 55 Or gbAbs > 55 Or brAbs > 55) Then
                    ClrPt += 1
                End If
            Next
            If ClrPt > 200 Then
                SourceRectangleF.Y = Ct - 20
                Exit For
            Else
                ClrPt = 0
            End If
        Next
        '枠の位置が取れるか？
        For Cj As Integer = SourceRectangleF.Y To SourceRectangleF.Y + 50
            If PTFLG Then
                Exit For
            End If
            For Ci As Integer = SourceRectangleF.X To SourceRectangleF.X + 50
                hColor = Source.GetPixel(Ci, Cj)
                rgAbs = System.Math.Abs(Convert.ToInt32(hColor.R) - Convert.ToInt32(hColor.G))
                gbAbs = System.Math.Abs(Convert.ToInt32(hColor.G) - Convert.ToInt32(hColor.B))
                brAbs = System.Math.Abs(Convert.ToInt32(hColor.B) - Convert.ToInt32(hColor.R))

                If (rgAbs > 55 Or gbAbs > 55 Or brAbs > 55) Then
                    clrCnt += 1
                    If clrCnt = 4 Then
                        SourceRectangleF.X = Ci - 20
                        SourceRectangleF.Y = Cj - 20
                        PTFLG = True
                        Exit For
                    End If
                Else
                    clrCnt = 0
                    Exit For
                End If
            Next
        Next
        '逆側に走査しなおす
        If PTFLG = False Then
            For Cj As Integer = SourceRectangleF.Y To SourceRectangleF.Y - 50 Step -1
                If PTFLG Then
                    Exit For
                End If
                For Ci As Integer = SourceRectangleF.X To SourceRectangleF.X - 50 Step -1
                    hColor = Source.GetPixel(Ci, Cj)
                    rgAbs = System.Math.Abs(Convert.ToInt32(hColor.R) - Convert.ToInt32(hColor.G))
                    gbAbs = System.Math.Abs(Convert.ToInt32(hColor.G) - Convert.ToInt32(hColor.B))
                    brAbs = System.Math.Abs(Convert.ToInt32(hColor.B) - Convert.ToInt32(hColor.R))

                    If (rgAbs > 55 Or gbAbs > 55 Or brAbs > 55) Then
                        clrCnt += 1
                        If clrCnt = 4 Then
                            SourceRectangleF.X = Ci - 20
                            SourceRectangleF.Y = Cj - 20
                            PTFLG = True
                            Exit For
                        End If
                    Else
                        clrCnt = 0
                        Exit For
                    End If
                Next
            Next
        End If


        'ただ保存するならtemp.Save("Q4045416-2.bmp")だけで良いが,
        '折角だから切り取った部分を任意の位置に描いてみよう。
        temp = New System.Drawing.Bitmap(rctSize, rctSize, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
        temp.SetResolution(hRes, vRes)
        temp = Source.Clone(SourceRectangleF, System.Drawing.Imaging.PixelFormat.Format24bppRgb)

        DestinationRectangleF.X = 0
        DestinationRectangleF.Y = 0
        DestinationRectangleF.Width = rctSize
        DestinationRectangleF.Height = rctSize

        Destination = New System.Drawing.Bitmap(rctSize, rctSize, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
        Destination.SetResolution(hRes, vRes)

        System.Drawing.Graphics.FromImage(Destination).DrawImage(temp, DestinationRectangleF)


        Destination.Save(OutFile)

    End Sub
    Public Function ptCheck(ByRef X As Integer, ByRef Y As Integer, ByRef Row As Integer, ByRef Col As Integer) As Boolean
        Dim trnRow As Integer = 0
        For i As Integer = 0 To 155
            If (Row = csvDat(i, 0) And Col = csvDat(i, 1)) Then
                '座標確認
                If (csvDat(i, 2) <= X And X <= csvDat(i, 4)) And csvDat(i, 3) <= Y And Y <= csvDat(i, 5) Then
                    Return True
                Else
                    For k As Integer = 0 To 155
                        If (Row = csvDat(k, 0) And csvDat(k, 1) = 10) Then
                            trnRow = k
                        End If
                    Next
                    If (csvDat(trnRow, 2) <= X And X <= csvDat(trnRow, 4)) And csvDat(trnRow, 3) <= Y And Y <= csvDat(trnRow, 5) Then
                        ptXY.X = Row - 1
                        ptXY.Y = 9
                        Return True
                    Else
                        For k As Integer = 0 To 155
                            If (Row = csvDat(k, 0) And csvDat(k, 1) = 11) Then
                                trnRow = k
                            End If
                        Next
                        If (csvDat(trnRow, 2) <= X And X <= csvDat(trnRow, 4)) And csvDat(trnRow, 3) <= Y And Y <= csvDat(trnRow, 5) Then
                            ptXY.X = Row - 1
                            ptXY.Y = 10
                            Return True
                        Else
                            For k As Integer = 0 To 155
                                If (Row = csvDat(k, 0) And csvDat(k, 1) = 12) Then
                                    trnRow = k
                                End If
                            Next
                            If (csvDat(trnRow, 2) <= X And X <= csvDat(trnRow, 4)) And csvDat(trnRow, 3) <= Y And Y <= csvDat(trnRow, 5) Then
                                ptXY.X = Row - 1
                                ptXY.Y = 11
                                Return True
                            Else
                                mes = inFile & "の座標X＝" & X & "：Y=" & Y & "は正しい座標ではありません。"
                                Return False
                            End If

                        End If

                    End If


                End If
            End If

        Next
    End Function

    '***************************************************************
    ' 機能：画面情報を入出力ファイルに追記する
    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        Dim strPar As String
        Dim Answer As Long
        ' 更新日時を取得する
        Dim dtUpdate As DateTime = System.IO.File.GetLastWriteTime(InOutfile)
        If dtUpdate.Day < Date.Now.Day Then
            Answer = MsgBox("入出力設定ファイルをクリアしますか？", vbYesNoCancel, "確認")
            Select Case Answer

                Case vbYes

                    WriterInOut = New System.IO.StreamWriter(InOutfile, False, System.Text.Encoding.Default)

                Case vbNo

                    WriterInOut = New System.IO.StreamWriter(InOutfile, True, System.Text.Encoding.Default)

                Case vbCancel

                    Exit Sub

            End Select
        Else
            WriterInOut = New System.IO.StreamWriter(InOutfile, True, System.Text.Encoding.Default)
        End If

 

        '１．入力チェック   =======================================================***
        initCheck()

        strPar = TextBox4.Text & "," & Inpath & "," & Outpath & "," & Imagepath
        WriterInOut.WriteLine(strPar)
        WriterInOut.Close()
    End Sub
    '***************************************************************
    ' 機能：入出力ファイルから処理を実行する
    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        'ボタン類を日活性にする
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button7.Enabled = False
        WriterLog = New System.IO.StreamWriter("D:\佐々木の資料\科警研\log\log.log", True, System.Text.Encoding.UTF8)

        Dim Reader As New IO.StreamReader(InOutfile, System.Text.Encoding.GetEncoding("Shift-JIS"))
        Dim Array() As String                'CSVの各項目を表す配列
        Dim Line As String = Reader.ReadLine 'CSVの一行
        Dim PostalCode As String             '郵便番号
        Dim Address As String                '住所
        Do Until IsNothing(Line)
            Array = Line.Split(",")
            monitID = Array(0)                     '被験者番号
            Inpath = Array(1)     '入力CSVく
            Outpath = Array(2)  '出力フォルダ
            Imagepath = Array(3)       '画像フォルダ
            '実処理
            putImage()

            Line = Reader.ReadLine                    '次の行を読み込む。

        Loop

        Reader.Close()

        '終了処理 ====================================================================***
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button7.Enabled = True

        mes = Date.Now.ToString & "処理終了します。"
        WriterLog.WriteLine(mes)
        WriterLog.Close()


    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub
    '*******************************************************************************************
    '入力チェック
    '内容：コントロール、または入出力設定ファイルの確認を行う
    '引数、戻り値・なし
    '*******************************************************************************************
    Private Sub initCheck()
        '１．入力チェック   =======================================================***
        '被験者番号チェック
        If TextBox4.Text.Equals("") Then
            MsgBox("被験者Noを入力してください。")
        Else
            Try
                Dim iA As Integer = Integer.Parse(TextBox4.Text, 0)
                If iA < 10 Then
                    TextBox4.Text = "00" & iA.ToString()
                ElseIf iA < 100 Then
                    TextBox4.Text = "0" & iA.ToString()
                End If
            Catch ex As Exception
                MsgBox("被験者Noは数字で入力してください。")
                Exit Sub
            End Try
        End If
        '入力ファイルチェック
        If TextBox1.Text.Equals("") Then
            MsgBox("CSVファイルのフォルダを選択してください。")
            Exit Sub
        End If
        '出力ファイルチェック
        If TextBox2.Text.Equals("") Then
            MsgBox("画像を保存するフォルダを選択してください。")
            Exit Sub
        End If

        '文字ファイルチェック
        If TextBox3.Text.Equals("") Then
            MsgBox("画像（原本）のファイルを選択してください。")
            Exit Sub
        End If

    End Sub
    '*******************************************************************************************
    '主処理
    '内容：該当のイメージの切り抜きを行う（画面、テキスト両方））
    '引数、戻り値・なし
    '*******************************************************************************************
    Private Sub putImage()
        inCSV = Nothing
        Dim Result As Boolean
        '//コントロールを初期化する
        ProgressBar1.Minimum = 0
        ProgressBar1.Maximum = 78
        ProgressBar1.Value = 0

        Dim Source As System.Drawing.Bitmap
        'ここから　====================================================================***
        For TotalCnt = 1 To 5
            '記入文字ファイルの読み込み
            kanaFileRead(TotalCnt)

            For paperCnt = 1 To 6
                '１次元まどぇ
                Dim strsht As String = shtMap(Integer.Parse(TotalCnt - 1.ToString & paperCnt - 1.ToString))
                'Imagepath = TextBox3.Text & "\" & _
                '            monitID & (paperCnt + 1).ToString & TotalCnt.ToString & ".bmp"
                ImageFile = Imagepath & "\" & _
                            monitID & strsht & ".bmp"
                Try
                    'Source = Nothing
                    GC.Collect()       ' System.GCクラス
                    Source = New System.Drawing.Bitmap(ImageFile)
                    '縦軸
                    For i As Integer = 0 To 12
                        '横軸
                        For j As Integer = 0 To 8
                            KANA = kanaDat((paperCnt - 1) * 13 + i, j)
                            If Char.IsLower(KANA) Then
                                KANA &= KANA
                            End If
                            '画像切出し===========================================***
                            '3次元まで
                            keyFileName = KANA & "_" & monitID & "_" & TotalCnt
                            OutFile = Outpath & "\" & _
                                        keyFileName & ".bmp"

                            '座標ファイル確認
                            inFile = Inpath & "\" & keyFileName & ".csv"
                            Try
                                inCSV = New System.IO.StreamReader(inFile, System.Text.Encoding.Default)
                                Dim temp() As String
                                Dim str As String
                                '一行目読み飛ばし
                                inCSV.ReadLine()
                                temp = Split(inCSV.ReadLine(), ",")
                                '座標チェック
                                ptXY.X = i
                                ptXY.Y = j
                                Result = ptCheck(CType(Double.Parse(temp(4)), Integer), CType(Double.Parse(temp(5)), Integer), i + 1, j + 1)
                                If Not Result Then
                                    WriterLog.WriteLine(mes)
                                End If
                                '書き出し
                                ProgressBar1.Value = ((paperCnt - 1) * 13 + i) / 5 * TotalCnt
                                TrimImageBMP(Source, ptXY.X, ptXY.Y)
                                '画像切出し=終了==========================================***

                            Catch ex As Exception
                                mes = inFile & ex.ToString
                                WriterLog.WriteLine(mes)
                                'MsgBox(keyFileName & ".csvファイルが作成されていません。")
                            End Try

                        Next
                    Next

                Catch ex As Exception
                    mes = Imagepath & ex.ToString
                    WriterLog.WriteLine(mes)
                    'MsgBox(mes)

                End Try


            Next
        Next

    End Sub
    '進行状況を表示する処理
    Private Sub CountUp()
        'デリゲートの作成
        Dim progressDlg As New SetProgressValueDelegate(AddressOf SetProgressValue)
        Dim completeDlg As New ThreadCompletedDelegate(AddressOf ThreadCompleted)
        Dim canceledDlg As New ThreadCanceledDelegate(AddressOf ThreadCanceled)

        '時間のかかる処理を開始する
        Dim i As Integer
        For i = 1 To 10
            'キャンセルボタンがクリックされたか調べる
            If canceled Then
                'キャンセルされたときにコントロールの値を変更する
                Me.Invoke(canceledDlg)
                '処理を終了させる
                Return
            End If

            '1秒間待機する（時間のかかる処理があるものとする）
            System.Threading.Thread.Sleep(1000)

            'コントロールの表示を変更する
            Me.Invoke(progressDlg, New Object() {i})
        Next

        '完了したときにコントロールの値を変更する
        Me.Invoke(completeDlg)
    End Sub

    'コントロールの値を変更する
    Private Sub SetProgressValue(ByVal num As Integer)
        'ProgressBar1の値を変更する
        ProgressBar1.Value = num
        'Label1のテキストを変更する
        Label1.Text = num.ToString()
    End Sub

    '処理が完了した時にコントロールの値を変更する
    Private Sub ThreadCompleted()
        Label1.Text = "完了しました。"
        Button1.Enabled = True
        Button2.Enabled = False
    End Sub

    '処理がキャンセルされた時にコントロールの値を変更する
    Private Sub ThreadCanceled()
        Label1.Text = "キャンセルされました。"
        Button1.Enabled = True
        Button2.Enabled = False
    End Sub

End Class
