Imports System.Drawing.Drawing2D

Public Class Form1

    Public Const 倍率 As Integer = 2
    Public Const 余白 As Integer = 5
    Public 分割() As String
    Public x1, x2, y1, y2 As Decimal
    Public flg As Integer
    Public II As Integer
    Public 抽出 As String
    'キー項目
    Public monitID As String
    Public KANA As String
    Public pCnt As String
    Public Count As String
    Dim strKeyb As String
    Dim strKeyc As String
    Public CharaPosition As New Rectangle ' ｷｬﾗｸﾀｰ位置
    Public CharacterImg As Image = Nothing
    Public myfnt As System.Drawing.Font = Nothing
    'フォントタイプ
    Dim FNTYPE As String = "ＭＳ 明朝"
    'フォントサイズ
    Public FNTPT As Integer = 12
    '筆順表示可能フラグ
    Dim HTJflg As Boolean = False
    'くり抜き幅
    Public rctSize As Integer = 100
    '行間
    Public rowW As Integer = 120
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CheckBox1.Checked = True
        CheckBox2.Checked = True
        CheckBox3.Checked = True
        CheckBox4.Checked = True
        CheckBox5.Checked = True
    End Sub

 

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        CheckBox5.Checked = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim canvas As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        Dim gr As Graphics = PictureBox1.CreateGraphics()
        Dim gr2 As Graphics = PictureBox2.CreateGraphics()
        ListBox1.Items.Clear()
        gr.Clear(Color.White)
        canvas = New Bitmap(PictureBox2.Width, PictureBox1.Height)
        gr2.Clear(Color.White)
        '処理の終了宣言
        Label3.Text = "検索画像を消去しました。"

 
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim files As String()
        Dim grPage As Graphics()
        ListBox1.Items.Clear()
        ''描画先とするImageオブジェクトを作成する
        Dim canvas As New Bitmap(1500, PictureBox1.Height)
        myfnt = New System.Drawing.Font(FNTYPE, FNTPT)

        ' ｷｬﾗｸﾀｰ位置指定
        CharaPosition.X = 0
        CharaPosition.Y = 0
        CharaPosition.Height = rctSize
        CharaPosition.Width = rctSize

        ' グラフィック用オブジェクトを生成
        Dim gr As Graphics = PictureBox1.CreateGraphics()
        ReDim grPage(10)
        grPage(0) = gr
        grPage(0).Clear(Color.White)

        '文字確認処理
        If CheckWord_author() = False Then
            Exit Sub
        End If

        '処理の開始宣言
        Label3.Text = "検索処理中です。" & " 文字:" & KANA & ",被験者番号:" & monitID


        '被験者が空の場合
        If monitID = "" Then
            strKeyb = KANA & "_" & "*.bmp"
            'チェックボックスの確認
            files = System.IO.Directory.GetFiles( _
                  "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
            'ListBox1に結果を表示する
            ListBox1.Items.AddRange(files)

        Else
            strKeyb = KANA & "_" & monitID & "_*.bmp"
            'チェックボックス判定
            If CheckBox1.Checked Then
                strKeyb = KANA & "_" & monitID & "_1.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            If CheckBox2.Checked Then
                strKeyb = KANA & "_" & monitID & "_2.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            If CheckBox3.Checked Then
                strKeyb = KANA & "_" & monitID & "_3.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            If CheckBox4.Checked Then
                strKeyb = KANA & "_" & monitID & "_4.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            If CheckBox5.Checked Then
                strKeyb = KANA & "_" & monitID & "_5.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            If CheckBox1.Checked = False And CheckBox2.Checked = False And CheckBox3.Checked = False And CheckBox4.Checked = False And CheckBox5.Checked = False Then
                strKeyb = KANA & "_" & monitID & "_*.bmp"
                'チェックボックスの確認
                files = System.IO.Directory.GetFiles( _
                      "D:\佐々木の資料\科警研", strKeyb, System.IO.SearchOption.AllDirectories)
                'ListBox1に結果を表示する
                ListBox1.Items.AddRange(files)
            End If
            HTJflg = True
            ListBox1.Sorted = True
            End If

        If ListBox1.Items.Count > 10 Then
            Dim v As Integer = ListBox1.Items.Count / 10
            canvas = New Bitmap(1500, v * rowW)
        End If


        Dim row As Integer = 0
        For Each strList As String In ListBox1.Items
            If row = 10 Then
                CharaPosition.X = 0
                CharaPosition.Y += rowW
                row = 0
            End If
            CharacterImg = New Bitmap(strList) ' 文字画像
            ' ｷｬﾗｸﾀｰ画像の描画
            grPage(0).DrawImage(CharacterImg, CharaPosition)
            grPage(0).DrawString(FindMonit(strList), myfnt, Brushes.Black, CharaPosition.X, CharaPosition.Y + rctSize)
            CharaPosition.X += rctSize
            CharaPosition.Y += 0
            row += 1
        Next
        If HTJflg And CheckBox6.Checked Then
            hitujyunWriter()
        End If

        ''リソースを解放する
        grPage(0).Dispose()
        'PictureBox1に表示する
        PictureBox1.Image = canvas
        '処理の終了宣言
        Label3.Text = "検索処理が完了しました。" & " 文字:" & KANA & ",被験者番号:" & monitID

    End Sub

    Private Sub hitujyunWriter()

        Dim 筆跡ラインポイント(2000) As String
        Dim jj As Integer
        Dim i As Integer
        Dim jjjj As Integer
        Dim csvpath As String
        Dim csvfile As String
        Dim files As String()

        '描画先とするImageオブジェクトを作成する
        Dim canvas As New Bitmap(PictureBox2.Width, PictureBox2.Height)
        'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim g As Graphics = Graphics.FromImage(canvas)
        g.Clear(Color.White)

        'GraphicsPathオブジェクトの作成
        Dim myPath As New GraphicsPath
        strKeyc = KANA & "_" & monitID & "_" & pCnt & ".csv"

        files = System.IO.Directory.GetFiles( _
                    "D:\佐々木の資料\科警研", strKeyc, System.IO.SearchOption.AllDirectories)
        If files.Length > 1 Then
            csvfile = files(0)
            MsgBox("CSVファイルが複数あります。" & files(1))
        Else
            csvfile = files(0)
        End If

        flg = 0

        myPath.AddRectangle(New Rectangle(0 * 倍率, 0 * 倍率, 34 * 倍率, 34 * 倍率))

        ' ''CSV形式ファイルパスを設定してください。
        Dim objFile As New System.IO.StreamReader(csvfile)


        Dim strLine As String       '１行

        '次の行へ
        strLine = objFile.ReadLine()

        jj = 0
        i = 0
        While (strLine <> "")

            '行単位データをカンマ部分で分割し、配列へ格納
            jj = jj + 1
            筆跡ラインポイント(jj) = strLine
            If Mid(strLine, 1, 4) = ",,,," Then
            Else
                If Mid(strLine, 1, 2) = ",," Then
                    'x1 = 分割(3)
                    'y1 = 分割(4)
                    'MsgBox (x1 & "  " & y1)
                    i = i + 1
                End If
            End If

            '次の行へ
            strLine = objFile.ReadLine()

        End While

        'ファイルのクローズ
        objFile.Close()


        For II = 1 To i
            jjjj = 1
            For jjj = 1 To jj
                strLine = 筆跡ラインポイント(jjj)
                分割 = Split(strLine, ",")

                If Mid(strLine, 1, 4) = ",,,," Then
                    If 分割(4) >= 100 - 余白 And 分割(4) <= 134 + 余白 Then
                        分割(4) = 分割(4) - 50
                    ElseIf 分割(4) >= 150 - 余白 And 分割(4) <= 184 + 余白 Then
                        分割(4) = 分割(4) - 100
                    ElseIf 分割(4) >= 200 - 余白 And 分割(4) <= 234 + 余白 Then
                        分割(4) = 分割(4) - 150
                    ElseIf 分割(4) >= 250 - 余白 And 分割(4) <= 284 + 余白 Then
                        分割(4) = 分割(4) - 200
                    ElseIf 分割(4) >= 300 - 余白 And 分割(4) <= 334 + 余白 Then
                        分割(4) = 分割(4) - 250
                    ElseIf 分割(4) >= 350 - 余白 And 分割(4) <= 384 + 余白 Then
                        分割(4) = 分割(4) - 300
                    ElseIf 分割(4) >= 400 - 余白 And 分割(4) <= 434 + 余白 Then
                        分割(4) = 分割(4) - 350
                    ElseIf 分割(4) >= 450 - 余白 And 分割(4) <= 484 + 余白 Then
                        分割(4) = 分割(4) - 400
                    ElseIf 分割(4) >= 516 - 余白 And 分割(4) <= 550 + 余白 Then
                        分割(4) = 分割(4) - 466
                    ElseIf 分割(4) >= 566 - 余白 And 分割(4) <= 600 + 余白 Then
                        分割(4) = 分割(4) - 516
                    ElseIf 分割(4) >= 616 - 余白 And 分割(4) <= 650 + 余白 Then
                        分割(4) = 分割(4) - 566
                    End If
                    分割(4) = 分割(4) - 50

                    If 分割(5) >= 150 - 余白 And 分割(5) <= 185 + 余白 Then
                        分割(5) = 分割(5) - 66
                    ElseIf 分割(5) >= 216 - 余白 And 分割(5) <= 251 + 余白 Then
                        分割(5) = 分割(5) - 133
                    ElseIf 分割(5) >= 283 - 余白 And 分割(5) <= 318 + 余白 Then
                        分割(5) = 分割(5) - 200
                    ElseIf 分割(5) >= 349 - 余白 And 分割(5) <= 385 + 余白 Then
                        分割(5) = 分割(5) - 267
                    ElseIf 分割(5) >= 416 - 余白 And 分割(5) <= 451 + 余白 Then
                        分割(5) = 分割(5) - 333
                    ElseIf 分割(5) >= 483 - 余白 And 分割(5) <= 518 + 余白 Then
                        分割(5) = 分割(5) - 400
                    ElseIf 分割(5) >= 549 - 余白 And 分割(5) <= 585 + 余白 Then
                        分割(5) = 分割(5) - 467
                    ElseIf 分割(5) >= 616 - 余白 And 分割(5) <= 652 + 余白 Then
                        分割(5) = 分割(5) - 533
                    ElseIf 分割(5) >= 683 - 余白 And 分割(5) <= 718 + 余白 Then
                        分割(5) = 分割(5) - 600
                    ElseIf 分割(5) >= 748 - 余白 And 分割(5) <= 785 + 余白 Then
                        分割(5) = 分割(5) - 667
                    ElseIf 分割(5) >= 815 - 余白 And 分割(5) <= 852 + 余白 Then
                        分割(5) = 分割(5) - 733
                    ElseIf 分割(5) >= 883 - 余白 And 分割(5) <= 919 + 余白 Then
                        分割(5) = 分割(5) - 800
                    End If
                    分割(5) = 分割(5) - 83


                    分割(4) = (分割(4) + (II - 1) * 34) * 倍率
                    分割(5) = 分割(5) * 倍率

                    If flg = 0 Then
                        x1 = 分割(4)
                        y1 = 分割(5)
                        ''MsgBox (x1 & "  " & 分割(4) & "     " & y1 & "   " & 分割(5))
                        If II = jjjj - 1 Then
                            myPath.AddRectangle(New Rectangle((II - 1) * 34 * 倍率, 0 * 倍率, 34 * 倍率, 34 * 倍率))
                        End If

                        flg = 1
                    Else
                        x2 = 分割(4)
                        y2 = 分割(5)

                        myPath.StartFigure()
                        myPath.AddLine(x1, y1, x2, y2)
                        myPath.CloseFigure()

                        x1 = 分割(4)
                        y1 = 分割(5)
                        ''MsgBox (x1 & "  " & y1)
                    End If
                Else
                    If Mid(strLine, 1, 2) = ",," Then
                        'x1 = 分割(3)
                        'y1 = 分割(4)
                        'MsgBox (x1 & "  " & y1)
                        flg = 0
                        If jjjj > II Then Exit For
                        jjjj = jjjj + 1

                    End If
                End If

            Next jjj
        Next II

        'パス図形を描画する
        g.DrawPath(Pens.Black, myPath)

        'リソースを解放する
        g.Dispose()

        'PictureBox1に表示する
        PictureBox2.Image = canvas

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        End
    End Sub


    Public Function CheckWord_author() As Boolean
        Dim result As Boolean = True
        '１．入力チェック   =======================================================***
        '被験者番号チェック
        If TextBox4.Text.Equals("") Then
            '被験者番号は空でもOK
            monitID = ""
        Else
            Try
                Dim iA As Integer = Integer.Parse(TextBox4.Text, 0)
                If iA < 10 Then
                    TextBox4.Text = "00" & iA.ToString()
                ElseIf iA < 100 Then
                    TextBox4.Text = "0" & iA.ToString()
                End If
                monitID = TextBox4.Text
            Catch ex As Exception
                MsgBox("被験者Noは数字で入力してください。")
                Return False
                Exit Function
            End Try
        End If
        '入力ファイルチェック
        If TextBox1.Text.Equals("") Then
            MsgBox("検索文字種は必ず指定してください。")
            Return False
            Exit Function
        Else
            TextBox1.Text = StrConv(TextBox1.Text, Microsoft.VisualBasic.VbStrConv.Narrow, &H411)
            KANA = TextBox1.Text
        End If
        '検索文字ファイルの作成
        If Char.IsLower(KANA) Then
            KANA &= KANA
        End If
        '筆順選択
        If CheckBox6.Checked Then
            Dim i As Integer = 0
            If CheckBox1.Checked Then
                i += 1
                pCnt = 1
            End If
            If CheckBox2.Checked Then
                i += 1
                pCnt = 2
            End If

            If CheckBox3.Checked Then
                i += 1
                pCnt = 3
            End If

            If CheckBox4.Checked Then
                i += 1
                pCnt = 4
            End If

            If CheckBox5.Checked Then
                i += 1
                pCnt = 5
            End If
            If i > 1 Then
                MsgBox("筆順表示は回数を指定してください。")
                pCnt = ""
                Return False
            End If

        End If
        Return True
    End Function

    Private Sub PictureBox1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
        myfnt = New System.Drawing.Font(FNTYPE, FNTPT)
        Dim row As Integer = 0


        e.Graphics.Clear(Color.White)
        ' ｷｬﾗｸﾀｰ位置指定
        CharaPosition.X = 0
        CharaPosition.Y = 0
        CharaPosition.Height = rctSize
        CharaPosition.Width = rctSize
        If ListBox1.Items.Count <> 0 Then
            For Each strList As String In ListBox1.Items
                If row = 10 Then
                    CharaPosition.X = 0
                    CharaPosition.Y += rowW
                    row = 0
                End If
                CharacterImg = New Bitmap(strList) ' 文字画像
                ' ｷｬﾗｸﾀｰ画像の描画
                e.Graphics.DrawImage(CharacterImg, CharaPosition)
                e.Graphics.DrawString(FindMonit(strList), myfnt, Brushes.Black, CharaPosition.X, CharaPosition.Y + rctSize)
                CharaPosition.X += rctSize
                CharaPosition.Y += 0
                row += 1
            Next
        End If
 

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim h, w As Integer
        'ディスプレイの高さ
        h = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height
        'ディスプレイの幅
        w = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width


    End Sub

    Private Sub Panel1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles Panel1.Scroll

    End Sub
    Private Function FindMonit(ByRef dir As String) As String
        Dim name As String = ""
        Dim cntN As String = ""
        Dim textLabel As String = ""
        Dim iName As Integer = InStr(dir, ".bmp")
        cntN = dir.Substring(iName - 2, 1)
        name = dir.Substring(iName - 6, 3)
        textLabel = name & " " & cntN & "回目"
        Return textLabel
    End Function

    Private Sub PictureBox1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseClick
        TextBox2.Text = "(X=" & e.X
        TextBox3.Text = ",Y=" & e.Y & "  )"
    End Sub

End Class