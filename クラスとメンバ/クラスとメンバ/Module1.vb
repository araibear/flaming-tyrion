Module Module1

    Sub Main()
        Dim InStr1, InStr2, InStr3 As String
        Dim height, bottom As String
        Dim area, hypo, pi As Double


        ' 直角三角形の面積を求める-------------------------------------*

        Console.WriteLine("高さを入力してEnterキーを押してください")
        InStr1 = Console.ReadLine()
        Console.WriteLine("底辺を入力してEnterキーを押してください")
        InStr2 = Console.ReadLine()

        ' 文字列を整数に変換して配列に保存
        height = Integer.Parse(InStr1)
        bottom = Integer.Parse(InStr2)
        'コンストラクタの作成
        Dim tr As Triangle = New Triangle()
        area = tr.TriangleArea(height, bottom)
        tr.DisplayAns(area)

        ' 直角三角形の斜辺を求める-------------------------------------*
        hypo = tr.TriangleHypo(height, bottom)
        tr.DisplayHypo(hypo)
        ' 直角三角形の３点をめぐる円の面積を求める------------------------*
        ' このような形でオブジェクトを定義することも可能
        Dim Tri2 As New Triangle
        Console.WriteLine("円周率は" & Tri2.Pi & "です")
        Console.WriteLine("細かい円周率に直してEnterキーを押してください。Pi = 301415")
        InStr3 = Console.ReadLine()
        ' 文字列を整数に変換して配列に保存
        pi = Double.Parse(InStr1)
        Tri2.Pi = pi
        Console.WriteLine("円周率は" & Tri2.Pi & "です")
        ' 斜辺を２で割って、半径とし、円の面積を求める
        area = (Tri2.Pi) * (hypo / 2) ^ 2
        tr.DisplayCircle(area)
    End Sub

End Module
'==============================================================
' 三角形のもろもろを計算したり、表示したりするクラスく
' フィールド：_Pi
' メソッド：TriangleArea,TriangleHypo,DisplayAns,DisplayHypo,DisplayCircle
' プロパティ：Pi

Public Class Triangle
    '==============================================================
    ' 円周率フィールドく
    ' 円周率（定数）を隠しておく
    '==============================================================
    Dim _Pi As Double = 3.14


    '***************************************************************
    ' 機能：三角形の面積を求める関数
    ' 引数：height高さ、bottom底辺
    ' 戻値：面積
    '******1*********2*********3*********4*********5**********6*****
    Public Function TriangleArea(height As Integer, bottom As Integer)
        Dim area As Integer
        area = height * bottom / 2
        Return area
    End Function
    '***************************************************************
    ' 機能：三角形の斜辺を求める関数
    ' 引数：height高さ、bottom底辺
    ' 戻値：斜辺
    '******1*********2*********3*********4*********5**********6*****

    Public Function TriangleHypo(height As Integer, bottom As Integer)
        Dim hypo As Double
        hypo = Math.Sqrt(height ^ 2 + bottom ^ 2)
        Return hypo
    End Function

    '***************************************************************
    ' 機能：メッセージを出力する関数
    ' 引数：表示する値
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****

    Public Sub DisplayAns(ans As Integer)
        Console.WriteLine("三角形の面積は" & ans & "cm2です")
    End Sub

    Public Sub DisplayHypo(ans As Double)
        Console.WriteLine("三角形の斜辺は" & ans & "cmです")
    End Sub

    Public Sub DisplayCircle(ans As Double)
        Console.WriteLine("三角形の外円斜の面積は" & ans & "cm2です")
    End Sub


    '*-------------------------------------------------------------*
    ' プロパティ名：円周率アクセサ
    ' 取出：円周率_pu
    ' 書込：円周率_pu
    '*-------------------------------------------------------------*
    Public Property Pi() As Double
        Get
            Console.WriteLine("Piプロパティの値を読み取ります。")
            Return _Pi
        End Get
        Set(ByVal value As Double)
            Console.WriteLine("Piプロパティに値を書き込みます。")
            _Pi = value
        End Set
    End Property
End Class
