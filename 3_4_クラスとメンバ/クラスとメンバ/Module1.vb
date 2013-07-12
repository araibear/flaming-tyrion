Module Module1
    '数字でない文字列を入力した場合に変換する数値：デフォルト：0
    Public ERRNUM As Integer = 1

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
        height = ParseInt_s(InStr1, ERRNUM)
        bottom = ParseInt_s(InStr2, ERRNUM)
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
        pi = ParseDbl_s(InStr3, 3.14)
        Tri2.Pi = pi
        Console.WriteLine("円周率は" & Tri2.Pi & "です")
        ' 斜辺を２で割って、半径とし、円の面積を求める
        area = (Tri2.Pi) * (hypo / 2) ^ 2
        tr.DisplayCircle(area)
    End Sub
    '***************************************************************
    ' 機能：入力文字列を安全に数字変換する関数
    ' 第１引数：変換する文字列
    ' 第２引数：エラーだった場合に置き換える数字
    ' 戻値：置き換わった結果
    '******1*********2*********3*********4*********5**********6*****
    Private Function ParseInt_s(ByRef strNum As String, ByRef errNum As Integer)
        Dim res As Integer
        If (Integer.TryParse(strNum, res)) Then
        Else
            res = errNum
            Console.WriteLine("********** ERROR :{0} 数字ではありません。 **********", strNum)
        End If
        Return res
    End Function
    Private Function ParseDbl_s(ByRef strNum As String, ByRef errNum As Double)
        Dim res As Double
        If (Double.TryParse(strNum, res)) Then
        Else
            res = errNum
            Console.WriteLine("********** ERROR :{0} 数字ではありません。 **********", strNum)
        End If
        Return res
    End Function

End Module
'==============================================================
' 三角形のもろもろを計算したり、表示したりするクラス
' フィールド：_Pi
' メソッド：TriangleArea,TriangleHypo,DisplayAns,DisplayHypo,DisplayCircle
' プロパティ：Pi

Public Class Triangle
    '==============================================================
    ' 円周率フィールドく
    ' 円周率（定数）を隠しておく
    '==============================================================
    Private _Pi As Double = 3.14


    '***************************************************************
    ' 機能：三角形の面積を求める関数
    ' 引数：height高さ、bottom底辺
    ' 戻値：面積
    '******1*********2*********3*********4*********5**********6*****
    Public Function TriangleArea(ByRef height As Integer, ByRef bottom As Integer)
        Dim area As Integer
        area = height * bottom / 2
        Return area
    End Function
    '***************************************************************
    ' 機能：三角形の斜辺を求める関数
    ' 引数：height高さ、bottom底辺
    ' 戻値：斜辺
    '******1*********2*********3*********4*********5**********6*****

    Public Function TriangleHypo(ByRef height As Integer, ByRef bottom As Integer)
        Dim hypo As Double
        hypo = Math.Sqrt(height ^ 2 + bottom ^ 2)
        Return hypo
    End Function

    '***************************************************************
    ' 機能：メッセージを出力する関数
    ' 引数：表示する値
    ' 戻値：なし
    '******1*********2*********3*********4*********5**********6*****

    Public Sub DisplayAns(ByRef ans As Integer)
        Console.WriteLine("三角形の面積は" & ans & "cm2です")
    End Sub

    Public Sub DisplayHypo(ByRef ans As Double)
        Console.WriteLine("三角形の斜辺は" & ans & "cmです")
    End Sub

    Public Sub DisplayCircle(ByRef ans As Double)
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
