Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim hBitmap As New System.Drawing.Bitmap("D:\佐々木の資料\科警研\いろファイル\tate_green.bmp")
        Dim hColor As System.Drawing.Color = Nothing
        Dim Writer As IO.StreamWriter = Nothing
        Dim num16 As Integer
        Dim str As String = ""

        Writer = New IO.StreamWriter("D:\佐々木の資料\科警研\いろファイル\tate_green.csv", False, System.Text.Encoding.Default)
        For i As Integer = 0 To hBitmap.Width - 1
            For j As Integer = 0 To hBitmap.Height - 1
                hColor = hBitmap.GetPixel(i, j)
                str = i & "," & j & ","
                str &= Convert.ToInt32(hColor.R, 10) & ","
                str &= Convert.ToInt32(hColor.G, 10) & ","
                str &= Convert.ToInt32(hColor.B, 10)
                Writer.WriteLine(str)
            Next
        Next
        Writer.Close()
 

    End Sub
End Class
