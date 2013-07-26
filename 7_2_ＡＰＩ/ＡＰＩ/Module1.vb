Imports TwitterVB2
Imports System.Web
Imports System.Net
Module Module1

    Dim id As String = ""
    Dim pass As String = ""

    Sub Main()
        Dim tw As New TwitterAPI
        'Dim Url As String = tw.GetAuthorizationLink("取得したConsumerkey", "取得したConsumersecret")

        'アクセスキーの宣言
        Dim ConsumerKey As String = "zr0VDF3eODE8ACmc812mg"
        Dim ConsumerSecret As String = "gAV3E9ZFa2KH55LazweLdW6ArNDWLpNy2zQE7UiFBCU"
        Dim TokenKey As String = "1003644312-DoTcXDItPOUFj8JEWCcJHJGvnISl6wAKcOIPDWx"
        Dim TokenSecret As String = "JM2vfWiYFRSaLRegesRgW0WSv2IzmRH0AJbCzfOqk"


        Dim req As WebRequest = CreateWebRequest("http://twitter.com/")
        req.Method = "GET"
        Try
            Dim wr As HttpWebResponse
            wr = req.GetResponse()
            Dim ht As HttpStatusCode = wr.StatusCode
            wr.Close()
            ht = HttpStatusCode.OK
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
        


        Dim a As TwitterVB2.TwitterOAuth.Method
        a.ToString()
        tw.AuthenticateWith(ConsumerKey, ConsumerSecret, TokenKey, TokenSecret)
        'タイムライン取得
        Try
            Dim tp As New TwitterParameters
            tp.Add(TwitterParameterNames.Count, 50)
            For Each status As TwitterStatus In tw.HomeTimeline(tp)
                Console.WriteLine(status.User.ScreenName & " : " & status.Text)
            Next
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
        

        'つぶやき投稿
        Try
            'Dim TweetStatus As TwitterStatus = tw.Update("Hello")
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try




    End Sub


    Private Function CreateWebRequest(ByRef Uri As String) As WebRequest
        Dim req As WebRequest = HttpWebRequest.Create(Uri)
        req.ContentType = "application/x-www-form-urlencoded"
        req.Headers.Add("a")
        Return req
    End Function
    'Private Function CreateAuthString() As String()
    '    Return "Authorization: Basic " & Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(id & ":" & pass))
    'End Function



End Module
