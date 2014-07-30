Imports TwitterVB2
Imports System.Web
Imports System.Net
Module Module1

    Dim id As String = ""
    Dim pass As String = ""

    Sub Main()
        Dim webreq As System.Net.HttpWebRequest
        Dim tw As New TwitterAPI
        'Dim Url As String = tw.GetAuthorizationLink("取得したConsumerkey", "取得したConsumersecret")
        Dim bypassUrls() As String = {"https://api.twitter.com/1.1/statuses/home_timeline.json"}
        webreq.Proxy = _
            New System.Net.WebProxy("http://192.168.99.8:8080", True, bypassUrls)

        'アクセスキーの宣言
        Dim ConsumerKey As String = "zr0VDF3eODE8ACmc812mg"
        Dim ConsumerSecret As String = "gAV3E9ZFa2KH55LazweLdW6ArNDWLpNy2zQE7UiFBCU"
        Dim TokenKey As String = "1003644312-DoTcXDItPOUFj8JEWCcJHJGvnISl6wAKcOIPDWx"
        Dim TokenSecret As String = "JM2vfWiYFRSaLRegesRgW0WSv2IzmRH0AJbCzfOqk"

        tw.ProxyUsername()
        tw.AuthenticateWith(ConsumerKey, ConsumerSecret, TokenKey, TokenSecret)

        'タイムライン取得
        For Each Tweet As TwitterStatus In tw.HomeTimeline
            Console.WriteLine(Tweet.User.ScreenName & " : " & Tweet.Text)
        Next

        'つぶやき投稿
        Dim TweetStatus As TwitterStatus = tw.Update("ありがとう")

        'タイムライン取得（ScreenName＆個数）
        Dim tp As New TwitterParameters
        tp.Add(TwitterParameterNames.Count, 10) '取得する個数
        tp.Add(TwitterParameterNames.ScreenName, "ScreenName") '取得したいScreenName

        For Each Tweet As TwitterStatus In tw.UserTimeline(tp)
            Console.WriteLine(Tweet.User.ScreenName & " : " & Tweet.Text)
        Next




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
