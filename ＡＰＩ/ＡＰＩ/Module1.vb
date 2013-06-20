Imports TwitterVB2

Module Module1



    Sub Main()
        Dim tw As New TwitterAPI
        'Dim Url As String = tw.GetAuthorizationLink("取得したConsumerkey", "取得したConsumersecret")

        'アクセスキーの宣言
        Dim ConsumerKey As String = "zr0VDF3eODE8ACmc812mg"
        Dim ConsumerSecret As String = "gAV3E9ZFa2KH55LazweLdW6ArNDWLpNy2zQE7UiFBCU"
        Dim TokenKey As String = "1003644312-DoTcXDItPOUFj8JEWCcJHJGvnISl6wAKcOIPDWx"
        Dim TokenSecret As String = "JM2vfWiYFRSaLRegesRgW0WSv2IzmRH0AJbCzfOqk"


        tw.AuthenticateWith(ConsumerKey, ConsumerSecret, TokenKey, TokenSecret)
        'タイムライン取得
        Try
            For Each Tweet As TwitterStatus In tw.HomeTimeline
                Console.WriteLine(Tweet.User.ScreenName & " : " & Tweet.Text)
            Next
        Catch ex As Exception
            'Throw ex
        End Try
        

        'つぶやき投稿
        'Dim TweetStatus As TwitterStatus = tw.Update("Hello")



    End Sub

End Module
