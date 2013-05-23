Module Module1
    ReadOnly log As log4net.ILog = _
    log4net.LogManager.GetLogger( _
    System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)


    Sub Main()
        log.Debug("Program Started")
        log.Warn("警告です")
        log.InfoFormat("mymethod:今日の日付 Date:{0}", DateTime.Now)


    End Sub

End Module

Public Class Logtest



End Class


