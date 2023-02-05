Public Class Util

    Shared Function parseOptions(ByVal optionString As String) As Dictionary(Of String, String)
        Dim options As New Dictionary(Of String, String)
        For Each optionString In optionString.Split("|")
            Dim optSplit = optionString.Split("=")
            options.Add(optSplit(0), optSplit(1))
        Next
        Return options
    End Function

End Class
