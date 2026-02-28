Public Class VbGreeter
    Public Function GetMessage(ByVal name As String) As String
        Return $"Hello from VB.NET: {name}"
    End Function

    Public Sub Run()
        Dim message As String = GetMessage("Developer")
        Dim localValue As Integer = message.Length
        Dim result As String = localValue.ToString()
    End Sub
End Class
