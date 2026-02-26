Public Class VbGreeter
    Public Function GetMessage(ByVal name As String) As String
        Return $"Hello from VB.NET: {name}"
    End Function

    Public Sub Run()
        Dim message As String = GetMessage("Developer")
        Dim length As Integer = message.Length
    End Sub
End Class
