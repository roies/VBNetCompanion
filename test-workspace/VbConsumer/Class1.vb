Imports CSharpLib

Public Class VbWorkflow
	Private ReadOnly _service As New GreeterService()

	Public Function BuildGreeting(ByVal name As String) As String
		Return _service.FormatMessage(name)
	End Function

	Public Function ComputeTotal(ByVal left As Integer, ByVal right As Integer) As Integer
		Return _service.Add(left, right)
	End Function
End Class
