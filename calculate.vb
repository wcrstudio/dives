
Public Class calculate
    Sub New()
        Main()
    End Sub
    Public Function Fisher(ByVal X As Double) As Double
        Try
            Fisher = (1 + X) / (1 - X)
            Fisher = Math.Log(Fisher, Math.E)
            Return 0.5 * Fisher
        Catch
            Return 0
        End Try
    End Function

    Public Function Factorial(ByVal Num As Integer) As Double
        If Num <= 1 Then
            Return (1)
        Else
            Return Num * Factorial(Num - 1)
        End If
    End Function
   
End Class
