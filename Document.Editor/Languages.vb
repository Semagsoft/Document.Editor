Public Class Languages

    Public NewString As String, OpenString As String, RecentString As String, PrintString As String

    Public Sub New(LangCode As String)
        If LangCode = "pt-br" Then
            NewString = "Novo"
            OpenString = "Aberto"
        End If
    End Sub
End Class