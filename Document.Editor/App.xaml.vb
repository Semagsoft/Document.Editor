Class Application
    Public StartUpFileNames As New Collection

    Private Sub Application_Startup(ByVal sender As Object, ByVal e As System.Windows.StartupEventArgs) Handles Me.Startup
        If e.Args.Length > 0 Then
            For Each s As String In e.Args
                StartUpFileNames.Add(s)
            Next
        End If
    End Sub
End Class