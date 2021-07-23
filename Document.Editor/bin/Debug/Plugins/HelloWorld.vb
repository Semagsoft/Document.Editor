Public Class HelloWorld

    Public Function IsStartupPlugin() As Boolean
        Return False
    End Function

    Public Function IsEventPlugin() As Boolean
        Return True
    End Function

    Public Function StartPlugin() As Object
        Dim i As String = "hello world!", win As New Form
        win.Text = i
        win.ShowDialog()
        Return i
    End Function

End Class