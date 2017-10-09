Public Class MessageBoxDialog
    Public Result As String = "Cancel"

    Public Sub New(text As String, title As String, buttons As String, icon As Integer)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        MessageBoxText.Text = text
        Me.Title = title
        If buttons = "YesNo" Then
            OKButton.Visibility = Visibility.Collapsed
            YesButton.Visibility = Visibility.Visible
            NoButton.Visibility = Visibility.Visible
            YesButton.IsDefault = True
        ElseIf buttons = "YesNoCancel" Then
            OKButton.Visibility = Visibility.Collapsed
            YesButton.Visibility = Visibility.Visible
            NoButton.Visibility = Visibility.Visible
            CancelButton.Visibility = Visibility.Visible
            YesButton.IsDefault = True
        End If

        If icon = Nothing Then
            Dim iconUri As Uri = New Uri("pack://application:,,,/app.ico", UriKind.RelativeOrAbsolute)
            Me.Icon = BitmapFrame.Create(iconUri)
        End If
    End Sub

    Private Sub OKButton_Click(sender As Object, e As RoutedEventArgs) Handles OKButton.Click
        Result = "OK"
        DialogResult = True
        Close()
    End Sub

    Private Sub YesButton_Click(sender As Object, e As RoutedEventArgs) Handles YesButton.Click
        Result = "Yes"
        DialogResult = True
        Close()
    End Sub

    Private Sub NoButton_Click(sender As Object, e As RoutedEventArgs) Handles NoButton.Click
        Result = "No"
        DialogResult = True
        Close()
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As RoutedEventArgs) Handles CancelButton.Click
        Result = "Cancel"
        DialogResult = True
        Close()
    End Sub
End Class