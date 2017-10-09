Imports MS.WindowsAPICodePack.Internal.CoreHelpers
Public Class LinkDialog
    Public Res As String = "Cancel"
    Public Link As String = Nothing

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Link = TextBox1.Text
        Res = "OK"
        Close()
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles TextBox1.KeyDown
        If e.Key = Key.Enter AndAlso TextBox1.Text.Length > 0 Then
            OKButton_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text IsNot Nothing Then
            OKButton.IsEnabled = True
        Else
            OKButton.IsEnabled = False
        End If
    End Sub

    Private Sub LinkDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        TextBox1.Focus()
    End Sub
End Class