Public Class SpellCheckDialog
    Public Res As String = "Cancel"

    Private Sub SpellCheckDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub

    Private Sub WordListBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles WordListBox.KeyDown
        If e.Key = Key.Enter And WordListBox.SelectedItem IsNot Nothing Then
            OKButton_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub WordListBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles WordListBox.SelectionChanged
        If WordListBox.SelectedItem IsNot Nothing Then
            OKButton.IsEnabled = True
        Else
            OKButton.IsEnabled = False
        End If
    End Sub
End Class
