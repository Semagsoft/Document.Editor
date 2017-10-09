Public Class TimeDialog
    Public Res As String = "Cancel", Time As String = Nothing

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub

    Private Sub RadioButton12_Checked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RadioButton12.Checked
        If AMPMCheckBox IsNot Nothing Then
            AMPMCheckBox.IsEnabled = True
        End If
    End Sub

    Private Sub RadioButton12_Unchecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RadioButton12.Unchecked
        AMPMCheckBox.IsEnabled = False
    End Sub

    Private Sub TimeDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
    End Sub
End Class