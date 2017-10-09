Public Class InsertLineDialog
    Public Res As String = "Cancel", h As Integer

    'Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TextBox1.TextChanged
    '    Try
    '        h = Convert.ToInt32(TextBox1.Text)
    '    Catch ex As Exception
    '        TextBox1.Clear()
    '    End Try
    '    If TextBox1.Text.Length > 0 Then
    '        OKButton.IsEnabled = True
    '    Else
    '        OKButton.IsEnabled = False
    '    End If
    'End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        h = Convert.ToInt32(TextBox1.Value)
        Res = "OK"
        Close()
    End Sub

    Private Sub InsertLineDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        TextBox1.Focus()
    End Sub
End Class