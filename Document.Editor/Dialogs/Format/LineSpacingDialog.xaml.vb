Public Class LineSpacingDialog
    Public Res As String = "Cancel"
    Public number As Double

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles TextBox1.KeyDown
        If e.Key = Key.Enter Then
            If OKButton.IsEnabled Then
                OKButton_Click(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub

    Private Sub LineSpacingDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        TextBox1.Value = number
        TextBox1.Focus()
    End Sub

    Private Sub TextBox1_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TextBox1.ValueChanged
        Try
            number = TextBox1.Value
        Catch ex As Exception
        End Try
    End Sub
End Class