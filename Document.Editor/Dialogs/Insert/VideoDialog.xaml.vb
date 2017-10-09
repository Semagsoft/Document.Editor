Public Class VideoDialog
    Public Res As String = "Cancel"
    Public w As Integer = Nothing, h As Integer = Nothing

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
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

    Private Sub TextBox1_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles TextBox1.ValueChanged, TextBox2.ValueChanged
        Try
            w = Convert.ToDouble(TextBox1.Value)
        Catch ex As Exception
            TextBox1.Value = 1
        End Try
        Try
            h = Convert.ToDouble(TextBox2.Text)
        Catch ex As Exception
            TextBox2.Value = 1
        End Try
        If TextBox1.Value > 0 AndAlso TextBox2.Value > 0 Then
            OKButton.IsEnabled = True
        Else
            OKButton.IsEnabled = False
        End If
    End Sub
End Class