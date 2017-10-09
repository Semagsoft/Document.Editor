Imports MS.WindowsAPICodePack.Internal.CoreHelpers
Partial Public Class LoadFileDialog

    Public i As Boolean = False

    Private Sub LoadFileDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If i = False Then
            e.Cancel = True
        End If
    End Sub

    Private Sub LoadFileDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
    End Sub
End Class