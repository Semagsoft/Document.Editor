Public Class StartupDialog

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        If My.Settings.Options_ShowStartupDialog Then
            ShowOnStartupCheckBox.IsChecked = True
        Else
            ShowOnStartupCheckBox.IsChecked = False
        End If
    End Sub

    Private Sub OnlineHelpButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OnlineHelpButton.Click
        Process.Start("http://documenteditor.net/documentation/")
    End Sub

    Private Sub GetPluginsButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles GetPluginsButton.Click
        Process.Start("http://documenteditor.net/plugins")
    End Sub

    Private Sub WebsiteButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles WebsiteButton.Click
        Process.Start("http://documenteditor.net")
    End Sub

    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CloseButton.Click
        If ShowOnStartupCheckBox.IsChecked Then
            My.Settings.Options_ShowStartupDialog = True
        Else
            My.Settings.Options_ShowStartupDialog = False
        End If
        Close()
    End Sub

    Private Sub ReportBugButton_Click(sender As Object, e As RoutedEventArgs) Handles ReportBugButton.Click
        Try
            Process.Start("mailto:semagsoft@gmail.com")
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub DonateButton_Click(sender As Object, e As RoutedEventArgs) Handles DonateButton.Click
        Process.Start("https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=K4QJBR4UJ3W5E&lc=US&item_name=Semagsoft&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donateCC_LG%2egif%3aNonHosted")
    End Sub
End Class