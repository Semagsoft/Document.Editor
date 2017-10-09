Public Class DateDialog
    Public Res As String = "Cancel"

    Private Sub DateDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        ListBox1.Items.Add(Date.Now.ToString("M/dd/yyyy"))
        ListBox1.Items.Add(Date.Now.ToString("M/dd/yy"))
        ListBox1.Items.Add(Date.Now.ToString("MM/dd/yy"))
        ListBox1.Items.Add(Date.Now.ToString("MM/dd/yyyy"))
        ListBox1.Items.Add(Date.Now.ToString("yy/MM/dd"))
        ListBox1.Items.Add(Date.Now.ToString("yyyy-MM-dd"))
        ListBox1.Items.Add(Date.Now.ToString("dd-MMM-yy"))
        ListBox1.Items.Add(Date.Now.ToString("dddd, MMMM dd, yyyy"))
        ListBox1.Items.Add(Date.Now.ToString("MMMM dd, yyyy"))
        ListBox1.Items.Add(Date.Now.ToString("dddd, dd MMMM, yyyy"))
        ListBox1.Items.Add(Date.Now.ToString("dd MMMM, yyyy"))

        ListBox1.SelectedIndex = 0
        ListBox1.Focus()
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub
End Class