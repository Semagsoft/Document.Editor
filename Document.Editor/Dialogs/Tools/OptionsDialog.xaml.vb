Imports System.Collections.ObjectModel, System.Speech.Synthesis
Public Class OptionsDialog
    Private speech As New SpeechSynthesizer

#Region "Loaded"

    Private Sub OptionsDialog_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles Me.Loaded
        StartUpComboBox.SelectedIndex = My.Settings.Options_StartupMode
        ShowStartupDialogCheckBox.IsChecked = My.Settings.Options_ShowStartupDialog
        CheckForUpdatesOnStartupCheckBox.IsChecked = My.Settings.Options_CheckForUpdatesOnStartup
        ThemeComboBox.SelectedIndex = My.Settings.Options_Theme
        EnableGlassCheckBox.IsChecked = My.Settings.Options_EnableGlass
        ComboBox1.SelectedIndex = My.Settings.Options_Tabs_SizeMode
        FontFaceComboBox.SelectedItem = My.Settings.Options_DefaultFont
        FontSizeTextBox.Value = My.Settings.Options_DefaultFontSize
        SpellCheckBox.IsChecked = My.Settings.Options_SpellCheck
        TabPlacementComboBox.SelectedIndex = My.Settings.Options_TabPlacement
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        Dim Voices As ReadOnlyCollection(Of InstalledVoice) = speech.GetInstalledVoices(Globalization.CultureInfo.CurrentCulture)
        Dim VoiceInformation As VoiceInfo = Voices(0).VoiceInfo
        For Each Voice As InstalledVoice In Voices
            VoiceInformation = Voice.VoiceInfo
            TTSComboBox.Items.Add(VoiceInformation.Name.ToString)
        Next
        TTSComboBox.SelectedIndex = My.Settings.Options_TTSV
        TTSSlider.Value = My.Settings.Options_TTSS
        CloseButtonComboBox.SelectedIndex = My.Settings.Options_Tabs_CloseButtonMode
        If My.Settings.Options_ShowRecentDocuments Then
            RecentDocumentsCheckBox.IsChecked = True
        Else
            RecentDocumentsCheckBox.IsChecked = False
        End If
        RulerMeasurementComboBox.SelectedIndex = My.Settings.Options_RulerMeasurement

        TemplatesFolderTextBlock.Text = My.Settings.Options_TemplatesFolder
        TemplatesFolderTextBlock.ToolTip = My.Settings.Options_TemplatesFolder
        'For Each f As String In My.Computer.FileSystem.GetFiles(My.Settings.Options_TemplatesFolder)
        '    Dim template As New IO.FileInfo(f)
        '    If template.Extension = ".xaml" Then
        '        Dim item As New ListBoxItem
        '        item.Height = 32
        '        item.Content = template.Name.Remove(template.Name.Length - 5)
        '        TemplatesListBox.Items.Add(item)
        '    End If
        'Next

        If My.Settings.Options_EnablePlugins Then
            PluginsCheckBox.IsChecked = True
        Else
            PluginsCheckBox.IsChecked = False
        End If
        For Each f As String In My.Computer.FileSystem.GetFiles(My.Application.Info.DirectoryPath + "\Plugins\")
            Dim plugin As New IO.FileInfo(f)
            Dim item As New ListBoxItem
            item.Height = 32
            item.Content = plugin.Name.Remove(plugin.Name.Length - 3)
            PluginsListBox.Items.Add(item)
        Next
    End Sub

#End Region

#Region "DialogButtons"

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles OKButton.Click
        My.Settings.Options_StartupMode = StartUpComboBox.SelectedIndex
        My.Settings.Options_ShowStartupDialog = ShowStartupDialogCheckBox.IsChecked
        My.Settings.Options_CheckForUpdatesOnStartup = CheckForUpdatesOnStartupCheckBox.IsChecked
        My.Settings.Options_Theme = ThemeComboBox.SelectedIndex
        My.Settings.Options_EnableGlass = EnableGlassCheckBox.IsChecked
        My.Settings.Options_DefaultFont = TryCast(FontFaceComboBox.SelectedItem, FontFamily)
        My.Settings.Options_DefaultFontSize = Convert.ToInt16(FontSizeTextBox.Value)
        If SpellCheckBox.IsChecked Then
            My.Settings.Options_SpellCheck = True
        Else
            My.Settings.Options_SpellCheck = False
        End If
        My.Settings.Options_TabPlacement = TabPlacementComboBox.SelectedIndex
        My.Settings.Options_TTSV = TTSComboBox.SelectedIndex
        My.Settings.Options_TTSS = Convert.ToInt16(TTSSlider.Value)
        My.Settings.Options_Tabs_CloseButtonMode = CloseButtonComboBox.SelectedIndex
        If RecentDocumentsCheckBox.IsChecked Then
            My.Settings.Options_ShowRecentDocuments = True
        Else
            My.Settings.Options_ShowRecentDocuments = False
        End If
        My.Settings.Options_RulerMeasurement = RulerMeasurementComboBox.SelectedIndex
        My.Settings.Options_TemplatesFolder = TemplatesFolderTextBlock.Text
        If PluginsCheckBox.IsChecked Then
            My.Settings.Options_EnablePlugins = True
        Else
            My.Settings.Options_EnablePlugins = False
        End If
        Close()
    End Sub

    Private Sub CancelButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles CancelButton.Click
        Close()
    End Sub

    Private Sub ResetButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ResetButton.Click
        Dim m As New MessageBoxDialog(My.Application.Info.ProductName + " needs to restart, restart now?", "Are You Sure?", "YesNo", Nothing)
        m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
        m.Owner = Me
        m.ShowDialog()
        If m.Result = "Yes" Then
            My.Settings.Reset()
            My.Application.Shutdown()
        End If
    End Sub

#End Region

#Region "Plugins"

    Private Sub PluginAddButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles PluginAddButton.Click
        Dim dialog As New Microsoft.Win32.OpenFileDialog
        dialog.Title = "Add Plugin"
        dialog.Filter = "Plugins(*.vb)|*.vb"
        Me.Hide()
        If dialog.ShowDialog Then
            Dim fileinfo As New IO.FileInfo(dialog.FileName)
            My.Computer.FileSystem.CopyFile(dialog.FileName, My.Application.Info.DirectoryPath + "\Plugins\" + fileinfo.Name)
            PluginsListBox.Items.Add(fileinfo.Name.Remove(fileinfo.Name.Length - 3))
        End If
        Me.ShowDialog()
    End Sub

    Private Sub PluginRemoveButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles PluginRemoveButton.Click
        My.Computer.FileSystem.DeleteFile(My.Application.Info.DirectoryPath + "\Plugins\" + PluginsListBox.SelectedItem.ToString + ".vb")
        PluginsListBox.Items.Remove(PluginsListBox.SelectedItem)
    End Sub

    Private Sub PluginsFolderButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles PluginsFolderButton.Click
        Try
            Process.Start(My.Application.Info.DirectoryPath + "\Plugins")
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub PluginsListBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles PluginsListBox.SelectionChanged
        If PluginsListBox.SelectedItem IsNot Nothing Then
            PluginRemoveButton.IsEnabled = True
        Else
            PluginRemoveButton.IsEnabled = False
        End If
    End Sub

#End Region

#Region "Templates"

    Private Sub SetTemplatesFolderButton_Click(sender As Object, e As RoutedEventArgs) Handles SetTemplatesFolderButton.Click
        Dim folderDialog As New Forms.FolderBrowserDialog()
        folderDialog.Description = "Set Templates Folder"
        If My.Computer.FileSystem.DirectoryExists(My.Settings.Options_TemplatesFolder) Then
            folderDialog.SelectedPath = My.Settings.Options_TemplatesFolder
        End If

        If folderDialog.ShowDialog Then
            TemplatesFolderTextBlock.Text = folderDialog.SelectedPath
            TemplatesFolderTextBlock.ToolTip = folderDialog.SelectedPath
        End If
    End Sub

    Private Sub AddTemplateButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles AddTemplateButton.Click
        Dim dialog As New Microsoft.Win32.OpenFileDialog
        dialog.Title = "Add Template"
        dialog.Filter = "Templates(*.xaml)|*.xaml"
        Me.Hide()
        If dialog.ShowDialog Then
            Dim fileinfo As New IO.FileInfo(dialog.FileName)
            My.Computer.FileSystem.CopyFile(dialog.FileName, My.Settings.Options_TemplatesFolder + fileinfo.Name)
            TemplatesListBox.Items.Add(fileinfo.Name.Remove(fileinfo.Name.Length - 5))
        End If
        Me.ShowDialog()
    End Sub

    Private Sub RemoveTemplateButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles RemoveTemplateButton.Click
        My.Computer.FileSystem.DeleteFile(My.Settings.Options_TemplatesFolder + TemplatesListBox.SelectedItem.ToString + ".xaml")
        TemplatesListBox.Items.Remove(TemplatesListBox.SelectedItem)
    End Sub

    Private Sub TemplatesFolderButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles TemplatesFolderButton.Click
        Try
            Process.Start(My.Settings.Options_TemplatesFolder)
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub TemplatesListBox_SelectionChanged(ByVal sender As Object, ByVal e As SelectionChangedEventArgs) Handles TemplatesListBox.SelectionChanged
        If TemplatesListBox.SelectedItem IsNot Nothing Then
            RemoveTemplateButton.IsEnabled = True
        Else
            RemoveTemplateButton.IsEnabled = False
        End If
    End Sub

#End Region

    Private Sub ClearRecentButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles ClearRecentButton.Click
        My.Settings.Options_RecentFiles.Clear()
        Dim m As New MessageBoxDialog("Documents Cleared", "Cleared", Nothing, Nothing)
        m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
        m.Owner = Me
        m.ShowDialog()
    End Sub
End Class