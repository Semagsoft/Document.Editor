Partial Public Class AboutDialog
    Private licenseicon As New BitmapImage(New Uri("pack://application:,,,/Images/Help/license16.png"))
    Private backicon As New BitmapImage(New Uri("pack://application:,,,/Images/Help/back16.png"))
    Public IsCheckingForUpdates As Boolean = False

#Region "Reuseable Code"

    Public Function PathExists(ByVal path As String, ByVal timeout As Integer) As Boolean
        Dim exists As Boolean = True
        Dim t As New Threading.Thread(DirectCast(Function() CheckPathFunction(path), Threading.ThreadStart))
        t.Start()
        Dim completed As Boolean = t.Join(timeout)
        If Not completed Then
            exists = False
            t.Abort()
        End If
        Return exists
    End Function

    Public Function CheckPathFunction(ByVal path As String) As Boolean
        Return IO.File.Exists(path)
    End Function

#End Region

#Region "Loaded"

    Private Sub AboutDialog_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        TextBox1.Text = My.Resources.License
        AppNameLabel.Content = My.Application.Info.ProductName + " " + My.Application.Info.Version.Major.ToString
        VersionLabel.Content = "Version: " + My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString
        CopyLabel.Content = My.Application.Info.Copyright.ToString + " By Semagsoft"
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        If IsCheckingForUpdates Then
            Title = "Update"
            ShowInTaskbar = True
            UpdateButton_Click(Nothing, Nothing)
        End If
    End Sub

#End Region

#Region "License"

    Private Sub LicenseButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles LicenseButton.Click
        If TextBox1.IsVisible Then
            TextBox1.Visibility = Visibility.Hidden
            UpdateButton.Visibility = Visibility.Visible
            LicenseButton.Icon = licenseicon
            LicenseButton.Header = "License"
        Else
            TextBox1.Visibility = Visibility.Visible
            UpdateButton.Visibility = Visibility.Collapsed
            LicenseButton.Icon = backicon
            LicenseButton.Header = "Back"
        End If
    End Sub

#End Region

#Region "Check For Update"

    Private WithEvents CheckForUpdateWorker As New ComponentModel.BackgroundWorker

    Public Sub UpdateButton_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles UpdateButton.Click
        If My.Computer.Network.IsAvailable Then
            AppLogo.Visibility = Visibility.Hidden
            LicenseButton.Visibility = Visibility.Hidden
            OKButton.Visibility = Visibility.Hidden
            UpdateButton.Visibility = Visibility.Hidden
            UpdateBox.Visibility = Visibility.Visible
            CheckForUpdateWorker.RunWorkerAsync()
        Else
            Dim m As New MessageBoxDialog("Internet not found.", "Error", Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End If
    End Sub

    Private Sub CheckForUpdateWorker_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles CheckForUpdateWorker.DoWork
        Try
            If PathExists("http://semagsoft.com/software/updates/document.editor.update", 5000) Then
                Dim fileReader As New Net.WebClient
                My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\")
                Dim filename As String = My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\updatechecker.ini"
                If IO.File.Exists(filename) Then
                    IO.File.Delete(filename)
                End If
                fileReader.DownloadFile(New Uri("http://semagsoft.com/software/updates/document.editor.update"), filename)
                e.Result = True
            Else
                e.Result = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ex.ToString, MessageBoxButton.OK, MessageBoxImage.Error)
            e.Result = False
        End Try
    End Sub

    Private Sub CheckForUpdateWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles CheckForUpdateWorker.RunWorkerCompleted
        If e.Result = True Then
            Dim textreader As IO.TextReader = IO.File.OpenText(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\updatechecker.ini")
            Dim version As String = textreader.ReadLine
            Dim versionyear As Integer = Convert.ToInt16(version.Substring(0, 4))
            Dim versionnumber As Integer = Convert.ToInt16(version.Substring(5))
            If versionyear >= My.Application.Info.Version.Major AndAlso versionnumber > My.Application.Info.Version.Minor Then
                Dim whatsnew As New Collection
                Dim line As String
                Do
                    line = textreader.ReadLine
                    If line IsNot Nothing Then
                        whatsnew.Add(line.ToString)
                    End If
                Loop Until line Is Nothing
                textreader.Close()
                UpdateText.Text = "An update(" + version + ") was found, do you want to apply it?"
                ProgressBox.Visibility = Visibility.Collapsed
                ApplyUpdateButton.Visibility = Visibility.Visible
                CancelUpdateButton.Visibility = Visibility.Visible
                WhatsNewTextBox.Clear()
                For Each s As String In whatsnew
                    WhatsNewTextBox.AppendText(s + vbNewLine)
                Next
                WhatsNewTextBlock.Visibility = Visibility.Visible
                WhatsNewTextBox.Visibility = Visibility.Visible
            Else
                Dim m As New MessageBoxDialog("Document.Editor is already up to date", "Up To Date", Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
                m.Owner = Me
                m.ShowDialog()
                UpdateBox.Visibility = Visibility.Hidden
                AppLogo.Visibility = Visibility.Visible
                LicenseButton.Visibility = Visibility.Visible
                OKButton.Visibility = Visibility.Visible
                UpdateButton.Visibility = Visibility.Visible
                If IsCheckingForUpdates Then
                    Close()
                End If
            End If
            textreader.Close()
        Else
            If IsCheckingForUpdates Then
                Close()
            End If
        End If
    End Sub

#End Region

#Region "Update"

    Public WithEvents webmanager As New Net.WebClient

    Private Sub ApplyUpdateButton_Click(sender As Object, e As RoutedEventArgs) Handles ApplyUpdateButton.Click
        UpdateProgressbar.IsIndeterminate = False
        UpdateProgressbar.Minimum = 0
        UpdateProgressbar.Maximum = 100
        UpdateProgressbar.Value = 0
        ProgressBox.Visibility = Visibility.Visible
        If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe") Then
            IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe")
        End If
        webmanager.DownloadFileAsync(New Uri("http://semagsoft.com/software/downloads/Document.Editor_Setup.exe"), My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe")
        UpdateText.Text = "Downloading Update, Please Wait..."
        ApplyUpdateButton.Visibility = Visibility.Collapsed
    End Sub

    Private Sub CancelUpdateButton_Click(sender As Object, e As RoutedEventArgs) Handles CancelUpdateButton.Click
        While webmanager.IsBusy
            webmanager.CancelAsync()
        End While

        If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe") Then
            IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe")
        End If
        AppLogo.Visibility = Visibility.Visible
        LicenseButton.Visibility = Visibility.Visible
        OKButton.Visibility = Visibility.Visible
        UpdateButton.Visibility = Visibility.Visible
        UpdateBox.Visibility = Visibility.Collapsed
        UpdateText.Text = "Checking for updates..."
        UpdateProgressbar.IsIndeterminate = True
        FilesizeTextBlock.Visibility = Visibility.Collapsed
        ApplyUpdateButton.Visibility = Visibility.Collapsed
        CancelUpdateButton.Visibility = Visibility.Collapsed
        WhatsNewTextBlock.Visibility = Visibility.Collapsed
        WhatsNewTextBox.Visibility = Visibility.Collapsed
    End Sub

    Public Sub webmanager_DownloadProgressChanged(sender As Object, e As Net.DownloadProgressChangedEventArgs) Handles webmanager.DownloadProgressChanged
        UpdateProgressbar.Value = e.ProgressPercentage
        FilesizeTextBlock.Visibility = Visibility.Visible
        Dim downloadedfilesize As Integer = e.BytesReceived / 1024
        Dim totalfilesize As Integer = e.TotalBytesToReceive / 1024
        FilesizeTextBlock.Text = downloadedfilesize.ToString + " KB" + "/" + totalfilesize.ToString + " KB"
    End Sub

    Public Sub webmanager_DownloadFileCompleted(ByVal sender As Object, ByVal e As ComponentModel.AsyncCompletedEventArgs) Handles webmanager.DownloadFileCompleted
        If Not e.Cancelled Then
            Try
                Process.Start(My.Computer.FileSystem.SpecialDirectories.Temp + "\Semagsoft\DocumentEditor\setup.exe", "/D=" + My.Application.Info.DirectoryPath)
                My.Application.Shutdown()
            Catch ex As Exception
                If ex.Message.EndsWith("canceled by the user") Then
                    Dim m As New MessageBoxDialog("The update has been canceled!", "Update Canceled", Nothing, Nothing)
                    m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/info32.png"))
                    m.Owner = Me
                    m.ShowDialog()
                    CancelUpdateButton_Click(Nothing, Nothing)
                Else
                    Dim m As New MessageBoxDialog("Error running update installer", "Error", Nothing, Nothing)
                    m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                    m.Owner = Me
                    m.ShowDialog()
                End If
            End Try
        End If
    End Sub

#End Region

End Class