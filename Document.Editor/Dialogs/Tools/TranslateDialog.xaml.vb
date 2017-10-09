Imports MS.WindowsAPICodePack.Internal.CoreHelpers
Public Class TranslateDialog

    Private intContent As String
    Public res As Boolean = False

    Public Sub New(ByVal s As String)
        InitializeComponent()
        intContent = s
    End Sub

    Private Sub TranslateButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles TranslateButton.Click
        Try
            If My.Computer.Network.IsAvailable Then
                Dim froml As New Microsoft.DetectedLanguage, tol As New Microsoft.Language
                froml.Code = FromBox.Items.Item(FromBox.SelectedIndex).Tag
                tol.Code = ToBox.Items.Item(ToBox.SelectedIndex).Tag
                Dim translator As New Semagsoft.Translator.TranslatorHelper
                Dim transres As String = translator.Translate(intContent, froml, tol)
                TranslatedText.Content = transres
                OKButton.IsEnabled = True
            Else
                Dim m As New MessageBoxDialog("No Internet Found", "Error", Nothing, Nothing)
                m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
                m.Owner = Me
                m.ShowDialog()
            End If
        Catch ex As Exception
            Dim m As New MessageBoxDialog(ex.Message, ex.ToString, Nothing, Nothing)
            m.MessageImage.Source = New BitmapImage(New Uri("pack://application:,,,/Images/Common/error32.png"))
            m.Owner = Me
            m.ShowDialog()
        End Try
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        res = True
        Close()
    End Sub

    Private Sub TranslateDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
    End Sub
End Class
