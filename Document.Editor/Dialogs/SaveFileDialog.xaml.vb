Imports MS.WindowsAPICodePack.Internal.CoreHelpers
Partial Public Class SaveFileDialog
    Public Res As String = Nothing

    Public Sub SetFileInfo(ByVal name As String, ByVal RTB As RichTextBox)
        Label1.Content = "Do you want to save " + name + "?"
        RichTextBox1.AppendText("")
    End Sub

#Region "Buttons"

    Private Sub YesButton_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles YesButton.Click
        Res = "Yes"
        Close()
    End Sub

    Private Sub NoButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles NoButton.Click
        Res = "No"
        Close()
    End Sub

    Private Sub CancelButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles CancelButton.Click
        Res = Nothing
        Close()
    End Sub

#End Region

    Private Sub SaveFileDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If Me.WindowState = Windows.WindowState.Maximized Then
            My.Settings.SaveDialog_IsMax = True
        Else
            My.Settings.SaveDialog_IsMax = False
        End If
    End Sub

    Private Sub SaveFileDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        Dim fs As IO.FileStream = IO.File.OpenRead(My.Computer.FileSystem.SpecialDirectories.Temp + "\TVPre.xaml")
        Dim tr As New TextRange(RichTextBox1.Document.ContentStart, RichTextBox1.Document.ContentEnd)
        Dim content As FlowDocument = TryCast(Markup.XamlReader.Load(fs), FlowDocument)
        RichTextBox1.Document = content
        fs.Close()
        If My.Settings.SaveDialog_IsMax Then
            Me.WindowState = Windows.WindowState.Maximized
        Else
            Me.WindowState = Windows.WindowState.Normal
        End If
    End Sub
End Class