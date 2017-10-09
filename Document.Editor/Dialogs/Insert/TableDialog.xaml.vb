Public Class TableDialog
    Public Res As String = "Cancel", BackgroundColor As SolidColorBrush = Brushes.Transparent, CellBackgroundColor As SolidColorBrush = Brushes.Transparent, BorderColor As SolidColorBrush = Brushes.Transparent, CellBorderColor As SolidColorBrush = Brushes.Black

    Private Sub UpdatePreview()
        Dim t As New Table, int As Integer = Convert.ToInt32(RowsTextBox.Value), int2 As Integer = Convert.ToInt32(CellsTextBox.Value)
        While Not int = 0
            Dim trg As New TableRowGroup, tr As New TableRow
            While Not int2 = 0
                Dim tc As New TableCell
                tc.Background = CellBackgroundColor
                tc.BorderBrush = CellBorderColor
                tc.BorderThickness = New Thickness(1, 1, 1, 1)
                tr.Cells.Add(tc)
                int2 -= 1
            End While
            int2 = Convert.ToInt32(CellsTextBox.Value)
            trg.Rows.Add(tr)
            t.RowGroups.Add(trg)
            int -= 1
        End While
        t.Background = BackgroundColor
        t.BorderBrush = BorderColor
        t.BorderThickness = New Thickness(1, 1, 1, 1)
        Dim flowdoc As New FlowDocument
        flowdoc.Blocks.Add(t)
        PreviewBox.Document = flowdoc
    End Sub

    Private Sub TableDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        UpdatePreview()
        RowsTextBox.Focus()
    End Sub

    Private Sub RowsTextBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles RowsTextBox.ValueChanged, CellsTextBox.ValueChanged
        If IsLoaded Then
            UpdatePreview()
        End If
    End Sub

    Private Sub BorderColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles BorderColorGallery.SelectedColorChanged
        BorderColor = New SolidColorBrush(BorderColorGallery.SelectedColor)
        UpdatePreview()
    End Sub

    Private Sub CellBorderColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles CellBorderColorGallery.SelectedColorChanged
        CellBorderColor = New SolidColorBrush(CellBorderColorGallery.SelectedColor)
        UpdatePreview()
    End Sub

    Private Sub BackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles BackgroundColorGallery.SelectedColorChanged
        BackgroundColor = New SolidColorBrush(BackgroundColorGallery.SelectedColor)
        UpdatePreview()
    End Sub

    Private Sub CellBackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles CellBackgroundColorGallery.SelectedColorChanged
        CellBackgroundColor = New SolidColorBrush(CellBackgroundColorGallery.SelectedColor)
        UpdatePreview()
    End Sub

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub
End Class