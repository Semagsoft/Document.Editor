Imports System.Windows.Controls.DataVisualization
Public Class ChartDialog
    Public Res As String = "Cancel"
    Private Series As Charting.ColumnSeries, Items As New Collection

#Region "Loaded"

    Private Sub ChartDialog_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        Items.Add(New KeyValuePair(Of String, Integer)("Item1", 1))
        Items.Add(New KeyValuePair(Of String, Integer)("Item2", 2))
        Items.Add(New KeyValuePair(Of String, Integer)("Item3", 3))
        LoadColumnData(PreviewChart.Series(0))
        LoadPieData(PreviewChart.Series(1))
        PreviewChart.Series.Remove(PieSeries)
        ItemsListBox.ItemsSource = Items
    End Sub

    Private Sub ColumnSeries_Loaded(sender As Object, e As RoutedEventArgs) Handles ColumnSeries.Loaded
        LoadColumnData(ColumnSeries)
    End Sub

    Private Sub PieSeries_Loaded(sender As Object, e As RoutedEventArgs) Handles PieSeries.Loaded
        LoadPieData(PieSeries)
    End Sub

#End Region

#Region "Chart Editor"

    Private Sub LoadColumnData(series As Charting.ISeries)
        DirectCast(series, Charting.ColumnSeries).ItemsSource = Items
    End Sub

    Private Sub LoadPieData(series As Charting.ISeries)
        DirectCast(series, Charting.PieSeries).ItemsSource = Items
    End Sub

    Private Sub UpdatePreview()
        If ChartTypeComboBox.SelectedIndex = 0 Then
            PreviewChart.Series.Remove(PieSeries)
            PreviewChart.Series.Add(ColumnSeries)
            LoadColumnData(ColumnSeries)
        ElseIf ChartTypeComboBox.SelectedIndex = 1 Then
            PreviewChart.Series.Remove(ColumnSeries)
            PreviewChart.Series.Add(PieSeries)
            LoadPieData(PieSeries)
        End If
    End Sub

#Region "Chart"

    Private Sub ChartTypeComboBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ChartTypeComboBox.SelectionChanged
        If IsLoaded Then
            UpdatePreview()
        End If
    End Sub

    Private Sub ChartTitleTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles ChartTitleTextBox.TextChanged
        If IsLoaded Then
            PreviewChart.Title = ChartTitleTextBox.Text
        End If
    End Sub

    Private Sub ForegroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles ForegroundColorGallery.SelectedColorChanged
        If IsLoaded Then
            Dim co As New SolidColorBrush
            co.Color = ForegroundColorGallery.SelectedColor
            PreviewChart.Foreground = co
        End If
    End Sub

    Private Sub BackgroundColorGallery_SelectedColorChanged(sender As Object, e As RoutedEventArgs) Handles BackgroundColorGallery.SelectedColorChanged
        If IsLoaded Then
            Dim co As New SolidColorBrush
            co.Color = BackgroundColorGallery.SelectedColor
            PreviewChart.Background = co
        End If
    End Sub

    Private Sub ChartHight_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ChartHight.ValueChanged
        If IsLoaded Then
            PreviewChart.Height = ChartHight.Value
        End If
    End Sub

    Private Sub ChartWidth_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ChartWidth.ValueChanged
        If IsLoaded Then
            PreviewChart.Width = ChartWidth.Value
        End If
    End Sub

#End Region

#Region "Series"

    Private Sub SeriesTitleTextBox_TextChanged(sender As Object, e As TextChangedEventArgs) Handles SeriesTitleTextBox.TextChanged
        If IsLoaded Then
            Dim s As Charting.ISeries = PreviewChart.Series.Item(0)
            If s.GetType Is GetType(Charting.ColumnSeries) Then
                Dim series As Charting.ColumnSeries = s
                series.Title = SeriesTitleTextBox.Text
            ElseIf s.GetType Is GetType(Charting.PieSeries) Then
                Dim series As Charting.PieSeries = s
                series.Title = SeriesTitleTextBox.Text
            End If
        End If
    End Sub

#End Region

#Region "Items"

    Private Sub ItemsListBox_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ItemsListBox.SelectionChanged
        If ItemsListBox.SelectedItem IsNot Nothing Then
            RemoveItemButton.IsEnabled = True
            ItemTitleTextBox.Visibility = Windows.Visibility.Visible
            ItemValueBox.Visibility = Windows.Visibility.Visible
            Dim i As KeyValuePair(Of String, Integer) = Nothing
            i = Items.Item(ItemsListBox.SelectedIndex + 1)
            ItemTitleTextBox.Text = i.Key
            IsEditing = False
            ItemValueBox.Value = i.Value
            IsEditing = True
        Else
            RemoveItemButton.IsEnabled = False
            ItemTitleTextBox.Visibility = Windows.Visibility.Collapsed
            ItemValueBox.Visibility = Windows.Visibility.Collapsed
        End If
    End Sub

    Private Sub AddItemButton_Click(sender As Object, e As RoutedEventArgs) Handles AddItemButton.Click
        Items.Add(New KeyValuePair(Of String, Integer)("Item" + Convert.ToString(Items.Count + 1), 1))
        ItemsListBox.ItemsSource = Nothing
        ItemsListBox.ItemsSource = Items
        ColumnSeries.ItemsSource = Nothing
        LoadColumnData(ColumnSeries)
        PieSeries.ItemsSource = Nothing
        LoadPieData(PieSeries)
    End Sub

    Private Sub RemoveItemButton_Click(sender As Object, e As RoutedEventArgs) Handles RemoveItemButton.Click
        Items.Remove(ItemsListBox.SelectedIndex + 1)
        ItemsListBox.ItemsSource = Nothing
        ItemsListBox.ItemsSource = Items
        ColumnSeries.ItemsSource = Nothing
        LoadColumnData(ColumnSeries)
        PieSeries.ItemsSource = Nothing
        LoadPieData(PieSeries)
    End Sub

    Private Sub ItemTitleTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles ItemTitleTextBox.KeyDown
        If e.Key = Key.Enter Then
            Dim i As KeyValuePair(Of String, Integer) = Items.Item(ItemsListBox.SelectedIndex + 1)
            Dim newi = New KeyValuePair(Of String, Integer)(ItemTitleTextBox.Text, i.Value)
            Dim c As New Collection
            For Each it As KeyValuePair(Of String, Integer) In Items
                If it.Key Is i.Key Then
                    c.Add(New KeyValuePair(Of String, Integer)(newi.Key, newi.Value))
                Else
                    c.Add(New KeyValuePair(Of String, Integer)(it.Key, it.Value))
                End If
            Next
            Items = c
            ItemsListBox.ItemsSource = Nothing
            ItemsListBox.ItemsSource = Items
            ColumnSeries.ItemsSource = Nothing
            LoadColumnData(ColumnSeries)
            PieSeries.ItemsSource = Nothing
            LoadPieData(PieSeries)
            e.Handled = True
        End If
    End Sub

    Private IsEditing As Boolean = True
    Private Sub ItemValueBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles ItemValueBox.ValueChanged
        If IsLoaded AndAlso IsEditing Then
            Dim i As KeyValuePair(Of String, Integer) = Items.Item(ItemsListBox.SelectedIndex + 1)
            Dim c As New Collection
            For Each it As KeyValuePair(Of String, Integer) In Items
                If it.Key Is i.Key AndAlso it.Value = i.Value Then
                    Dim int As Integer = ItemValueBox.Value
                    c.Add(New KeyValuePair(Of String, Integer)(i.Key, int))
                Else
                    c.Add(New KeyValuePair(Of String, Integer)(it.Key, it.Value))
                End If
            Next
            Items = c
            ItemsListBox.ItemsSource = Nothing
            ItemsListBox.ItemsSource = Items
            ColumnSeries.ItemsSource = Nothing
            LoadColumnData(ColumnSeries)
            PieSeries.ItemsSource = Nothing
            LoadPieData(PieSeries)
        End If
    End Sub

#End Region

#End Region

    Private Sub OKButton_Click(sender As Object, e As RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub
End Class