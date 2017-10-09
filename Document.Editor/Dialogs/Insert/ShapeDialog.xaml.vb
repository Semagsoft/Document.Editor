Public Class ShapeDialog
    Public Res As String = "Cancel"
    Public Shape As Shapes.Shape = Nothing

    Private Sub OKButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles OKButton.Click
        Res = "OK"
        Close()
    End Sub

    Private Sub ShapeDialog_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If My.Computer.Info.OSVersion >= "6.0" Then
            If My.Settings.Options_EnableGlass Then
                AppHelper.ExtendGlassFrame(Me, New Thickness(-1, -1, -1, -1))
            End If
        End If
        TypeComboBox.Items.Add("Circle")
        TypeComboBox.Items.Add("Square")
        TypeComboBox.SelectedIndex = 0
        If TypeComboBox.SelectedIndex = 0 Then
            Shape = New Shapes.Ellipse
            Shape.Height = 32
            Shape.Width = 32
            Dim int2 As Integer = Convert.ToInt32(BorderSizeTextBox.Value)
            Shape.StrokeThickness = int2
            Shape.Stroke = Brushes.Black
        ElseIf TypeComboBox.SelectedIndex = 1 Then
            Shape = New Shapes.Rectangle
            Shape.Height = 32
            Shape.Width = 32
            Dim int2 As Integer = Convert.ToInt32(BorderSizeTextBox.Value)
            Shape.StrokeThickness = int2
            Shape.Stroke = Brushes.Black
        End If
        ScrollViewer1.Content = Shape
    End Sub

    Private Sub TypeComboBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles TypeComboBox.SelectionChanged
        If TypeComboBox.SelectedIndex = 0 Then
            Shape = New Shapes.Ellipse
            Dim int As Integer = Convert.ToInt32(SizeTextBox.Value)
            Shape.Height = int
            Shape.Width = int
            Dim int2 As Integer = Convert.ToInt32(BorderSizeTextBox.Value)
            Shape.StrokeThickness = int2
            Shape.Stroke = Brushes.Black
        ElseIf TypeComboBox.SelectedIndex = 1 Then
            Shape = New Shapes.Rectangle
            Dim int As Integer = Convert.ToInt32(SizeTextBox.Value)
            Shape.Height = int
            Shape.Width = int
            Dim int2 As Integer = Convert.ToInt32(BorderSizeTextBox.Value)
            Shape.StrokeThickness = int2
            Shape.Stroke = Brushes.Black
        End If
        ScrollViewer1.Content = Shape
    End Sub

    Private Sub SizeTextBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles SizeTextBox.ValueChanged
        Try
            Dim int As Integer = Convert.ToInt32(SizeTextBox.Value)
            Shape.Height = int
            Shape.Width = int
        Catch ex As Exception
            SizeTextBox.Value = 32
        End Try
    End Sub

    Private Sub BorderSizeTextBox_ValueChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Double)) Handles BorderSizeTextBox.ValueChanged
        Try
            Dim int As Integer = Convert.ToInt32(BorderSizeTextBox.Value)
            Shape.StrokeThickness = int
        Catch ex As Exception
            BorderSizeTextBox.Value = 4
        End Try
    End Sub
End Class