Public Class SymbolPanel
    Public Event Click(symbol As String)

#Region "Currency"

    Private Sub DollarButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles DollarButton.MouseDown
        RaiseEvent Click("$")
    End Sub

    Private Sub CentButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CentButton.MouseDown
        RaiseEvent Click("¢")
    End Sub

    Private Sub PoundButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles PoundButton.MouseDown
        RaiseEvent Click("£")
    End Sub

    Private Sub YenButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles YenButton.MouseDown
        RaiseEvent Click("¥")
    End Sub

    Private Sub EuroButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles EuroButton.MouseDown
        RaiseEvent Click("€")
    End Sub

#End Region

#Region "Misc"

    Private Sub CopyrightButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles CopyrightButton.MouseDown
        RaiseEvent Click("©")
    End Sub

    Private Sub RegisteredTrademarkButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles RegisteredTrademarkButton.MouseDown
        RaiseEvent Click("®")
    End Sub

    Private Sub TrademarkButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TrademarkButton.MouseDown
        RaiseEvent Click("™")
    End Sub

#End Region

    Private Sub SuperscriptOneButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles SuperscriptOneButton.MouseDown
        RaiseEvent Click("¹")
    End Sub

    Private Sub SuperscriptTwoButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles SuperscriptTwoButton.MouseDown
        RaiseEvent Click("²")
    End Sub

    Private Sub SuperscriptThreeButton_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles SuperscriptThreeButton.MouseDown
        RaiseEvent Click("³")
    End Sub

End Class