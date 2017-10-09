Public Class FluentTableGrid

    Public Event Click(y As Integer, x As Integer)

#Region "Highlight"

#Region "Row 1"

    Private Sub TableGrid1X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        SetTitle(1, 1)
    End Sub

    Private Sub TableGrid1X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        SetTitle(1, 2)
    End Sub

    Private Sub TableGrid1X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X3.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        SetTitle(1, 3)
    End Sub

    Private Sub TableGrid1X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X3.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X4.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        TableGridItem_MouseEnter(TableGrid1X4)
        SetTitle(1, 4)
    End Sub

    Private Sub TableGrid1X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X4.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        TableGridItem_MouseLeave(TableGrid1X4)
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 5 Then
                Exit For
            End If
        Next
        SetTitle(1, 5)
    End Sub

    Private Sub TableGrid1X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 5 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 6 Then
                Exit For
            End If
        Next
        SetTitle(1, 6)
    End Sub

    Private Sub TableGrid1X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 6 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 7 Then
                Exit For
            End If
        Next
        SetTitle(1, 7)
    End Sub

    Private Sub TableGrid1X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 7 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 8 Then
                Exit For
            End If
        Next
        SetTitle(1, 8)
    End Sub

    Private Sub TableGrid1X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 8 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 9 Then
                Exit For
            End If
        Next
        SetTitle(1, 9)
    End Sub

    Private Sub TableGrid1X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 9 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

    Private Sub TableGrid1X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid1X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 10 Then
                Exit For
            End If
        Next
        SetTitle(1, 10)
    End Sub

    Private Sub TableGrid1X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid1X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 10 Then
                Exit For
            End If
        Next
        SetTitle(Nothing, Nothing)
    End Sub

#End Region

#Region "Row 2"

    Private Sub TableGrid2X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        SetTitle(2, 1)
    End Sub

    Private Sub TableGrid2X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
    End Sub

    Private Sub TableGrid2X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        SetTitle(2, 2)
    End Sub

    Private Sub TableGrid2X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
    End Sub

    Private Sub TableGrid2X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X3.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid2X3)
        SetTitle(2, 3)
    End Sub

    Private Sub TableGrid2X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X3.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid2X3)
    End Sub

    Private Sub TableGrid2X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X4.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        TableGridItem_MouseEnter(TableGrid1X4)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid2X3)
        TableGridItem_MouseEnter(TableGrid2X4)
        SetTitle(2, 4)
    End Sub

    Private Sub TableGrid2X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X4.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        TableGridItem_MouseLeave(TableGrid1X4)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid2X3)
        TableGridItem_MouseLeave(TableGrid2X4)
    End Sub

    Private Sub TableGrid2X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X5.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        TableGridItem_MouseEnter(TableGrid1X4)
        TableGridItem_MouseEnter(TableGrid1X5)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid2X3)
        TableGridItem_MouseEnter(TableGrid2X4)
        TableGridItem_MouseEnter(TableGrid2X5)
        SetTitle(2, 5)
    End Sub

    Private Sub TableGrid2X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X5.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        TableGridItem_MouseLeave(TableGrid1X4)
        TableGridItem_MouseLeave(TableGrid1X5)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid2X3)
        TableGridItem_MouseLeave(TableGrid2X4)
        TableGridItem_MouseLeave(TableGrid2X5)
    End Sub

    Private Sub TableGrid2X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(2, 6)
    End Sub

    Private Sub TableGrid2X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid2X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(2, 7)
    End Sub

    Private Sub TableGrid2X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid2X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(2, 8)
    End Sub

    Private Sub TableGrid2X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid2X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(2, 9)
    End Sub

    Private Sub TableGrid2X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 Then
                count += 1
            ElseIf count = 21 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid2X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid2X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 20 Then
                Exit For
            End If
        Next
        SetTitle(2, 10)
    End Sub

    Private Sub TableGrid2X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid2X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 20 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 3"

    Private Sub TableGrid3X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        SetTitle(3, 1)
    End Sub

    Private Sub TableGrid3X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
    End Sub

    Private Sub TableGrid3X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        SetTitle(3, 2)
    End Sub

    Private Sub TableGrid3X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
    End Sub

    Private Sub TableGrid3X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X3.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid1X3)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid2X3)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid3X3)
        SetTitle(3, 3)
    End Sub

    Private Sub TableGrid3X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X3.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid1X3)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid2X3)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid3X3)
    End Sub

    Private Sub TableGrid3X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 Then
                count += 1
            ElseIf count = 37 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 4)
    End Sub

    Private Sub TableGrid3X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 Then
                count += 1
            ElseIf count = 37 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 Then
                count += 1
            ElseIf count = 36 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 5)
    End Sub

    Private Sub TableGrid3X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 Then
                count += 1
            ElseIf count = 36 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 Then
                count += 1
            ElseIf count = 35 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 6)
    End Sub

    Private Sub TableGrid3X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 Then
                count += 1
            ElseIf count = 35 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 Then
                count += 1
            ElseIf count = 34 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 7)
    End Sub

    Private Sub TableGrid3X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 Then
                count += 1
            ElseIf count = 34 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 Then
                count += 1
            ElseIf count = 33 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 8)
    End Sub

    Private Sub TableGrid3X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 Then
                count += 1
            ElseIf count = 33 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 Then
                count += 1
            ElseIf count = 32 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(3, 9)
    End Sub

    Private Sub TableGrid3X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 Then
                count += 1
            ElseIf count = 32 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid3X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid3X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 30 Then
                Exit For
            End If
        Next
        SetTitle(3, 10)
    End Sub

    Private Sub TableGrid3X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid3X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 30 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 4"

    Private Sub TableGrid4X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid4X1)
        SetTitle(4, 1)
    End Sub

    Private Sub TableGrid4X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid4X1)
    End Sub

    Private Sub TableGrid4X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid4X2)
        SetTitle(4, 2)
    End Sub

    Private Sub TableGrid4X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid4X2)
    End Sub

    Private Sub TableGrid4X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X3.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 Then
                count += 1
            ElseIf count = 55 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 3)
    End Sub

    Private Sub TableGrid4X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X3.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 Then
                count += 1
            ElseIf count = 55 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 Then
                count += 1
            ElseIf count = 53 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 4)
    End Sub

    Private Sub TableGrid4X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 Then
                count += 1
            ElseIf count = 53 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 Then
                count += 1
            ElseIf count = 51 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 5)
    End Sub

    Private Sub TableGrid4X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 Then
                count += 1
            ElseIf count = 51 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 Then
                count += 1
            ElseIf count = 49 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 6)
    End Sub

    Private Sub TableGrid4X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 Then
                count += 1
            ElseIf count = 49 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 Then
                count += 1
            ElseIf count = 47 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 7)
    End Sub

    Private Sub TableGrid4X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 Then
                count += 1
            ElseIf count = 47 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 Then
                count += 1
            ElseIf count = 45 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 8)
    End Sub

    Private Sub TableGrid4X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 Then
                count += 1
            ElseIf count = 45 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 Then
                count += 1
            ElseIf count = 43 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(4, 9)
    End Sub

    Private Sub TableGrid4X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 Then
                count += 1
            ElseIf count = 43 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid4X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid4X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 40 Then
                Exit For
            End If
        Next
        SetTitle(4, 10)
    End Sub

    Private Sub TableGrid4X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid4X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 40 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 5"

    Private Sub TableGrid5X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid5X1)
        SetTitle(5, 1)
    End Sub

    Private Sub TableGrid5X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid5X1)
    End Sub

    Private Sub TableGrid5X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid4X2)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid5X2)
        SetTitle(5, 2)
    End Sub

    Private Sub TableGrid5X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid4X2)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid5X2)
    End Sub

    Private Sub TableGrid5X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X3.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 Then
                count += 1
            ElseIf count = 72 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 3)
    End Sub

    Private Sub TableGrid5X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X3.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 Then
                count += 1
            ElseIf count = 72 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 Then
                count += 1
            ElseIf count = 69 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 4)
    End Sub

    Private Sub TableGrid5X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 Then
                count += 1
            ElseIf count = 69 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 Then
                count += 1
            ElseIf count = 66 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 5)
    End Sub

    Private Sub TableGrid5X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 Then
                count += 1
            ElseIf count = 66 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 Then
                count += 1
            ElseIf count = 63 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 6)
    End Sub

    Private Sub TableGrid5X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 Then
                count += 1
            ElseIf count = 63 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 Then
                count += 1
            ElseIf count = 60 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 7)
    End Sub

    Private Sub TableGrid5X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 Then
                count += 1
            ElseIf count = 60 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 Then
                count += 1
            ElseIf count = 57 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 8)
    End Sub

    Private Sub TableGrid5X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 Then
                count += 1
            ElseIf count = 57 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 Then
                count += 1
            ElseIf count = 54 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(5, 9)
    End Sub

    Private Sub TableGrid5X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 Then
                count += 1
            ElseIf count = 54 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid5X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid5X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 50 Then
                Exit For
            End If
        Next
        SetTitle(5, 10)
    End Sub

    Private Sub TableGrid5X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid5X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 50 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 6"

    Private Sub TableGrid6X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid6X1)
        SetTitle(6, 1)
    End Sub

    Private Sub TableGrid6X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid6X1)
    End Sub

    Private Sub TableGrid6X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid4X2)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid5X2)
        TableGridItem_MouseEnter(TableGrid6X1)
        TableGridItem_MouseEnter(TableGrid6X2)
        SetTitle(6, 2)
    End Sub

    Private Sub TableGrid6X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid4X2)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid5X2)
        TableGridItem_MouseLeave(TableGrid6X1)
        TableGridItem_MouseLeave(TableGrid6X2)
    End Sub

    Private Sub TableGrid6X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X3.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 Then
                count += 1
            ElseIf count = 89 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 3)
    End Sub

    Private Sub TableGrid6X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X3.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 Then
                count += 1
            ElseIf count = 89 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 Then
                count += 1
            ElseIf count = 85 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 4)
    End Sub

    Private Sub TableGrid6X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 Then
                count += 1
            ElseIf count = 85 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 Then
                count += 1
            ElseIf count = 81 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 5)
    End Sub

    Private Sub TableGrid6X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 Then
                count += 1
            ElseIf count = 81 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 Then
                count += 1
            ElseIf count = 77 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 6)
    End Sub

    Private Sub TableGrid6X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 Then
                count += 1
            ElseIf count = 77 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 Then
                count += 1
            ElseIf count = 73 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 7)
    End Sub

    Private Sub TableGrid6X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 Then
                count += 1
            ElseIf count = 73 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 Then
                count += 1
            ElseIf count = 69 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 8)
    End Sub

    Private Sub TableGrid6X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 Then
                count += 1
            ElseIf count = 69 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 Then
                count += 1
            ElseIf count = 65 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(6, 9)
    End Sub

    Private Sub TableGrid6X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 Then
                count += 1
            ElseIf count = 65 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid6X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid6X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 60 Then
                Exit For
            End If
        Next
        SetTitle(6, 10)
    End Sub

    Private Sub TableGrid6X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid6X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 60 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 7"

    Private Sub TableGrid7X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid6X1)
        TableGridItem_MouseEnter(TableGrid7X1)
        SetTitle(7, 1)
    End Sub

    Private Sub TableGrid7X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid6X1)
        TableGridItem_MouseLeave(TableGrid7X1)
    End Sub

    Private Sub TableGrid7X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid4X2)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid5X2)
        TableGridItem_MouseEnter(TableGrid6X1)
        TableGridItem_MouseEnter(TableGrid6X2)
        TableGridItem_MouseEnter(TableGrid7X1)
        TableGridItem_MouseEnter(TableGrid7X2)
        SetTitle(7, 2)
    End Sub

    Private Sub TableGrid7X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid4X2)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid5X2)
        TableGridItem_MouseLeave(TableGrid6X1)
        TableGridItem_MouseLeave(TableGrid6X2)
        TableGridItem_MouseLeave(TableGrid7X1)
        TableGridItem_MouseLeave(TableGrid7X2)
    End Sub

    Private Sub TableGrid7X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X3.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 OrElse count = 99 OrElse count = 101 Then
                count += 1
            ElseIf count = 106 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 3)
    End Sub

    Private Sub TableGrid7X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X3.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 OrElse count = 99 OrElse count = 101 Then
                count += 1
            ElseIf count = 106 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 Then
                count += 1
            ElseIf count = 101 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 4)
    End Sub

    Private Sub TableGrid7X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 Then
                count += 1
            ElseIf count = 101 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 OrElse count = 81 OrElse count = 83 OrElse count = 85 OrElse count = 87 OrElse count = 89 Then
                count += 1
            ElseIf count = 96 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 5)
    End Sub

    Private Sub TableGrid7X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 OrElse count = 81 OrElse count = 83 OrElse count = 85 OrElse count = 87 OrElse count = 89 Then
                count += 1
            ElseIf count = 96 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 OrElse count = 77 OrElse count = 79 OrElse count = 81 OrElse count = 83 Then
                count += 1
            ElseIf count = 91 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 6)
    End Sub

    Private Sub TableGrid7X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 OrElse count = 77 OrElse count = 79 OrElse count = 81 OrElse count = 83 Then
                count += 1
            ElseIf count = 91 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 OrElse count = 73 OrElse count = 75 OrElse count = 77 Then
                count += 1
            ElseIf count = 86 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 7)
    End Sub

    Private Sub TableGrid7X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 OrElse count = 73 OrElse count = 75 OrElse count = 77 Then
                count += 1
            ElseIf count = 86 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 OrElse count = 69 OrElse count = 71 Then
                count += 1
            ElseIf count = 81 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 8)
    End Sub

    Private Sub TableGrid7X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 OrElse count = 69 OrElse count = 71 Then
                count += 1
            ElseIf count = 81 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 OrElse count = 65 Then
                count += 1
            ElseIf count = 76 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(7, 9)
    End Sub

    Private Sub TableGrid7X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 OrElse count = 65 Then
                count += 1
            ElseIf count = 76 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid7X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid7X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 70 Then
                Exit For
            End If
        Next
        SetTitle(7, 10)
    End Sub

    Private Sub TableGrid7X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid7X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 70 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#Region "Row 8"

    Private Sub TableGrid8X1_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X1.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid6X1)
        TableGridItem_MouseEnter(TableGrid7X1)
        TableGridItem_MouseEnter(TableGrid8X1)
        SetTitle(8, 1)
    End Sub

    Private Sub TableGrid8X1_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X1.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid6X1)
        TableGridItem_MouseLeave(TableGrid7X1)
        TableGridItem_MouseLeave(TableGrid8X1)
    End Sub

    Private Sub TableGrid8X2_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X2.MouseEnter
        TableGridItem_MouseEnter(TableGrid1X1)
        TableGridItem_MouseEnter(TableGrid1X2)
        TableGridItem_MouseEnter(TableGrid2X1)
        TableGridItem_MouseEnter(TableGrid2X2)
        TableGridItem_MouseEnter(TableGrid3X1)
        TableGridItem_MouseEnter(TableGrid3X2)
        TableGridItem_MouseEnter(TableGrid4X1)
        TableGridItem_MouseEnter(TableGrid4X2)
        TableGridItem_MouseEnter(TableGrid5X1)
        TableGridItem_MouseEnter(TableGrid5X2)
        TableGridItem_MouseEnter(TableGrid6X1)
        TableGridItem_MouseEnter(TableGrid6X2)
        TableGridItem_MouseEnter(TableGrid7X1)
        TableGridItem_MouseEnter(TableGrid7X2)
        TableGridItem_MouseEnter(TableGrid8X1)
        TableGridItem_MouseEnter(TableGrid8X2)
        SetTitle(8, 2)
    End Sub

    Private Sub TableGrid8X2_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X2.MouseLeave
        TableGridItem_MouseLeave(TableGrid1X1)
        TableGridItem_MouseLeave(TableGrid1X2)
        TableGridItem_MouseLeave(TableGrid2X1)
        TableGridItem_MouseLeave(TableGrid2X2)
        TableGridItem_MouseLeave(TableGrid3X1)
        TableGridItem_MouseLeave(TableGrid3X2)
        TableGridItem_MouseLeave(TableGrid4X1)
        TableGridItem_MouseLeave(TableGrid4X2)
        TableGridItem_MouseLeave(TableGrid5X1)
        TableGridItem_MouseLeave(TableGrid5X2)
        TableGridItem_MouseLeave(TableGrid6X1)
        TableGridItem_MouseLeave(TableGrid6X2)
        TableGridItem_MouseLeave(TableGrid7X1)
        TableGridItem_MouseLeave(TableGrid7X2)
        TableGridItem_MouseLeave(TableGrid8X1)
        TableGridItem_MouseLeave(TableGrid8X2)
    End Sub

    Private Sub TableGrid8X3_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X3.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 OrElse count = 99 OrElse count = 101 OrElse count = 106 OrElse count = 108 OrElse count = 110 OrElse count = 112 OrElse count = 114 OrElse count = 116 OrElse count = 118 Then
                count += 1
            ElseIf count = 123 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 3)
    End Sub

    Private Sub TableGrid8X3_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X3.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 4 OrElse count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 16 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 33 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 46 OrElse count = 48 OrElse count = 50 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 72 OrElse count = 74 OrElse count = 76 OrElse count = 78 OrElse count = 80 OrElse count = 82 OrElse count = 84 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 OrElse count = 99 OrElse count = 101 OrElse count = 106 OrElse count = 108 OrElse count = 110 OrElse count = 112 OrElse count = 114 OrElse count = 116 OrElse count = 118 Then
                count += 1
            ElseIf count = 123 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X4_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X4.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 101 OrElse count = 103 OrElse count = 105 OrElse count = 107 OrElse count = 109 OrElse count = 111 Then
                count += 1
            ElseIf count = 117 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 4)
    End Sub

    Private Sub TableGrid8X4_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X4.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 5 OrElse count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 15 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 31 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 43 OrElse count = 45 OrElse count = 47 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 61 OrElse count = 63 OrElse count = 69 OrElse count = 71 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 79 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 101 OrElse count = 103 OrElse count = 105 OrElse count = 107 OrElse count = 109 OrElse count = 111 Then
                count += 1
            ElseIf count = 117 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X5_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X5.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 OrElse count = 81 OrElse count = 83 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 96 OrElse count = 98 OrElse count = 100 OrElse count = 102 OrElse count = 104 Then
                count += 1
            ElseIf count = 111 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 5)
    End Sub

    Private Sub TableGrid8X5_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X5.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 6 OrElse count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 14 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 29 OrElse count = 36 OrElse count = 38 OrElse count = 40 OrElse count = 42 OrElse count = 44 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 57 OrElse count = 59 OrElse count = 66 OrElse count = 68 OrElse count = 70 OrElse count = 72 OrElse count = 74 OrElse count = 81 OrElse count = 83 OrElse count = 85 OrElse count = 87 OrElse count = 89 OrElse count = 96 OrElse count = 98 OrElse count = 100 OrElse count = 102 OrElse count = 104 Then
                count += 1
            ElseIf count = 111 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X6_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X6.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 OrElse count = 77 OrElse count = 79 OrElse count = 81 OrElse count = 83 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 Then
                count += 1
            ElseIf count = 105 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 6)
    End Sub

    Private Sub TableGrid8X6_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X6.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 7 OrElse count = 9 OrElse count = 11 OrElse count = 13 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 27 OrElse count = 35 OrElse count = 37 OrElse count = 39 OrElse count = 41 OrElse count = 49 OrElse count = 51 OrElse count = 53 OrElse count = 55 OrElse count = 63 OrElse count = 65 OrElse count = 67 OrElse count = 69 OrElse count = 77 OrElse count = 79 OrElse count = 81 OrElse count = 83 OrElse count = 91 OrElse count = 93 OrElse count = 95 OrElse count = 97 Then
                count += 1
            ElseIf count = 105 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X7_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X7.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 86 OrElse count = 88 OrElse count = 90 Then
                count += 1
            ElseIf count = 99 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 7)
    End Sub

    Private Sub TableGrid8X7_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X7.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 8 OrElse count = 10 OrElse count = 12 OrElse count = 21 OrElse count = 23 OrElse count = 25 OrElse count = 34 OrElse count = 36 OrElse count = 38 OrElse count = 47 OrElse count = 49 OrElse count = 51 OrElse count = 60 OrElse count = 62 OrElse count = 64 OrElse count = 73 OrElse count = 75 OrElse count = 77 OrElse count = 86 OrElse count = 88 OrElse count = 90 Then
                count += 1
            ElseIf count = 99 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X8_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X8.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 OrElse count = 69 OrElse count = 71 OrElse count = 81 OrElse count = 83 Then
                count += 1
            ElseIf count = 93 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 8)
    End Sub

    Private Sub TableGrid8X8_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X8.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 9 OrElse count = 11 OrElse count = 21 OrElse count = 23 OrElse count = 33 OrElse count = 35 OrElse count = 45 OrElse count = 47 OrElse count = 57 OrElse count = 59 OrElse count = 69 OrElse count = 71 OrElse count = 81 OrElse count = 83 Then
                count += 1
            ElseIf count = 93 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X9_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X9.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 OrElse count = 65 OrElse count = 76 Then
                count += 1
            ElseIf count = 87 Then
                Exit For
            Else
                TableGridItem_MouseEnter(g)
            End If
        Next
        SetTitle(8, 9)
    End Sub

    Private Sub TableGrid8X9_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X9.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            count += 1
            If count = 10 OrElse count = 21 OrElse count = 32 OrElse count = 43 OrElse count = 54 OrElse count = 65 OrElse count = 76 Then
                count += 1
            ElseIf count = 87 Then
                Exit For
            Else
                TableGridItem_MouseLeave(g)
            End If
        Next
    End Sub

    Private Sub TableGrid8X10_MouseEnter(sender As Object, e As MouseEventArgs) Handles TableGrid8X10.MouseEnter
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseEnter(g)
            count += 1
            If count = 80 Then
                Exit For
            End If
        Next
        SetTitle(8, 10)
    End Sub

    Private Sub TableGrid8X10_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid8X10.MouseLeave
        Dim count As Integer = 0
        For Each g As Grid In TableGrid.Items
            TableGridItem_MouseLeave(g)
            count += 1
            If count = 80 Then
                Exit For
            End If
        Next
    End Sub

#End Region

#End Region

#Region "Click"

#Region "Row 1"

    Private Sub TableGrid1X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X1.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 1)
    End Sub

    Private Sub TableGrid1X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X2.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 2)
    End Sub

    Private Sub TableGrid1X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X3.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 3)
    End Sub

    Private Sub TableGrid1X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X4.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 4)
    End Sub

    Private Sub TableGrid1X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X5.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 5)
    End Sub

    Private Sub TableGrid1X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X6.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 6)
    End Sub

    Private Sub TableGrid1X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X7.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 7)
    End Sub

    Private Sub TableGrid1X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X8.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 8)
    End Sub

    Private Sub TableGrid1X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X9.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 9)
    End Sub

    Private Sub TableGrid1X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid1X10.MouseDown
        e.Handled = True
        RaiseEvent Click(1, 10)
    End Sub

#End Region

#Region "Row 2"

    Private Sub TableGrid2X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X1.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 1)
    End Sub

    Private Sub TableGrid2X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X2.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 2)
    End Sub

    Private Sub TableGrid2X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X3.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 3)
    End Sub

    Private Sub TableGrid2X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X4.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 4)
    End Sub

    Private Sub TableGrid2X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X5.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 5)
    End Sub

    Private Sub TableGrid2X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X6.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 6)
    End Sub

    Private Sub TableGrid2X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X7.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 7)
    End Sub

    Private Sub TableGrid2X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X8.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 8)
    End Sub

    Private Sub TableGrid2X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X9.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 9)
    End Sub

    Private Sub TableGrid2X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid2X10.MouseDown
        e.Handled = True
        RaiseEvent Click(2, 10)
    End Sub

#End Region

#Region "Row 3"

    Private Sub TableGrid3X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X1.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 1)
    End Sub

    Private Sub TableGrid3X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X2.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 2)
    End Sub

    Private Sub TableGrid3X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X3.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 3)
    End Sub

    Private Sub TableGrid3X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X4.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 4)
    End Sub

    Private Sub TableGrid3X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X5.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 5)
    End Sub

    Private Sub TableGrid3X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X6.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 6)
    End Sub

    Private Sub TableGrid3X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X7.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 7)
    End Sub

    Private Sub TableGrid3X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X8.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 8)
    End Sub

    Private Sub TableGrid3X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X9.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 9)
    End Sub

    Private Sub TableGrid3X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid3X10.MouseDown
        e.Handled = True
        RaiseEvent Click(3, 10)
    End Sub

#End Region

#Region "Row 4"

    Private Sub TableGrid4X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X1.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 1)
    End Sub

    Private Sub TableGrid4X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X2.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 2)
    End Sub

    Private Sub TableGrid4X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X3.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 3)
    End Sub

    Private Sub TableGrid4X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X4.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 4)
    End Sub

    Private Sub TableGrid4X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X5.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 5)
    End Sub

    Private Sub TableGrid4X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X6.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 6)
    End Sub

    Private Sub TableGrid4X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X7.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 7)
    End Sub

    Private Sub TableGrid4X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X8.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 8)
    End Sub

    Private Sub TableGrid4X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X9.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 9)
    End Sub

    Private Sub TableGrid4X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid4X10.MouseDown
        e.Handled = True
        RaiseEvent Click(4, 10)
    End Sub

#End Region

#Region "Row 5"

    Private Sub TableGrid5X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X1.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 1)
    End Sub

    Private Sub TableGrid5X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X2.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 2)
    End Sub

    Private Sub TableGrid5X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X3.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 3)
    End Sub

    Private Sub TableGrid5X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X4.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 4)
    End Sub

    Private Sub TableGrid5X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X5.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 5)
    End Sub

    Private Sub TableGrid5X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X6.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 6)
    End Sub

    Private Sub TableGrid5X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X7.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 7)
    End Sub

    Private Sub TableGrid5X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X8.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 8)
    End Sub

    Private Sub TableGrid5X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X9.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 9)
    End Sub

    Private Sub TableGrid5X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid5X10.MouseDown
        e.Handled = True
        RaiseEvent Click(5, 10)
    End Sub

#End Region

#Region "Row 6"

    Private Sub TableGrid6X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X1.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 1)
    End Sub

    Private Sub TableGrid6X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X2.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 2)
    End Sub

    Private Sub TableGrid6X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X3.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 3)
    End Sub

    Private Sub TableGrid6X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X4.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 4)
    End Sub

    Private Sub TableGrid6X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X5.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 5)
    End Sub

    Private Sub TableGrid6X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X6.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 6)
    End Sub

    Private Sub TableGrid6X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X7.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 7)
    End Sub

    Private Sub TableGrid6X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X8.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 8)
    End Sub

    Private Sub TableGrid6X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X9.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 9)
    End Sub

    Private Sub TableGrid6X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid6X10.MouseDown
        e.Handled = True
        RaiseEvent Click(6, 10)
    End Sub

#End Region

#Region "Row 7"

    Private Sub TableGrid7X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X1.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 1)
    End Sub

    Private Sub TableGrid7X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X2.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 2)
    End Sub

    Private Sub TableGrid7X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X3.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 3)
    End Sub

    Private Sub TableGrid7X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X4.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 4)
    End Sub

    Private Sub TableGrid7X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X5.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 5)
    End Sub

    Private Sub TableGrid7X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X6.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 6)
    End Sub

    Private Sub TableGrid7X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X7.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 7)
    End Sub

    Private Sub TableGrid7X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X8.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 8)
    End Sub

    Private Sub TableGrid7X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X9.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 9)
    End Sub

    Private Sub TableGrid7X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid7X10.MouseDown
        e.Handled = True
        RaiseEvent Click(7, 10)
    End Sub

#End Region

#Region "Row 8"

    Private Sub TableGrid8X1_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X1.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 1)
    End Sub

    Private Sub TableGrid8X2_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X2.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 2)
    End Sub

    Private Sub TableGrid8X3_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X3.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 3)
    End Sub

    Private Sub TableGrid8X4_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X4.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 4)
    End Sub

    Private Sub TableGrid8X5_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X5.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 5)
    End Sub

    Private Sub TableGrid8X6_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X6.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 6)
    End Sub

    Private Sub TableGrid8X7_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X7.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 7)
    End Sub

    Private Sub TableGrid8X8_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X8.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 8)
    End Sub

    Private Sub TableGrid8X9_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X9.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 9)
    End Sub

    Private Sub TableGrid8X10_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles TableGrid8X10.MouseDown
        e.Handled = True
        RaiseEvent Click(8, 10)
    End Sub

#End Region

#End Region

    Public Sub TableGridItem_MouseEnter(item As Grid)
        Dim b1 As Border = TryCast(item.Children.Item(0), Border)
        Dim b2 As Border = TryCast(b1.Child, Border)
        b1.BorderBrush = Brushes.Red
        b2.BorderBrush = Brushes.Yellow
    End Sub

    Public Sub TableGridItem_MouseLeave(item As Grid)
        Dim b1 As Border = TryCast(item.Children.Item(0), Border)
        Dim b2 As Border = TryCast(b1.Child, Border)
        b1.BorderBrush = Brushes.Transparent
        b2.BorderBrush = Brushes.Black
    End Sub

    Public Sub SetTitle(x As Integer, y As Integer)
        If x = Nothing OrElse y = Nothing Then
            TableGrid.Header = "Insert Table"
        Else
            TableGrid.Header = y.ToString + "x" + x.ToString + " Table"
        End If
    End Sub

    Private Sub TableGrid_MouseLeave(sender As Object, e As MouseEventArgs) Handles TableGrid.MouseLeave
        SetTitle(Nothing, Nothing)
    End Sub
End Class