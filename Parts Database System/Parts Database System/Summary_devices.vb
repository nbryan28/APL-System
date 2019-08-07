Public Class Summary_devices
    Private Sub Summary_devices_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim total As Double : total = 0

        'summary all inputs qty

        For i = 0 To myQuote.Panel_grid.Rows.Count - 1
            total = 0
            For j = 1 To myQuote.Panel_grid.Columns.Count - 1
                total = total + If(IsNumeric(myQuote.Panel_grid.Rows(i).Cells(j).Value) = True, myQuote.Panel_grid.Rows(i).Cells(j).Value, 0)
            Next

            If total > 0 Then
                setup_grid.Rows.Add(New String() {myQuote.Panel_grid.Rows(i).Cells(0).Value, total})
            End If
        Next


        For i = 0 To myQuote.PLC_grid.Rows.Count - 1
            total = 0
            For j = 1 To myQuote.PLC_grid.Columns.Count - 1
                total = total + If(IsNumeric(myQuote.PLC_grid.Rows(i).Cells(j).Value) = True, myQuote.PLC_grid.Rows(i).Cells(j).Value, 0)
            Next

            If total > 0 Then
                setup_grid.Rows.Add(New String() {myQuote.PLC_grid.Rows(i).Cells(0).Value, total})
            End If
        Next


        For i = 0 To myQuote.Control_grid.Rows.Count - 1
            total = 0
            For j = 1 To myQuote.Control_grid.Columns.Count - 1
                total = total + If(IsNumeric(myQuote.Control_grid.Rows(i).Cells(j).Value) = True, myQuote.Control_grid.Rows(i).Cells(j).Value, 0)
            Next

            If total > 0 Then
                setup_grid.Rows.Add(New String() {myQuote.Control_grid.Rows(i).Cells(0).Value, total})
            End If
        Next


        For i = 0 To myQuote.Field_grid.Rows.Count - 1
            total = 0
            For j = 1 To myQuote.Field_grid.Columns.Count - 1
                total = total + If(IsNumeric(myQuote.Field_grid.Rows(i).Cells(j).Value) = True, myQuote.Field_grid.Rows(i).Cells(j).Value, 0)
            Next

            If total > 0 Then
                setup_grid.Rows.Add(New String() {myQuote.Field_grid.Rows(i).Cells(0).Value, total})
            End If
        Next

    End Sub
End Class