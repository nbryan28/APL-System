Public Class Panel_editor
    Private Sub Panel_editor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        p_grid.Rows.Clear()

        If myQuote.Panel_grid.Columns.Count > 1 Then

            For i = 1 To myQuote.Panel_grid.Columns.Count - 1

                p_grid.Rows.Add(New String() {myQuote.Panel_grid.Columns(i).HeaderText, myQuote.pick_panel(myQuote.Panel_grid.Columns(i).HeaderText), myQuote.real_panels(myQuote.Panel_grid.Columns(i).HeaderText), myQuote.my_qtys_panels(myQuote.Panel_grid.Columns(i).HeaderText), myQuote.real_panels_30(myQuote.Panel_grid.Columns(i).HeaderText), myQuote.real_panels_30(myQuote.Panel_grid.Columns(i).HeaderText)})
                '  p_grid.Rows(i - 1).Cells(1).Value = "All 24' Panels"

            Next

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'add data to myquote.dimen_table
        myQuote.dimen_table.Rows.Clear()

        For i = 0 To p_grid.Rows.Count - 1
            myQuote.dimen_table.Rows.Add(p_grid.Rows(i).Cells(0).Value, p_grid.Rows(i).Cells(1).Value, p_grid.Rows(i).Cells(2).Value, p_grid.Rows(i).Cells(3).Value, p_grid.Rows(i).Cells(4).Value, p_grid.Rows(i).Cells(5).Value)
        Next

        myQuote.switch_b = False
        Call myQuote.General_calculation()
        Me.Close()
    End Sub
End Class