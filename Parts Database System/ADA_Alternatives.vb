Public Class ADA_Alternatives
    Private Sub alt_grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles alt_grid.CellContentClick

        'When double clicked remove the selected row and replace it with a new with the alternative ADA part

        'ADA_Setup.PR_grid.Rows.RemoveAt(ADA_Setup.row_in)
        'ADA_Setup.PR_grid.Rows.Insert(ADA_Setup.row_in, 1)
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(0).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(0).Value 'ADA Number                        
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(1).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(1).Value  'part name
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(2).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(2).Value   'Part description
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(3).Value = ""  'Notes
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(4).Value = ""  'Type
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(5).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(3).Value  'vendor   
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(6).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(4).Value  'price
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(7).Value = ADA_Setup.myQTY 'qty
        'ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(8).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(4).Value * ADA_Setup.myQTY   'subtotal

        'Me.Visible = False

        'ADA_Setup.PR_grid.CurrentCell = ADA_Setup.PR_grid.Rows(ADA_Setup.row_in).Cells(0)
        'ADA_Setup.PR_grid.CurrentRow.Selected = True

        'MessageBox.Show("Part replaced.. Alternative Part " & alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(1).Value & " added Succesfully")


        Purchase_Request.PR_grid.Rows.RemoveAt(Purchase_Request.row_in)
        Purchase_Request.PR_grid.Rows.Insert(Purchase_Request.row_in, 1)
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(0).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(1).Value  'part name
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(1).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(1).Value  'part desc
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(2).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(0).Value   'ada
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(4).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(3).Value   'vendor
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(5).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(4).Value   'price
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(6).Value = Purchase_Request.myQTY  'qty
        Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(7).Value = alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(4).Value * Purchase_Request.myQTY 'subtotal


        Me.Visible = False

        Purchase_Request.PR_grid.CurrentCell = Purchase_Request.PR_grid.Rows(Purchase_Request.row_in).Cells(0)
        Purchase_Request.PR_grid.CurrentRow.Selected = True

        MessageBox.Show("Part replaced.. Alternative Part " & alt_grid.Rows(alt_grid.CurrentRow.Index).Cells(1).Value & " added Succesfully")

    End Sub
End Class