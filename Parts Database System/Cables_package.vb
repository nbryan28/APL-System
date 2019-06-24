Public Class Cables_package

    Public f_t As Double

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Me.Visible = False
    End Sub


    Private Sub part_assembly_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles feet_table.CellValueChanged

        Dim t_sum As Double : t_sum = 0

        For i = 0 To feet_table.Rows.Count - 1

            t_sum = t_sum + If(IsNumeric(feet_table.Rows(i).Cells(0).Value), feet_table.Rows(i).Cells(0).Value, 0) * If(IsNumeric(feet_table.Rows(i).Cells(1).Value), feet_table.Rows(i).Cells(1).Value, 0)

        Next

        If t_sum > f_t Then
            warning_l.Visible = True
        Else
            warning_l.Visible = False
        End If

    End Sub

    Private Sub Cables_package_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        f_t = 0

        If IsNumeric(feet_qty.Text) = True Then
            f_t = feet_qty.Text
        End If
    End Sub
End Class