Public Class Spec_inv
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If part_t.Text Is Nothing = False And String.Equals(part_t.Text, "") = False Then

            Call Inventory_manage.add_specific(part_t.Text, desc_t.Text, mfg_t.Text)
            Call Inventory_manage.refresh_table()
            MessageBox.Show("Inventory updated")
            Me.Close()
        Else
            MessageBox.Show("Please enter part name")
        End If
    End Sub
End Class