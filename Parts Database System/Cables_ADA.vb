Public Class Cables_ADA
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

        If alt_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(alt_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If

    End Sub
End Class