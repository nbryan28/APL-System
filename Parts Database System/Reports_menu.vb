Public Class Reports_menu
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label1.Visible = True
        Application.DoEvents()

        Manufacturing_Report.Visible = True
        Label1.Visible = False
        Me.Visible = False
    End Sub
End Class