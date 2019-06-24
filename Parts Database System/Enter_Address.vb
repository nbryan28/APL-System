Public Class Enter_Address
    Private Sub Enter_Address_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ship_box.Text = BOM_types.shipping_ad
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        BOM_types.shipping_ad = ship_box.Text
        Me.Visible = False
    End Sub

    Private Sub Enter_Address_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            ship_box.Text = ""
        End If
    End Sub
End Class