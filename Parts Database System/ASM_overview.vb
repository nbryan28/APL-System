Imports MySql.Data.MySqlClient

Public Class ASM_overview
    Private Sub ASM_overview_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", Procurement_Overview.Text)
            cmd4.CommandText = "SELECT ASM, qty_needed from Tracking_Reports.ASM_BOM where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    asm_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString})
                End While

                Me.Text = Procurement_Overview.Text

            Else
                MessageBox.Show("Not a Assembly BOM")
            End If

            reader4.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class