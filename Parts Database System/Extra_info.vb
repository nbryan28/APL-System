Imports MySql.Data.MySqlClient

Public Class Extra_info


    Private Sub Extra_info_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '-- when visibility change load numbers

        Dim part_name As String : part_name = Me.Text

        Dim current_inv As String : current_inv = 0
        Dim min_inv As String : min_inv = 0
        Dim max_inv As String : max_inv = 0
        Dim transit_inv As String : transit_inv = 0
        Dim demand_inv As String : demand_inv = 0


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part_name)
            cmd4.CommandText = "SELECT current_qty, min_qty, max_qty, Qty_in_order from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    current_inv = reader4(0).ToString
                    min_inv = reader4(1).ToString
                    max_inv = reader4(2).ToString
                    transit_inv = reader4(3).ToString
                End While
            End If

            reader4.Close()

            demand_inv = Inventory_manage.cal_demand(part_name)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        Current_l.Text = current_inv
        min_l.Text = min_inv
        max_l.Text = max_inv
        qti_l.Text = transit_inv
        demand_l.Text = demand_inv
    End Sub


End Class