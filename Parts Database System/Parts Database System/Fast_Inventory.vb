Imports MySql.Data.MySqlClient

Public Class Fast_Inventory
    Private Sub Fast_Inventory_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged

        If Me.Visible = True Then

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@part_name", Me.Text)
                cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        current_qty.Text = reader4(0).ToString
                    End While
                End If

                reader4.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try



        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Try
            Dim qty_in_order As Double : qty_in_order = 0
            Dim date_t As String : date_t = ""


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", Me.Text)
            cmd4.CommandText = "SELECT Qty_in_order, es_date_of_arrival from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    qty_in_order = reader4(0).ToString
                    date_t = reader4(1).ToString
                End While
            End If

            reader4.Close()

            If IsNumeric(current_qty.Text) = True Then

                If qty_in_order > 0 Then

                    Dim result As DialogResult = MessageBox.Show("Are you still going to receive " & qty_in_order & " " & Me.Text & " on " & date_t & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.No) Then

                        '--fulfille mr

                        '--------------- find latest version --------------
                        Dim latest_v As String : latest_v = Proc_Material_R.get_last_revision(Procurement_Overview.bom_open.Text)
                        '---------------------------------------------------
                        Call Procurement_Overview.Update_mr_fullfill(latest_v)   'update latest revision mr
                        Call Procurement_Overview.Update_mr_fullfill(Procurement_Overview.bom_open.Text)    'update current mr which can be the same as above
                        Call Procurement_Overview.update_mb_fullfill(Proc_Material_R.get_last_revision(Procurement_Overview.Text)) 'update latest MB

                        Call Procurement_Overview.Update_current_inv()  'update inventory current qtys

                        Procurement_Overview.go_ahead = False
                        '--- reload entire table except current inv----
                        Procurement_Overview.fullfill_grid.Rows.Clear()
                        Procurement_Overview.go_ahead = True



                        '--- update upcoming part qty
                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@part_name", Me.Text)
                        Create_cmd.Parameters.AddWithValue("@current_qty", If(String.Equals(current_qty.Text, ""), 0, current_qty.Text))
                        Create_cmd.Parameters.AddWithValue("@Qty_in_order", 0)
                        Create_cmd.Parameters.AddWithValue("@es_date_of_arrival", DBNull.Value)

                        Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty, Qty_in_order = @Qty_in_order, es_date_of_arrival = @es_date_of_arrival where part_name = @part_name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()
                        Call Inventory_manage.General_inv_cal()


                        Call Procurement_Overview.Open_recent(Procurement_Overview.bom_open.Text, Procurement_Overview.open_job)
                        Call Procurement_Overview.Color_Module()
                        Me.Visible = False

                    End If
                Else

                    '--------------- find latest version --------------
                    Dim latest_v As String : latest_v = Proc_Material_R.get_last_revision(Procurement_Overview.bom_open.Text)
                    '---------------------------------------------------
                    Call Procurement_Overview.Update_mr_fullfill(latest_v)   'update latest revision mr
                    Call Procurement_Overview.Update_mr_fullfill(Procurement_Overview.bom_open.Text)    'update current mr which can be the same as above
                    Call Procurement_Overview.update_mb_fullfill(Proc_Material_R.get_last_revision(Procurement_Overview.Text)) 'update latest MB

                    Call Procurement_Overview.Update_current_inv()  'update inventory current qtys

                    Procurement_Overview.go_ahead = False
                    '--- reload entire table except current inv----
                    Procurement_Overview.fullfill_grid.Rows.Clear()
                    Procurement_Overview.go_ahead = True



                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@part_name", Me.Text)
                    Create_cmd.Parameters.AddWithValue("@current_qty", If(String.Equals(current_qty.Text, ""), 0, current_qty.Text))

                    Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Call Inventory_manage.General_inv_cal()

                    Call Procurement_Overview.Open_recent(Procurement_Overview.bom_open.Text, Procurement_Overview.open_job)
                    Call Procurement_Overview.Color_Module()
                    Me.Visible = False

                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class