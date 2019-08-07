Imports MySql.Data.MySqlClient

Public Class UR_inventory

    Sub findrows(part As String)

        Dim rowindex As Integer

        For Each row As DataGridViewRow In Inventory_manage.fullfill_grid.Rows
            If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
                If String.Compare(row.Cells.Item("Column10").Value.ToString, part, True) = 0 Then
                    rowindex = row.Index
                    Inventory_manage.fullfill_grid.CurrentCell = Inventory_manage.fullfill_grid.Rows(rowindex).Cells(0)

                    Exit For
                End If
            End If
        Next
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '---------- Update min, max and location
        Try

            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@part_name", part_name.Text)
            Create_cmd.Parameters.AddWithValue("@min_qty", If(String.Equals(Min_qty.Text, ""), 0, Min_qty.Text))
            Create_cmd.Parameters.AddWithValue("@max_qty", If(String.Equals(max_qty.Text, ""), 0, max_qty.Text))
            Create_cmd.Parameters.AddWithValue("@location", Location_t.Text)

            Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET min_qty = @min_qty, max_qty = @max_qty, location = @location where part_name = @part_name"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()

            Call Inventory_manage.refresh_table()


            '---- select last selected part

            Call findrows(part_name.Text)
            MessageBox.Show("Inventory updated")


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '-- update current_inventory qty

        Try
            Dim qty_in_order As Double : qty_in_order = 0
            Dim date_t As String : date_t = ""
            Dim max As Double = max = 0

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part_name.Text)
            cmd4.CommandText = "SELECT Qty_in_order, es_date_of_arrival, max_qty from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    qty_in_order = reader4(0).ToString
                    date_t = reader4(1).ToString
                    max = reader4(2).ToString
                End While
            End If

            reader4.Close()

            If IsNumeric(current_qty.Text) = True Then

                '  If CType(current_qty.Text, Double) <= max Then

                If qty_in_order > 0 Then

                    Dim result As DialogResult = MessageBox.Show("Are you still going to receive " & qty_in_order & " " & part_name.Text & " on " & date_t & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.No) Then

                        '--- update upcoming part qty
                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@part_name", part_name.Text)
                        Create_cmd.Parameters.AddWithValue("@current_qty", If(String.Equals(current_qty.Text, ""), 0, current_qty.Text))
                        Create_cmd.Parameters.AddWithValue("@Qty_in_order", 0)
                        Create_cmd.Parameters.AddWithValue("@es_date_of_arrival", DBNull.Value)

                        Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty, Qty_in_order = @Qty_in_order, es_date_of_arrival = @es_date_of_arrival where part_name = @part_name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()
                        Call Inventory_manage.refresh_table()
                        MessageBox.Show("Inventory updated")

                    End If
                Else

                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@part_name", part_name.Text)
                    Create_cmd.Parameters.AddWithValue("@current_qty", If(String.Equals(current_qty.Text, ""), 0, current_qty.Text))

                    Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Call Inventory_manage.refresh_table()

                    Call findrows(part_name.Text)
                    MessageBox.Show("Inventory updated")

                End If
                ' Else
                'MessageBox.Show("value is greater than the max qty")
                ' End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'add qty to current inv

        Try
            Dim old_inv_qty As Double : old_inv_qty = 0
            Dim date_t As String : date_t = ""
            Dim Qty_in_order As Double : Qty_in_order = 0
            Dim max As Double : max = 0

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part_name.Text)
            cmd4.CommandText = "SELECT current_qty, Qty_in_order, es_date_of_arrival, max_qty from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    old_inv_qty = reader4(0).ToString
                    Qty_in_order = reader4(1).ToString
                    date_t = reader4(2).ToString
                    max = reader4(3).ToString

                End While
            End If

            reader4.Close()

            '-------- update ---------
            Dim add_qty As Double : add_qty = 0

            If IsNumeric(add_box.Text) = True Then
                If add_box.Text >= 0 Then
                    add_qty = add_box.Text
                End If

            End If

            '  If old_inv_qty + add_qty <= max Then


            '--------- reduce qty in order -----------
            If Qty_in_order > 0 And add_qty > 0 Then

                Dim result As DialogResult = MessageBox.Show("Are these some of the " & Qty_in_order & " " & part_name.Text & " you were expecting to received on " & date_t & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                If (result = DialogResult.Yes) Then
                    Qty_in_order = If(Qty_in_order - add_qty < 0, 0, Qty_in_order - add_qty)
                End If

            End If

            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@part_name", part_name.Text)
            Create_cmd.Parameters.AddWithValue("@current_qty", old_inv_qty + add_qty)
            Create_cmd.Parameters.AddWithValue("@Qty_in_order", Qty_in_order)

            If Qty_in_order = 0 Then
                Create_cmd.Parameters.AddWithValue("@es_date_of_arrival", DBNull.Value)
            Else
                Create_cmd.Parameters.AddWithValue("@es_date_of_arrival", date_t)
            End If


            Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty, Qty_in_order = @Qty_in_order, es_date_of_arrival = @es_date_of_arrival where part_name = @part_name"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()

            Call Inventory_manage.refresh_table()
            Call findrows(part_name.Text)

            MessageBox.Show("Inventory updated")
            current_qty.Text = old_inv_qty + add_qty


            ' Else
            '     MessageBox.Show("New current qty is greater than max qty!")
            'End If




        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '-- remove qty
        Try
            Dim old_inv_qty As Double : old_inv_qty = 0

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part_name.Text)
            cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    old_inv_qty = reader4(0).ToString
                End While
            End If

            reader4.Close()

            '-------- update ---------
            Dim rem_qty As Double : rem_qty = 0

            If IsNumeric(remove_box.Text) = True Then
                If remove_box.Text >= 0 Then
                    rem_qty = remove_box.Text
                End If

            End If

            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@part_name", part_name.Text)
            Create_cmd.Parameters.AddWithValue("@current_qty", If(old_inv_qty - rem_qty < 0, 0, old_inv_qty - rem_qty))

            Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()

            Call Inventory_manage.refresh_table()
            MessageBox.Show("Inventory updated")
            current_qty.Text = If(old_inv_qty - rem_qty < 0, 0, old_inv_qty - rem_qty)


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub


End Class