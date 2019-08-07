Imports MySql.Data.MySqlClient
Public Class edit_DEVICE

    Public name_device As String


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'look for an specific part
        Dim found_po As Boolean
        Dim rowindex As Integer
        For Each row As DataGridViewRow In dev_sel.Rows
            If String.Compare(row.Cells.Item("Device_Name").Value.ToString, dev_sel_box.Text) = 0 Then
                rowindex = row.Index
                dev_sel.CurrentCell = dev_sel.Rows(rowindex).Cells(0)
                found_po = True
                Exit For
            End If
        Next
        If found_po = False Then
            MsgBox(" Device not found!")
        End If
    End Sub

    Private Sub edit_DEVICE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------------------------------- LOAD DEVICES ----------------------------------------
        Try

            'Fill first DEVICE table where the user can select the device to be modified
            Dim tabl As New DataTable
            Dim ad As New MySqlDataAdapter("Select DEVICE_Name from devices", Login.Connection)
            ad.Fill(tabl)
            dev_sel.DataSource = tabl

            'Setting Columns size 
            For i = 0 To dev_sel.ColumnCount - 1
                With dev_sel.Columns(i)
                    .Width = 690
                End With
            Next i

            'Fill the all parts table where the user can select the parts to be added to the device member
            Dim tabl2 As New DataTable
            Dim ad2 As New MySqlDataAdapter("Select Part_Name, Part_Description, Part_Status, Part_Type from parts_table", Login.Connection)
            ad2.Fill(tabl2)
            allParts.DataSource = tabl2

            'Setting Columns size 
            For i = 0 To allParts.ColumnCount - 1
                With allParts.Columns(i)
                    .Width = 340
                End With
            Next i

        Catch ex As Exception
        End Try
    End Sub

    Private Sub dev_sel_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dev_sel.CellDoubleClick
        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW


        name_device = dev_sel.CurrentCell.Value.ToString

        Try
            '------------- get AKT values -----------------------
            Dim cmd As New MySqlCommand
            Dim cmd_a As New MySqlCommand
            Dim reader As MySqlDataReader


            cmd.Parameters.AddWithValue("@Device_Name", name_device)
            cmd.CommandText = "SELECT * from devices where DEVICE_Name = @Device_Name"

            'cmd_a.CommandText = "SELECT ACN_Number from ACN where Part_Name = @Device_Name"
            '  cmd_a.Parameters.AddWithValue("@Device_Name", name_device)

            cmd.Connection = Login.Connection
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    deviceName.Text = reader(0).ToString
                    advName.Text = reader(1).ToString
                    descr.Text = reader(2).ToString
                    bulk_c.Text = reader(3).ToString
                    labor_c.Text = reader(4).ToString
                    ' legacy.Text = reader(5).ToString
                End While
            End If

            reader.Close()



            '--------------- populate DEVICE members -----------------
            Call Populate_devices()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'delete device data in parts, akn, vendor, acn and Kits table

        If String.Equals(deviceName.Text, "") = True Then
            MessageBox.Show("Please Fill Device Name and ADV Number")
        Else
            Try
                Dim cmd As New MySqlCommand 'use this mysql command to check if the kit exist
                Dim Create_cmd As New MySqlCommand 'parts command        
                Dim Create_akn As New MySqlCommand  'akn command
                Dim Create_device As New MySqlCommand  'kit command
                Dim Create_adv As New MySqlCommand  'dev command

                cmd.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                cmd.Parameters.AddWithValue("@adv", advName.Text)
                cmd.CommandText = "SELECT DEVICE_Name from devices where Device_Name = @Device_Name and ADV_Number = @adv"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    Dim dlgR As DialogResult
                    dlgR = MessageBox.Show("Are you sure you want to delete this Device?", "Attention!", MessageBoxButtons.YesNo)
                    If dlgR = DialogResult.Yes Then

                        reader.Close()

                        Create_cmd.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                        Create_adv.Parameters.AddWithValue("@adv", advName.Text)
                        Create_device.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                        Create_akn.Parameters.AddWithValue("@Device_Name", deviceName.Text)

                        Create_cmd.CommandText = "DELETE FROM parts_table where Part_Name = @Device_Name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                        Create_adv.CommandText = "DELETE FROM adv where ADV_Number = @adv"
                        Create_adv.Connection = Login.Connection
                        Create_adv.ExecuteNonQuery()
                        Create_device.CommandText = "DELETE FROM devices where DEVICE_Name = @Device_Name"
                        Create_device.Connection = Login.Connection
                        Create_device.ExecuteNonQuery()
                        Create_akn.CommandText = "DELETE FROM akn where Part_Name = @Device_Name"
                        Create_akn.Connection = Login.Connection
                        Create_akn.ExecuteNonQuery()

                        Call edit_KIT.Reload_visualizer()
                        '---------------refresh device datagridview
                        Dim tabl As New DataTable
                        Dim ad As New MySqlDataAdapter("Select DEVICE_Name from devices", Login.Connection)
                        ad.Fill(tabl)
                        dev_sel.DataSource = tabl

                        'Setting Columns size 
                        For i = 0 To dev_sel.ColumnCount - 1
                            With dev_sel.Columns(i)
                                .Width = 690
                            End With
                        Next i
                        MessageBox.Show("Device deleted succesfully")

                    Else
                        reader.Close()

                    End If
                Else
                    MessageBox.Show("We could not find the  Device specified to be removed")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '------------------------ Remove a member of the device -------------------------------------------

        If (current_p.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(current_p.CurrentCell.Value.ToString, "") = False Then

                Try
                    Dim cmd As New MySqlCommand
                    cmd.Parameters.AddWithValue("@Part_Name", current_p.CurrentCell.Value.ToString)
                    cmd.Parameters.AddWithValue("@ada", advname.Text.ToString)
                    cmd.CommandText = "DELETE FROM adv where Part_Name = @Part_Name and  ADV_Number = @ada"
                    cmd.Connection = Login.Connection
                    cmd.ExecuteNonQuery()

                    Dim cmdstr As String : cmdstr = "Select Part_Name, Qty from adv where ADV_Number = @adv"
                    Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
                    cmd_k.Parameters.AddWithValue("@adv", advname.Text)

                    Dim table_s As New DataTable
                    Dim ad As New MySqlDataAdapter(cmd_k)
                    ad.Fill(table_s)
                    current_p.DataSource = table_s

                    'Setting Columns size 
                    For i = 0 To current_p.ColumnCount - 1
                        With current_p.Columns(i)
                            .Width = 280
                        End With
                    Next i

                    MessageBox.Show("Part removed")
                    ' Call edit_KIT.Reload_visualizer()


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '---------------------------------------- CHANGE QTY OF DEVICE MEMBER ----------------------

        If (current_p.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(current_p.CurrentCell.Value.ToString, "") = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please input the new quantity.")
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ update qty ----------------------------------
                        Try
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@Part_Name", current_p.CurrentCell.Value.ToString)
                            Create_cmd.Parameters.AddWithValue("@ada", advname.Text.ToString)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.CommandText = "UPDATE adv  SET  Qty = @qty where Part_Name = @Part_Name and ADV_Number = @ada"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Call edit_KIT.Reload_visualizer()
                            Call Populate_devices()

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try
                        '-----------------------------------------------------
                    Else
                        MsgBox("Please input an integer.")
                    End If
                Catch
                    MsgBox("Please input a whole number.")
                End Try

            End If

        End If
    End Sub

    Sub Populate_devices()

        '--------------- populate DEVICE members -----------------
        Try
            Dim cmdstr As String : cmdstr = "Select Part_Name, Qty from adv where ADV_Number = @adv"
            Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
            cmd_k.Parameters.AddWithValue("@adv", advname.Text)

            Dim table_s As New DataTable
            Dim ad As New MySqlDataAdapter(cmd_k)
            ad.Fill(table_s)
            current_p.DataSource = table_s

            'Setting Columns size 
            For i = 0 To current_p.ColumnCount - 1
                With current_p.Columns(i)
                    .Width = 280
                End With
            Next i
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '--------------- Update DEVICE --------------------
        If String.Equals(deviceName.Text, "") = False Or String.Equals(advname.Text, "") = False Then

            Try
                Dim Create_cmd As New MySqlCommand
                Dim Create_part As New MySqlCommand
                Dim Create_adv As New MySqlCommand
                Dim Create_inv As New MySqlCommand
                Dim Create_mo As New MySqlCommand



                '----------- device table par --------------------
                Create_cmd.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                Create_cmd.Parameters.AddWithValue("@ADV_Number", advname.Text)
                Create_cmd.Parameters.AddWithValue("@Description", descr.Text)
                Create_cmd.Parameters.AddWithValue("@labor", labor_c.Text)
                Create_cmd.Parameters.AddWithValue("@bulk", bulk_c.Text)
                Create_cmd.Parameters.AddWithValue("@legacy", deviceName.Text)

                '------- adv table ------------
                Create_adv.Parameters.AddWithValue("@legacy", deviceName.Text)
                Create_adv.Parameters.AddWithValue("@ADV_Number", advname.Text)

                '------------- part table -------------------------------
                Create_part.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                Create_part.Parameters.AddWithValue("@Description", descr.Text)
                Create_part.Parameters.AddWithValue("@legacy", deviceName.Text)
                Create_part.Parameters.AddWithValue("@oldname", name_device)

                '--------- inventory assembly -----------
                Create_inv.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                Create_inv.Parameters.AddWithValue("@oldname", name_device)

                '-------- update material order ---------
                Create_mo.Parameters.AddWithValue("@Device_Name", deviceName.Text)
                Create_mo.Parameters.AddWithValue("@oldname", name_device)

                ' update device table
                Create_cmd.CommandText = "UPDATE devices  SET Device_Name = @Device_Name, Description = @Description, Bulk_Cost = @bulk, Labor_Cost = @labor, Legacy_ADA_Number = @legacy  where ADV_Number = @ADV_Number"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                '--------------------------------------------------

                'update adv
                Create_adv.CommandText = "UPDATE adv  SET  Legacy_ADA = @legacy  where ADV_Number = @ADV_Number"
                Create_adv.Connection = Login.Connection
                Create_adv.ExecuteNonQuery()
                '-------------------------------------------------
                '--update part
                Create_part.CommandText = "UPDATE parts_table SET Part_Name = @Device_Name, Part_Description = @Description, Legacy_ADA_Number = @legacy where Part_Name = @oldname"
                Create_part.Connection = Login.Connection
                Create_part.ExecuteNonQuery()

                '------- update inventory ---
                Create_inv.CommandText = "UPDATE inventory.inventory_qty SET part_name = @Device_Name where part_name = @oldname"
                Create_inv.Connection = Login.Connection
                Create_inv.ExecuteNonQuery()

                '------- update Material order ---
                Create_mo.CommandText = "UPDATE inventory.Material_orders SET Part_No = @Device_Name where Part_No = @oldname"
                Create_mo.Connection = Login.Connection
                Create_mo.ExecuteNonQuery()



                Call edit_KIT.Reload_visualizer()
                MessageBox.Show("Device updated succesfully")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Else
            MessageBox.Show("Please Fill Device Name and ADV Number")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '-------- Add components to the Device ----------------
        If (allParts.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(allParts.CurrentCell.Value.ToString, "") = False And String.Equals(advname.Text, "") = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString & " this device need")
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ add components to adv table ----------------------------------
                        Try
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@Part_Name", allParts.CurrentCell.Value.ToString)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.Parameters.AddWithValue("@adv", advname.Text)
                            Create_cmd.Parameters.AddWithValue("@legacy", deviceName.Text)
                            Create_cmd.CommandText = "INSERT INTO adv VALUES (@Part_Name, @adv, @qty, @legacy)"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Call Populate_devices()
                            ' MessageBox.Show(allParts.CurrentCell.Value.ToString & " added succesfully")

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try
                        '-----------------------------------------------------
                    Else
                        MsgBox("Please input an integer.")
                    End If
                Catch
                    MsgBox("Please input a whole number.")
                End Try

            End If

        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer
        For Each row As DataGridViewRow In allParts.Rows
            If String.Compare(row.Cells.Item("Part_Name").Value.ToString, TextBox3.Text, True) = 0 Then
                rowindex = row.Index
                allParts.CurrentCell = allParts.Rows(rowindex).Cells(0)
                found_po = True
                Exit For
            End If
        Next
        If found_po = False Then
            MsgBox("Part not found!")
        End If
    End Sub

End Class