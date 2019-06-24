Imports MySql.Data.MySqlClient

Public Class UR_part

    Public temp_vendor As String

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    '------------ clear textboxes
    Private Sub UR_part_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        Dim ctrl As Control
        For Each ctrl In Me.GroupBox1.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next

        For Each ctrl In Me.GroupBox2.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next




    End Sub

    Private Sub UR_part_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        type_ur.Text = "All Types"

        '----------------------------------- LOAD PARTS TABLE ----------------------------------------
        Try
            Dim tabl As New DataTable
            Dim ad As New MySqlDataAdapter("Select * from parts_table", Login.Connection)
            ad.Fill(tabl)
            Update_grid.DataSource = tabl

            'Setting Columns size 
            For i = 0 To Update_grid.ColumnCount - 1
                With Update_grid.Columns(i)
                    .Width = 140
                End With
            Next i

        Catch ex As Exception
        End Try


        '----------- Populate type combobox --------
        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT Part_Type from parts_type_table"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    type_ur.Items.Add(readerv(0))
                    type_t.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '----------------Populate Units combobox -----------------------
        Try
            Dim cmd_u As New MySqlCommand
            cmd_u.CommandText = "SELECT distinct Units from parts_table where Units is not null"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    units_t.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Update_grid_DoubleClick(sender As Object, e As EventArgs) Handles Update_grid.DoubleClick
        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW

        Dim part_name_t = Update_grid.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_Name", part_name_t)
            cmd.CommandText = "SELECT * from parts_table where part_name = @Part_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    part_t.Text = reader(0).ToString
                    manu_t.Text = reader(1).ToString
                    desc_t.Text = reader(2).ToString
                    notes_t.Text = reader(3).ToString
                    status_t.Text = reader(4).ToString
                    type_t.Text = reader(5).ToString
                    units_t.Text = reader(6).ToString
                    Min_t.Text = reader(7).ToString
                    legacy_t.Text = reader(8).ToString
                    p_vendor_t.Text = reader(9).ToString
                    mfg_t.Text = reader(10).ToString
                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try



    End Sub

    Private Sub type_ur_SelectedIndexChanged(sender As Object, e As EventArgs) Handles type_ur.SelectedIndexChanged
        '-----------UPDATE PART TABLE --------------------

        If ok_press = True Then
            Try
                Dim selectedItem As Object
                Dim query_t As String
                Dim Type_name As String
                selectedItem = type_ur.SelectedItem
                Type_name = selectedItem.ToString

                If String.Equals(Type_name, "All Types") = True Then
                    query_t = ""
                Else
                    query_t = "and Part_Type = '" & Type_name & "'"
                End If

                Dim ta As New DataTable
                Dim adapt As New MySqlDataAdapter("SELECT * from parts_table where 1=1 " & query_t, Login.Connection)
                adapt.Fill(ta)
                Update_grid.DataSource = ta

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '--------------SEARCHING USER STRING------------------
        Dim found_user As Boolean : found_user = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In Update_grid.Rows
            If String.Compare(row.Cells.Item("Part_Name").Value.ToString, Part_textbox.Text) = 0 Then
                rowindex = row.Index
                Update_grid.CurrentCell = Update_grid.Rows(rowindex).Cells(0)
                found_user = True
                Exit For
            End If
        Next
        If found_user = False Then
            MsgBox("Part not found")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '================================ UPDATE PART ================================

        Dim update_flag As Boolean : update_flag = True
        Dim status As New Collection
        Dim types As New Collection


        'Populate part status
        status.Add("Preferred")
        status.Add("Special Order")
        status.Add("")



        'Populate Part Types
        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT Part_Type from parts_type_table"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    types.Add(readerv(0).ToString)
                End While
            End If

            readerv.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        '---------------------- data validation --------------------------------------------------------------------

        '----------- Part name, status and type cannot be null
        If String.Equals(part_t.Text, "") = True Or String.Equals(status_t.Text, "") = True Or String.Equals(type_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name, Part Status and Part Type boxes")
            update_flag = False

            '---- Check if part type is valid
        ElseIf ImportCSV.IsOntheList(types, type_t.Text) = False Then
            MessageBox.Show("Please select a valid Part Type")
            update_flag = False

            '---- Check if part status is valid
        ElseIf ImportCSV.IsOntheList(status, status_t.Text) = False Then
            MessageBox.Show("Please select a valid Part Status")
            update_flag = False

            '---- Check if Min qty is valid
        ElseIf String.Equals(Min_t.Text, "") = False And Not IsNumeric(Min_t.Text) Then
            MessageBox.Show("Invalid Minimum Quantity")
            update_flag = False

        ElseIf IsNumeric(Min_t.Text) Then
            If CType(Min_t.Text, Integer) <= 0 Then
                MessageBox.Show("Minimum Quantity should be greater than zero")
                update_flag = False
            End If
            '--- Make sure Min_qty is 1 if the textbox is empty
        ElseIf String.Equals(Min_t.Text, "") = True Then
            Min_t.Text = "1"

        End If


        Try
            Dim cmd As New MySqlCommand
            cmd.CommandText = "SELECT Part_name from parts_table where Part_name = """ & part_t.Text & """"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader


            'MAKE SURE THE PART EXIST SO IT CAN BE UPDATED
            If Not reader.HasRows Then
                MessageBox.Show("Part does not exist! ... Update incomplete")
                update_flag = False
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        '--------------------------------------------------------------------------------------------------------------------------------------

        If update_flag = True Then
            '-- Start updating the part ---------

            Try
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                Create_cmd.Parameters.AddWithValue("@Manufacturer", manu_t.Text)
                Create_cmd.Parameters.AddWithValue("@Part_Description", desc_t.Text)
                Create_cmd.Parameters.AddWithValue("@Notes", notes_t.Text)
                Create_cmd.Parameters.AddWithValue("@Part_Status", status_t.Text)
                Create_cmd.Parameters.AddWithValue("@Part_Type", type_t.Text)
                Create_cmd.Parameters.AddWithValue("@Units", units_t.Text)
                Create_cmd.Parameters.AddWithValue("@Min_Order_Qty", Min_t.Text)
                Create_cmd.Parameters.AddWithValue("@Legacy_ADA_Number", legacy_t.Text)
                Create_cmd.Parameters.AddWithValue("@Primary_Vendor", p_vendor_t.Text)
                Create_cmd.Parameters.AddWithValue("@MFG_type", mfg_t.Text)

                '   Create_cmd.CommandText = "UPDATE parts_table  SET  Manufacturer = """ & manu_t.Text & """, Part_Description = """ & desc_t.Text & """, Notes = """ & notes_t.Text & """, Part_Status = """ & status_t.Text & """, Part_Type = """ & type_t.Text & """, Units = """ & units_t.Text & """, Min_Order_Qty = """ & Min_t.Text & """, Legacy_ADA_Number = """ & legacy_t.Text & """, Primary_Vendor = """ & p_vendor_t.Text & """ where Part_Name = """ & part_t.Text & """"
                Create_cmd.CommandText = "UPDATE parts_table  SET  Manufacturer = @Manufacturer, Part_Description = @Part_Description, Notes =  @Notes, Part_Status = @Part_Status, Part_Type = @Part_Type, Units =  @Units, Min_Order_Qty = @Min_Order_Qty, Legacy_ADA_Number = @Legacy_ADA_Number, Primary_Vendor = @Primary_Vendor, MFG_type = @MFG_type where Part_Name = @Part_Name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

                '----update inventory_qty table
                Dim Create_cmd2 As New MySqlCommand
                Create_cmd2.Parameters.AddWithValue("@part_name", part_t.Text)
                Create_cmd2.Parameters.AddWithValue("@manufacturer", manu_t.Text)
                Create_cmd2.Parameters.AddWithValue("@description", desc_t.Text)
                Create_cmd2.Parameters.AddWithValue("@MFG_Type", mfg_t.Text)
                Create_cmd2.Parameters.AddWithValue("@units", units_t.Text)

                Create_cmd2.CommandText = "UPDATE inventory.inventory_qty  SET  manufacturer = @manufacturer, description = @description, units =  @units, MFG_type = @MFG_type where part_name = @part_name"
                Create_cmd2.Connection = Login.Connection
                Create_cmd2.ExecuteNonQuery()
                '---------------------------------------


                Call Refresh_tables(Login.Connection)
                MessageBox.Show("Part updated succesfully")
                Call clear_boxes()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub


    Sub Refresh_tables(c As MySqlConnection)

        '---------- REFRESH DATA IN GRID ----------------------
        Try
            'REFRESH PARTS TABLE 
            Dim tabl2 As New DataTable
            Dim ad2 As New MySqlDataAdapter("Select * from parts_table", c)
            ad2.Fill(tabl2)
            Update_grid.DataSource = tabl2

            '----type filter---

            If ok_press = True Then
                Try
                    Dim selectedItem As String : selectedItem = "xxxxxxx"
                    Dim query_t As String
                    Dim Type_name As String : Type_name = "xxxxxx"
                    selectedItem = type_ur.SelectedItem
                    Type_name = selectedItem.ToString

                    If Not type_ur.SelectedItem Is Nothing Then
                        If String.Equals(Type_name, "All Types") = True Then
                            query_t = ""
                        Else
                            query_t = "and Part_Type = '" & Type_name & "'"
                        End If
                    Else
                        query_t = ""
                    End If

                    Dim ta As New DataTable
                    Dim adapt As New MySqlDataAdapter("SELECT * from parts_table where 1=1 " & query_t, Login.Connection)
                    adapt.Fill(ta)
                    Update_grid.DataSource = ta

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If

            '--------- search

            Dim rowindex As Integer

            For Each row As DataGridViewRow In Update_grid.Rows
                If String.Compare(row.Cells.Item("Part_Name").Value.ToString, Part_textbox.Text) = 0 Then
                    rowindex = row.Index
                    Update_grid.CurrentCell = Update_grid.Rows(rowindex).Cells(0)
                    Exit For
                End If
            Next

        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '-----------DELETE PART --------------------

        If String.Equals(part_t.Text, "") = True Or String.Equals(status_t.Text, "") = True Or String.Equals(type_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name, Part Status and Part Type")
        Else
            Try
                '--check if part current_qty > 0
                Dim current_q As Boolean : current_q = False

                Dim cmdq As New MySqlCommand
                cmdq.Parameters.AddWithValue("@part_name", part_t.Text)
                cmdq.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmdq.Connection = Login.Connection
                Dim readerq As MySqlDataReader
                readerq = cmdq.ExecuteReader

                If readerq.HasRows Then
                    While readerq.Read
                        If readerq(0) > 0 Then
                            current_q = True
                        End If
                    End While
                End If

                readerq.Close()
                '-----------------------------------------
                If current_q = False Then

                    Dim cmd As New MySqlCommand
                    Dim Create_cmd As New MySqlCommand
                    Dim Create_vd As New MySqlCommand
                    Dim Create_acn As New MySqlCommand
                    cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                    cmd.CommandText = "SELECT Part_Name from parts_table where Part_Name = @Part_Name"
                    cmd.Connection = Login.Connection
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader

                    If reader.HasRows Then

                        Dim dlgR As DialogResult
                        dlgR = MessageBox.Show("Are you sure you want to delete this Part?", "Attention!", MessageBoxButtons.YesNo)
                        If dlgR = DialogResult.Yes Then

                            reader.Close()

                            Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                            Create_vd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                            Create_acn.Parameters.AddWithValue("@Part_Name", part_t.Text)
                            Create_cmd.CommandText = "DELETE FROM parts_table where Part_Name = @Part_Name"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Create_vd.CommandText = "DELETE FROM vendors_table where Part_Name = @Part_Name"
                            Create_vd.Connection = Login.Connection
                            Create_vd.ExecuteNonQuery()
                            Create_acn.CommandText = "DELETE FROM adv where Part_Name = @Part_Name"
                            Create_acn.Connection = Login.Connection
                            Create_acn.ExecuteNonQuery()

                            '--delete from inventort_qty
                            Dim check_cmd2 As New MySqlCommand
                            check_cmd2.Parameters.AddWithValue("@part_name", part_t.Text)
                            check_cmd2.CommandText = "delete from inventory.inventory_qty where part_name = @part_name"
                            check_cmd2.Connection = Login.Connection
                            check_cmd2.ExecuteNonQuery()
                            '-------------------------


                            MessageBox.Show("Record deleted succesfully")
                            Call Refresh_tables(Login.Connection)
                            Call clear_boxes()
                        Else
                            reader.Close()

                        End If
                    Else
                        MessageBox.Show("We could not find the Part specified to be removed")
                        reader.Close()
                    End If

                Else
                    MessageBox.Show("There are some items on the shelf! Please go to General Inventory and set the current qty to zero before deleting this Part")
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        '----------------------------------- ENTER A NEW PART FROM AN EXISTING PART------------------------------------------

        Dim min_flag As Boolean : min_flag = True


        '---------MAKE SURE TEXTBOXES ARE FILL UP (PART, STATUS AND TYPE) ------------------------
        If String.Equals(part_t.Text, "") = True Or String.Equals(status_t.Text, "") = True Or String.Equals(type_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name, Part Status and Part Type fields")
        Else

            '----------- MAKE SURE THE PART DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                cmd.CommandText = "SELECT Part_Name from parts_table where Part_Name = @Part_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If String.Equals(Min_t.Text, "") = False And IsNumeric(Min_t.Text) = False Then
                    min_flag = False
                    MessageBox.Show("Invalid Minimum Order Quantity")
                    reader.Close()
                End If

                If min_flag = True Then

                    If Not reader.HasRows Then

                        'IF IT DOES NOT EXIST THEN CREATE IT
                        reader.Close()

                        'check for empty and invalid min qty
                        If String.Equals(Min_t.Text, "") Then
                            Min_t.Text = 1
                        ElseIf CType(Min_t.Text, Integer) <= 0 Then
                            Min_t.Text = 1
                        End If

                        Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Manufacturer", manu_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Description", desc_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Notes", notes_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Status", status_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Type", type_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Units", units_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Min_Order_Qty", Min_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Legacy_ADA_Number", legacy_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Primary_Vendor", p_vendor_t.Text)
                        Create_cmd.Parameters.AddWithValue("@MFG_type", mfg_t.Text)

                        Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Legacy_ADA_Number, Primary_Vendor, MFG_type) VALUES (@Part_Name, @Manufacturer, @Part_Description, @Notes, @Part_Status, @Part_Type, @Units, @Min_Order_Qty, @Legacy_ADA_Number, @Primary_Vendor, @MFG_type)"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                        '--- insert into inventory
                        Dim Create_cmdi As New MySqlCommand
                        Create_cmdi.Parameters.AddWithValue("@part_name", part_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@description", desc_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@manufacturer", manu_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@MFG_type", mfg_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@units", units_t.Text)

                        Create_cmdi.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, units, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, @units,0,0,0,0)"
                        Create_cmdi.Connection = Login.Connection
                        Create_cmdi.ExecuteNonQuery()

                        '---------------------------

                        Call Refresh_tables(Login.Connection)
                        MessageBox.Show("Part " & part_t.Text & " added succesfully!")
                        Call clear_boxes()

                    Else
                        MessageBox.Show("Part " & part_t.Text & " already exist!")
                        'DO NOT FORGET TO CLOSE THE READER
                        reader.Close()
                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If

    End Sub

    Sub clear_boxes()
        '-- clear textboxes

        part_t.Text = ""
        Min_t.Text = ""
        manu_t.Text = ""
        p_vendor_t.Text = ""
        legacy_t.Text = ""
        desc_t.Text = ""
        notes_t.Text = ""

    End Sub


End Class