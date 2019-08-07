Imports MySql.Data.MySqlClient

Public Class Feature_manager

    Public temp_Feature_code As String
    Public temp_description As String
    Public temp_solution As String
    Public temp_sol_desc As String
    Public temp_type As String
    Public temp_vfd_type As String
    Public temp_specific_Type As String
    Public temp_show As String
    Public temp_sol_def As String
    Public temp_assembly As Boolean


    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        TabControl1.SelectedTab = tab1_manu
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TabControl1.SelectedTab = Panel_alloc
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TabControl1.SelectedTab = BOM_alloc
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TabControl1.SelectedTab = IO
    End Sub



    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Feature_manager_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        show_s_t.Text = "N"
        vfd_type.Text = "none"


        Try

            '--------- load solutions ---------------

            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT solution_name from quote_table.feature_solutions order by solution_name"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    sol_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

            sol_box.Text = "EIP-MS/EIP-RIO" '"SWDMS/EIPRIO"

            '-------- load feature code tables -----------
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT feature_code, description, type ,VFD_Type, specific_type, show_menu from quote_table.feature_codes where Solution = @sol"
            cmd.Connection = Login.Connection

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter(cmd)    'DataViewGrid1 fill
            adapter.Fill(table)
            Feature_Table.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Feature_Table.ColumnCount - 1
                With Feature_Table.Columns(i)
                    .Width = 300
                End With
            Next i

            'Fill the all parts table where the user can select the parts to be added to the feature code
            Dim tabl2 As New DataTable
            Dim ad2 As New MySqlDataAdapter("Select Part_Name, Part_Description, Part_Type from parts_table", Login.Connection)
            ad2.Fill(tabl2)
            allParts.DataSource = tabl2

            'Setting Columns size 
            For i = 0 To allParts.ColumnCount - 1
                With allParts.Columns(i)
                    .Width = 240
                End With
            Next i

            allParts.Columns(1).Width = 600



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

    Private Sub Feature_Table_DoubleClick(sender As Object, e As EventArgs) Handles Feature_Table.DoubleClick
        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW
        ListBox1.Items.Clear()
        remote_check.Checked = False
        FDR.Text = ""
        HDR.Text = ""
        FLA.Text = ""
        c_o.Text = "" : c_c.Text = "" : c_i.Text = "" : c_io.Text = "" : c_m.Text = ""
        o_b.Text = "" : i_b.Text = "" : io_b.Text = "" : c_b.Text = "" : m_b.Text = ""

        Dim feature_n = Feature_Table.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@feature_c", feature_n)
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT * from quote_table.feature_codes where Feature_code = @feature_c and Solution = @sol"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    feature_t.Text = reader(0).ToString
                    desc_t.Text = reader(1).ToString
                    type_t.Text = reader(4).ToString
                    vfd_type.Text = reader(5).ToString
                    specific_t.Text = reader(6).ToString
                    show_s_t.Text = reader(7).ToString
                    sol_def.Text = reader(8).ToString
                    labor_box.Text = reader(10).ToString
                    bulk_box.Text = reader(11).ToString

                End While

            End If

            reader.Close()

            If String.Equals(type_t.Text, "Panel") = True Or String.Equals(type_t.Text, "Control Panel") = True Then

                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@feature_c", feature_t.Text)
                cmd3.Parameters.AddWithValue("@sol", sol_box.Text)
                cmd3.CommandText = "SELECT Full_DR, Half_DR, FLA from quote_table.f_dimensions where feature_code = @feature_c and Solution = @sol"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        FDR.Text = reader3(0).ToString
                        HDR.Text = reader3(1).ToString
                        FLA.Text = reader3(2).ToString
                    End While
                End If

                reader3.Close()

            End If

            If String.Equals(type_t.Text, "Field") = True Then

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@feature_c", feature_t.Text)
                cmd4.Parameters.AddWithValue("@sol", sol_box.Text)
                cmd4.CommandText = "SELECT input, output, motion  from quote_table.IO_points where feature_code = @feature_c and solution = @sol"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        i_b.Text = reader4(0).ToString
                        o_b.Text = reader4(1).ToString
                        m_b.Text = reader4(2).ToString
                    End While
                End If

                reader4.Close()

                '------- Remote IO --------
                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.AddWithValue("@feature_c", feature_t.Text)
                cmd5.Parameters.AddWithValue("@sol", sol_box.Text)
                cmd5.CommandText = "SELECT inputs, output, motion, iolink, counter  from quote_table.Remote_IO where feature_code = @feature_c and solutiom = @sol"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    remote_check.Checked = True
                    While reader5.Read
                        c_i.Text = reader5(0).ToString
                        c_o.Text = reader5(1).ToString
                        c_m.Text = reader5(2).ToString
                        c_io.Text = reader5(3).ToString
                        c_c.Text = reader5(4).ToString

                    End While
                Else
                    remote_check.Checked = False
                End If

                reader5.Close()

                '------------------------------
            End If


            '------- formulas --------
            Dim cmd_v As New MySqlCommand
            cmd_v.Parameters.AddWithValue("@feau", feature_t.Text)
            cmd_v.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd_v.CommandText = "SELECT formula from quote_table.call_feature where feature_code = @feau and solution = @sol"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    ListBox1.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()
            '-------------------------------


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        'get temp values

        Call Populate_devices()  'fill BOM

        temp_Feature_code = feature_t.Text
        temp_description = desc_t.Text
        temp_type = type_t.Text
        temp_vfd_type = vfd_type.Text
        temp_specific_Type = specific_t.Text
        temp_show = show_s_t.Text
        temp_sol_def = sol_def.Text





    End Sub

    Sub Populate_devices()

        '--------------- populate part members -----------------
        Try
            Dim cmdstr As String : cmdstr = "Select part_name, qty from quote_table.feature_parts where Feature_code  = @feature_code and solution = @solution"
            Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
            cmd_k.Parameters.AddWithValue("@feature_code", feature_t.Text)
            cmd_k.Parameters.AddWithValue("@solution", sol_box.Text)

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

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        '----------- update feature code -------
        If String.Equals(feature_t.Text, "") = True Or String.Equals(sol_box.Text, "") = True Or String.IsNullOrEmpty(feature_t.Text) = True Then
            MessageBox.Show("Please Fill Feature code textbox")
        Else

            Dim update_flag As Boolean : update_flag = True

            Try
                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Feature_code", temp_Feature_code)
                cmd.Parameters.AddWithValue("@description", temp_description)
                cmd.Parameters.AddWithValue("@Solution", sol_box.Text)
                ' cmd.Parameters.AddWithValue("@solution_description", sol_des)
                cmd.Parameters.AddWithValue("@type", temp_type)
                cmd.Parameters.AddWithValue("@VFD_TYPE", If(String.IsNullOrEmpty(temp_vfd_type) = True, "none", temp_vfd_type))
                cmd.Parameters.AddWithValue("@specific_type", temp_specific_Type)
                ' cmd.Parameters.AddWithValue("@show_menu", temp_show)
                ' cmd.Parameters.AddWithValue("@solution_default", temp_sol_def)
                ' cmd.Parameters.AddWithValue("@assembly", If((temp_assembly = True), "Y", ""))

                cmd.CommandText = "SELECT Feature_code from quote_table.feature_codes where Feature_code = @Feature_code and description = @description and Solution = @Solution and type = @type and VFD_TYPE = @VFD_TYPE and specific_type = @specific_type"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                'MAKE SURE THE PART EXIST SO IT CAN BE UPDATED
                If Not reader.HasRows Then
                    MessageBox.Show("Feature code does not exist! ... Update incomplete")
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
                    'temps (where clause)
                    Create_cmd.Parameters.AddWithValue("@Feature_code", temp_Feature_code)
                    ' Create_cmd.Parameters.AddWithValue("@description", temp_description)
                    Create_cmd.Parameters.AddWithValue("@Solution", sol_box.Text)
                    ' Create_cmd.Parameters.AddWithValue("@type", temp_type)
                    ' Create_cmd.Parameters.AddWithValue("@VFD_TYPE", If(String.IsNullOrEmpty(temp_vfd_type) = True, "none", temp_vfd_type))
                    ' Create_cmd.Parameters.AddWithValue("@specific_type", temp_specific_Type)
                    '  Create_cmd.Parameters.AddWithValue("@show_menu", temp_show)
                    '  Create_cmd.Parameters.AddWithValue("@solution_default", temp_sol_def)

                    'new values
                    Create_cmd.Parameters.AddWithValue("@nFeature_code", feature_t.Text)
                    Create_cmd.Parameters.AddWithValue("@ndescription", desc_t.Text)
                    Create_cmd.Parameters.AddWithValue("@ntype", type_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nVFD_TYPE", If(String.IsNullOrEmpty(vfd_type.Text) = True, "none", vfd_type.Text))
                    Create_cmd.Parameters.AddWithValue("@nspecific_type", specific_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nshow_menu", show_s_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nsolution_default", sol_def.Text)
                    Create_cmd.Parameters.AddWithValue("@nlabor_cost", labor_box.Text)
                    Create_cmd.Parameters.AddWithValue("@nbulk_Cost", bulk_box.Text)

                    Create_cmd.CommandText = "UPDATE quote_table.feature_codes SET Feature_code = @nFeature_code, description = @ndescription, type = @ntype, VFD_TYPE =  @nVFD_TYPE, specific_type = @nspecific_type, show_menu = @nshow_menu, solution_default = @nsolution_default, labor_cost = @nlabor_cost, bulk_cost = @nbulk_cost where Feature_code = @Feature_code and Solution =  @Solution"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    '---call feature

                    Create_cmd.CommandText = "UPDATE quote_table.call_feature SET feature_code = @nFeature_code, description = @ndescription where feature_code = @Feature_code and Solution =  @Solution"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    '---------- f_dimensions
                    Create_cmd.CommandText = "UPDATE quote_table.f_dimensions SET feature_code = @nFeature_code where feature_code = @Feature_code  and Solution =  @Solution"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    '----- feature_parts
                    Create_cmd.CommandText = "UPDATE quote_table.feature_parts SET feature_code = @nFeature_code, type = @ntype where feature_code = @Feature_code  and Solution =  @Solution"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()


                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Feature code updated succesfully")

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '----------------------------------- CREATE FEATURE CODE ------------------------------------------


        '---------MAKE SURE TEXTBOXES ARE FILL UP (PART, STATUS AND TYPE) ------------------------
        If String.Equals(feature_t.Text, "") = True Or String.Equals(type_t.Text, "") = True Or String.Equals(specific_t.Text, "") = True Then
            MessageBox.Show("Please Fill Feature code name, solution, type and specific type")
        Else

            '----------- MAKE SURE THE FEATURE_CODE DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@feature", feature_t.Text)
                cmd.Parameters.AddWithValue("@solution", sol_box.Text)
                cmd.Parameters.AddWithValue("@type", type_t.Text)
                cmd.Parameters.AddWithValue("@specific_type", specific_t.Text)
                cmd.CommandText = "SELECT Feature_code from quote_table.feature_codes where Feature_code = @feature and Solution = @solution and type = @type and specific_type = @specific_type"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@nFeature_code", feature_t.Text)
                    Create_cmd.Parameters.AddWithValue("@ndescription", desc_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nSolution", sol_box.Text)
                    Create_cmd.Parameters.AddWithValue("@nsolution_description", sol_des.Text)
                    Create_cmd.Parameters.AddWithValue("@ntype", type_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nVFD_TYPE", If(String.IsNullOrEmpty(vfd_type.Text) = True, "none", vfd_type.Text))
                    Create_cmd.Parameters.AddWithValue("@nspecific_type", specific_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nshow_menu", show_s_t.Text)
                    Create_cmd.Parameters.AddWithValue("@nsolution_default", sol_def.Text)
                    Create_cmd.Parameters.AddWithValue("@nlabor_cost", labor_box.Text)
                    Create_cmd.Parameters.AddWithValue("@nbulk_Cost", bulk_box.Text)
                    Create_cmd.Parameters.AddWithValue("@assembly", "")
                    Create_cmd.CommandText = "INSERT INTO quote_table.feature_codes(Feature_code, description, Solution, solution_description, type, VFD_TYPE, specific_type, show_menu, solution_default, assembly, labor_cost, bulk_cost) VALUES (@nFeature_code, @ndescription, @nSolution, @nsolution_description, @ntype, @nVFD_TYPE, @nspecific_type, @nshow_menu, @nsolution_default, @assembly, @nlabor_cost, @nbulk_cost)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Feature code added succesfully!")

                Else
                    MessageBox.Show("Feature code already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If


            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Sub Refresh_tables(c As MySqlConnection)

        '---------- REFRESH DATA IN GRID ----------------------
        Try
            'REFRESH PARTS TABLE 
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT feature_code, description, type ,VFD_Type, specific_type, show_menu from quote_table.feature_codes where Solution = @sol"
            cmd.Connection = Login.Connection

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter(cmd)    'DataViewGrid1 fill
            adapter.Fill(table)
            Feature_Table.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Feature_Table.ColumnCount - 1
                With Feature_Table.Columns(i)
                    .Width = 300
                End With
            Next i

        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '-------- Add components to the feature parts bom ----------------
        If (allParts.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(allParts.CurrentCell.Value.ToString, "") = False And String.Equals(feature_t.Text, "") = False And String.IsNullOrEmpty(feature_t.Text) = False And String.IsNullOrEmpty(type_t.Text) = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString & " this feature code will allocate")
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ add components to adv table ----------------------------------
                        Try
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@Part_Name", allParts.CurrentCell.Value.ToString)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.Parameters.AddWithValue("@type", type_t.Text)
                            Create_cmd.Parameters.AddWithValue("@feature_code", feature_t.Text)
                            Create_cmd.Parameters.AddWithValue("@solution", sol_box.Text)
                            Create_cmd.CommandText = "INSERT INTO quote_table.feature_parts VALUES (@feature_code, @solution, @type, @Part_Name, @qty)"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Call Populate_devices()
                            MessageBox.Show(allParts.CurrentCell.Value.ToString & " added succesfully")

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

    Private Sub sol_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles sol_box.SelectedValueChanged
        Try

            '-------- get solution description ------
            Dim cmd_v As New MySqlCommand
            cmd_v.Parameters.AddWithValue("@sol_name", sol_box.Text)
            cmd_v.CommandText = "SELECT sol_description from quote_table.feature_solutions where solution_name = @sol_name"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    sol_des.Text = readerv(0).ToString
                End While
            End If

            readerv.Close()


            '-------- load feature code tables -----------
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT feature_code, description, type ,VFD_Type, specific_type, show_menu from quote_table.feature_codes where Solution = @sol"
            cmd.Connection = Login.Connection

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter(cmd)    'DataViewGrid1 fill
            adapter.Fill(table)
            Feature_Table.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Feature_Table.ColumnCount - 1
                With Feature_Table.Columns(i)
                    .Width = 300
                End With
            Next i

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If (current_p.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(current_p.CurrentCell.Value.ToString, "") = False Then

                Try
                    Dim cmd As New MySqlCommand
                    cmd.Parameters.AddWithValue("@Part_Name", current_p.CurrentCell.Value.ToString)
                    cmd.Parameters.AddWithValue("@sol", sol_box.Text)
                    cmd.Parameters.AddWithValue("@fea", feature_t.Text)
                    cmd.CommandText = "DELETE FROM quote_table.feature_parts where Part_Name = @Part_Name and feature_code = @fea and solution = @sol"
                    cmd.Connection = Login.Connection
                    cmd.ExecuteNonQuery()

                    Dim cmdstr As String : cmdstr = "Select part_name, qty from quote_table.feature_parts where Feature_code  = @feature_code and solution = @solution"
                    Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
                    cmd_k.Parameters.AddWithValue("@feature_code", feature_t.Text)
                    cmd_k.Parameters.AddWithValue("@solution", sol_box.Text)

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

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
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
                            Create_cmd.Parameters.AddWithValue("@fea", feature_t.Text.ToString)
                            Create_cmd.Parameters.AddWithValue("@sol", sol_box.Text)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.CommandText = "UPDATE quote_table.feature_parts  SET  Qty = @qty where Part_Name = @Part_Name and feature_code = @fea and solution = @sol"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
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

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Add_solutions.ShowDialog()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try

            If String.Equals(feature_t.Text, "") = True Or String.Equals(sol_box.Text, "") = True Or String.IsNullOrEmpty(feature_t.Text) = True Then
                MessageBox.Show("Please Fill Feature code textbox")
            Else
                Try
                    Dim cmd As New MySqlCommand
                    Dim Create_cmd As New MySqlCommand
                    Dim Create_vd As New MySqlCommand
                    Dim Create_acn As New MySqlCommand
                    Dim Create_cmd2 As New MySqlCommand
                    Dim Create_vd2 As New MySqlCommand
                    Dim Create_acn2 As New MySqlCommand

                    cmd.Parameters.AddWithValue("@feature_code", feature_t.Text)
                    cmd.CommandText = "SELECT *  from quote_table.feature_codes where Feature_code = @feature_code"
                    cmd.Connection = Login.Connection
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader

                    If reader.HasRows Then

                        Dim dlgR As DialogResult
                        dlgR = MessageBox.Show("Are you sure you want to delete this feature code?", "Attention!", MessageBoxButtons.YesNo)
                        If dlgR = DialogResult.Yes Then

                            reader.Close()

                            Create_cmd.Parameters.AddWithValue("@feature_code", feature_t.Text)  'call feature
                            Create_vd.Parameters.AddWithValue("@feature_code", feature_t.Text)   'f_dimensions
                            Create_acn.Parameters.AddWithValue("@feature_code", feature_t.Text)  'feature_codes
                            Create_cmd2.Parameters.AddWithValue("@feature_code", feature_t.Text)  'feature parts
                            Create_vd2.Parameters.AddWithValue("@feature_code", feature_t.Text)   'io_points
                            Create_acn2.Parameters.AddWithValue("@feature_code", feature_t.Text)  'Remote io

                            Create_cmd.Parameters.AddWithValue("@sol", sol_box.Text)  'call feature
                            Create_vd.Parameters.AddWithValue("@sol", sol_box.Text)   'f_dimensions
                            Create_acn.Parameters.AddWithValue("@sol", sol_box.Text)  'feature_codes
                            Create_cmd2.Parameters.AddWithValue("@sol", sol_box.Text)  'feature parts
                            Create_vd2.Parameters.AddWithValue("@sol", sol_box.Text)   'io_points
                            Create_acn2.Parameters.AddWithValue("@sol", sol_box.Text)

                            Create_cmd.CommandText = "DELETE FROM quote_table.call_feature where feature_code = @feature_code and solution = @sol"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Create_vd.CommandText = "DELETE FROM quote_table.f_dimensions where feature_code = @feature_code and Solution = @sol"
                            Create_vd.Connection = Login.Connection
                            Create_vd.ExecuteNonQuery()
                            Create_acn.CommandText = "DELETE FROM quote_table.feature_codes where feature_code = @feature_code and Solution = @sol"
                            Create_acn.Connection = Login.Connection
                            Create_acn.ExecuteNonQuery()

                            Create_cmd2.CommandText = "DELETE FROM quote_table.feature_parts where feature_code = @feature_code and solution = @sol"
                            Create_cmd2.Connection = Login.Connection
                            Create_cmd2.ExecuteNonQuery()
                            Create_vd2.CommandText = "DELETE FROM quote_table.IO_points where feature_code = @feature_code and solution = @sol"
                            Create_vd2.Connection = Login.Connection
                            Create_vd2.ExecuteNonQuery()
                            Create_acn2.CommandText = "DELETE FROM quote_table.Remote_IO where feature_code = @feature_code and solutiom = @sol"
                            Create_acn2.Connection = Login.Connection
                            Create_acn2.ExecuteNonQuery()


                            MessageBox.Show("feature coded deleted succesfully")
                            Call Refresh_tables(Login.Connection)

                            feature_t.Text = ""
                            desc_t.Text = ""
                            type_t.Text = ""
                            vfd_type.Text = ""
                            specific_t.Text = ""
                            show_s_t.Text = ""
                            sol_def.Text = ""

                            ListBox1.Items.Clear()
                            remote_check.Checked = False
                            FDR.Text = ""
                            HDR.Text = ""
                            FLA.Text = ""
                            c_o.Text = "" : c_c.Text = "" : c_i.Text = "" : c_io.Text = "" : c_m.Text = ""
                            o_b.Text = "" : i_b.Text = "" : io_b.Text = "" : c_b.Text = "" : m_b.Text = ""

                            temp_Feature_code = feature_t.Text
                            temp_description = desc_t.Text
                            temp_type = type_t.Text
                            temp_vfd_type = vfd_type.Text
                            temp_specific_Type = specific_t.Text
                            temp_show = show_s_t.Text
                            temp_sol_def = sol_def.Text

                        Else
                            reader.Close()

                        End If
                    Else
                        MessageBox.Show("We could not find the feature code specified to be removed")
                        reader.Close()
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)

                End Try
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
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