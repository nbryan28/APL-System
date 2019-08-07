Imports MySql.Data.MySqlClient

Public Class Open_file

    Public mpl_y As Boolean  'check if you are dealing with an mpl


    Private Sub Open_file_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Clear()
        open_grid.Rows.Clear()

        Select Case My_Material_r.mode

            '-- In progress Mr
            Case "mr_inpro"

                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.CommandText = "select mr_name, Date_Created, created_by, last_modified from Material_Request.mr  where released is null and (BOM_type = 'old_BOM') order by mr_name "
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        Dim i As Integer : i = 0
                        While reader.Read
                            open_grid.Rows.Add(New String() {})
                            open_grid.Rows(i).Cells(0).Value = reader(0).ToString
                            open_grid.Rows(i).Cells(1).Value = reader(1).ToString
                            open_grid.Rows(i).Cells(2).Value = reader(2).ToString
                            open_grid.Rows(i).Cells(3).Value = reader(3).ToString
                            i = i + 1
                        End While
                    End If

                    reader.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                '------ Released Mrs
            Case "mr_rel"

                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            ComboBox1.Items.Add(reader(0))
                        End While
                    End If

                    reader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                '--- released mr for procurement
            Case "pc_rel"


                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            ComboBox1.Items.Add(reader(0))
                        End While
                    End If

                    reader.Close()


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Case "inv_rel"

                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    ' check_cmd.CommandText = "select distinct job from Material_Request.mr where released_inv = 'Y' order by job"
                    check_cmd.CommandText = "select distinct job from Material_Request.mr where job is not null order by job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            ComboBox1.Items.Add(reader(0))
                        End While
                    End If

                    reader.Close()


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            Case "mpl_mfg"

                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            ComboBox1.Items.Add(reader(0))
                        End While
                    End If

                    reader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Case "br_mfg"

                mpl_y = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            ComboBox1.Items.Add(reader(0))
                        End While
                    End If

                    reader.Close()

                    '-- remove jobs that dont have mr with assemblies or panels
                    For i = ComboBox1.Items.Count - 1 To 0 Step -1
                        If My_Material_r.job_with_build(ComboBox1.Items(i)) = False Then
                            ComboBox1.Items.RemoveAt(i)
                        End If
                    Next

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
        End Select
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        open_grid.Rows.Clear()

        Try

            If mpl_y = False Then

                '------ get all BOM of the job Note: unelegant way of do this. Try to find a better Query
                Dim n_bom As Double : n_bom = 0
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
                check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        n_bom = reader(0)
                    End While
                End If

                reader.Close()

                For i = 1 To n_bom

                    Dim check_cmd1 As New MySqlCommand
                    check_cmd1.Parameters.Clear()
                    check_cmd1.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
                    check_cmd1.Parameters.AddWithValue("@id_bom", i)
                    check_cmd1.CommandText = "select mr_name,  Date_Created, created_by, last_modified , release_date, released_by from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then
                        ' Dim j As Integer : j = 0
                        While reader1.Read
                            open_grid.Rows.Add(New String() {reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString, reader1(5).ToString})
                        End While
                    End If

                    reader1.Close()
                Next
                '---------------------------------------------------------

            Else

            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim name As String = ""
        Dim index As Integer = 1

        If open_grid.Rows.Count > 0 Then
            name = If(String.IsNullOrEmpty(open_grid.CurrentCell.Value) = True, "", open_grid.CurrentCell.Value.ToString)
            index = open_grid.CurrentCell.ColumnIndex
        End If

        If String.Equals(name, "") = False And index = 0 Then

            Select Case My_Material_r.mode

                Case "mr_inpro"

                    My_Material_r.PR_grid.Rows.Clear()

                    Try
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", name)
                        cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name, Notes, Part_status, Part_type, full_panel, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then
                            Dim i As Integer : i = 0
                            While reader4.Read
                                My_Material_r.PR_grid.Rows.Add(New String() {})
                                My_Material_r.PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(14).Value = reader4(14).ToString

                                i = i + 1
                            End While

                        End If

                        reader4.Close()
                        My_Material_r.Text = name
                        My_Material_r.single_grid.Rows.Clear()

                        Me.Visible = False
                        My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage1

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                    For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
                        My_Material_r.PR_grid.Rows(i).ReadOnly = False
                    Next
                    My_Material_r.isitreleased = False
                    My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
                    Inventory_manage.part_sel = ""
                    My_Material_r.rev_mode = False

                Case "mr_rel"
                    '----------- Open a release MR ----------
                    My_Material_r.PR_grid.Rows.Clear()
                    ' My_Material_r.mrbox1.Items.Clear()
                    ' My_Material_r.mrbox2.Items.Clear()

                    Try
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", name)
                        cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name ,Notes, Part_status, Part_type, full_panel, need_by_date, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then
                            Dim i As Integer : i = 0
                            While reader4.Read
                                My_Material_r.PR_grid.Rows.Add(New String() {})
                                My_Material_r.PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(14).Value = reader4(14).ToString
                                My_Material_r.PR_grid.Rows(i).Cells(15).Value = reader4(15).ToString

                                i = i + 1
                            End While

                        End If

                        reader4.Close()
                        My_Material_r.Text = name
                        My_Material_r.single_grid.Rows.Clear()
                        My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage1
                        Me.Visible = False

                        '---------- fill combobox with all revisions -------
                        My_Material_r.ComboBox2.Items.Clear()
                        Dim check_cmd As New MySqlCommand
                        check_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
                        check_cmd.CommandText = "select distinct mr_name from Material_Request.mr where released = 'Y' and job = @job"
                        check_cmd.Connection = Login.Connection
                        check_cmd.ExecuteNonQuery()

                        Dim reader As MySqlDataReader
                        reader = check_cmd.ExecuteReader

                        If reader.HasRows Then
                            While reader.Read
                                My_Material_r.ComboBox2.Items.Add(reader(0))
                                'My_Material_r.mrbox1.Items.Add(reader(0))
                                ' My_Material_r.mrbox2.Items.Add(reader(0))
                            End While
                        End If

                        reader.Close()

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                    For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
                        My_Material_r.PR_grid.Rows(i).ReadOnly = True
                    Next

                    My_Material_r.isitreleased = True
                    My_Material_r.open_job = ComboBox1.SelectedItem.ToString
                    My_Material_r.job_label.Text = "Open Project: " & ComboBox1.SelectedItem.ToString
                    My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
                    Inventory_manage.part_sel = ""
                    My_Material_r.rev_mode = False

                Case "pc_rel"

                    '----------- Open a release MR ----------
                    Proc_Material_R.total_grid.Rows.Clear()
                    Proc_Material_R.start_e = False

                    Try
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", name)
                        cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, mfg_type, ADA_Number, Vendor, Price, subtotal, Qty, qty_fullfilled, qty_needed, status_m , Assembly_name, Part_status, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name and (full_panel is null or full_panel <> 'Yes')"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then
                            Dim i As Integer : i = 0
                            While reader4.Read
                                Proc_Material_R.total_grid.Rows.Add(New String() {})
                                Proc_Material_R.total_grid.Rows(i).Cells(0).Value = reader4(8).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(1).Value = reader4(0).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(2).Value = reader4(1).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(3).Value = reader4(2).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(4).Value = reader4(3).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(6).Value = If(IsNumeric(reader4(6)) = True, reader4(6).ToString, 0)
                                Proc_Material_R.total_grid.Rows(i).Cells(7).Value = If(IsNumeric(reader4(8)) = True, reader4(8).ToString, 0) * If(IsNumeric(reader4(6)) = True, reader4(6).ToString, 0)
                                Proc_Material_R.total_grid.Rows(i).Cells(8).Value = If(IsNumeric(reader4(9)) = True, reader4(9).ToString, 0)
                                Proc_Material_R.total_grid.Rows(i).Cells(9).Value = If(IsNumeric(reader4(10)) = True, reader4(10).ToString, 0)
                                Proc_Material_R.total_grid.Rows(i).Cells(10).Value = reader4(11).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(11).Value = reader4(15).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(12).Value = reader4(14).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(13).Value = reader4(12).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(14).Value = reader4(13).ToString
                                Proc_Material_R.total_grid.Rows(i).Cells(15).Value = reader4(4).ToString
                                i = i + 1
                            End While

                        End If

                        reader4.Close()

                        '==================================================
                        Dim cmd5 As New MySqlCommand
                        cmd5.Parameters.Clear()
                        cmd5.Parameters.AddWithValue("@mr_name", name)
                        cmd5.CommandText = "SELECT shipping_ad from Material_Request.mr where mr_name = @mr_name"
                        cmd5.Connection = Login.Connection
                            Dim reader5 As MySqlDataReader
                        reader5 = cmd5.ExecuteReader

                        If reader5.HasRows Then
                            While reader5.Read
                                Proc_Material_R.shipping_b.Text = reader5(0).ToString
                            End While
                        End If

                        reader5.Close()
                        '==================================================

                        Proc_Material_R.Text = name
                        Proc_Material_R.open_job = ComboBox1.SelectedItem.ToString
                        Proc_Material_R.job_label.Text = ComboBox1.SelectedItem.ToString
                        Proc_Material_R.PR_grid.Rows.Clear()
                        Proc_Material_R.compare_grid.Rows.Clear()

                        Me.Visible = False

                        Proc_Material_R.ComboBox2.Items.Clear()
                        Proc_Material_R.mrbox1.Items.Clear()
                        Proc_Material_R.mrbox2.Items.Clear()



                        Dim check_cmd As New MySqlCommand
                        check_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
                        check_cmd.CommandText = "select distinct mr_name from Material_Request.mr where released = 'Y' and job = @job"
                        check_cmd.Connection = Login.Connection
                        check_cmd.ExecuteNonQuery()

                        Dim reader As MySqlDataReader
                        reader = check_cmd.ExecuteReader

                        If reader.HasRows Then
                            While reader.Read
                                Proc_Material_R.ComboBox2.Items.Add(reader(0))
                                Proc_Material_R.mrbox1.Items.Add(reader(0))
                                Proc_Material_R.mrbox2.Items.Add(reader(0))
                            End While
                        End If

                        reader.Close()

                        Proc_Material_R.start_e = True

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                Case "inv_rel"

                    '----------- Open a release MR ----------
                    Inv_Material_r.fullfill_grid.Rows.Clear()

                    '---- datatable store MR without assemblies --------
                    Dim dimen_table = New DataTable
                    dimen_table.Columns.Add("Part_No", GetType(String))
                    dimen_table.Columns.Add("Description", GetType(String))
                    dimen_table.Columns.Add("Manufacturer", GetType(String))
                    dimen_table.Columns.Add("ADA_Number", GetType(String))
                    dimen_table.Columns.Add("Vendor", GetType(String))
                    dimen_table.Columns.Add("Price", GetType(String))
                    dimen_table.Columns.Add("mfg_type", GetType(String))
                    dimen_table.Columns.Add("Qty", GetType(String))
                    dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                    dimen_table.Columns.Add("qty_needed", GetType(String))
                    dimen_table.Columns.Add("Assembly_name", GetType(String))
                    dimen_table.Columns.Add("Part_status", GetType(String))
                    dimen_table.Columns.Add("Notes", GetType(String))


                    Try

                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", name)
                        cmd4.CommandText = "SELECT Part_No, Description, Manufacturer,  ADA_Number, Vendor, Price, mfg_type, Qty, qty_fullfilled, qty_needed , Assembly_name, Part_status, Notes from Material_Request.mr_data where mr_name = @mr_name and (full_panel is null or full_panel <> 'Yes')"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader


                        If reader4.HasRows Then
                            Dim i As Integer : i = 0
                            While reader4.Read
                                dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString, reader4(12).ToString)
                            End While

                        End If

                        reader4.Close()


                        '------- now copy dimen_table_a to inventroy grid
                        For i = 0 To dimen_table.Rows.Count - 1
                            Inv_Material_r.fullfill_grid.Rows.Add(New String() {})
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(8).Value = 0
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(9).Value = 0
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(10).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(11).Value = If(IsNumeric(dimen_table.Rows(i).Item(9).ToString) = True, dimen_table.Rows(i).Item(9).ToString, 0)
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(12).ToString
                            Inv_Material_r.fullfill_grid.Rows(i).Cells(14).Value = dimen_table.Rows(i).Item(10).ToString
                        Next


                        '////////////////////  ---------- fill current inventory values -----------------------

                        For i = 0 To Inv_Material_r.fullfill_grid.Rows.Count - 1

                            Dim cmd5 As New MySqlCommand
                            cmd5.Parameters.Clear()
                            cmd5.Parameters.AddWithValue("@part_name", Inv_Material_r.fullfill_grid.Rows(i).Cells(0).Value)
                            cmd5.CommandText = "SELECT current_qty, min_qty from inventory.inventory_qty where part_name = @part_name"
                            cmd5.Connection = Login.Connection
                            Dim reader5 As MySqlDataReader
                            reader5 = cmd5.ExecuteReader

                            If reader5.HasRows Then
                                While reader5.Read
                                    Inv_Material_r.fullfill_grid.Rows(i).Cells(8).Value = reader5(0).ToString

                                    If reader5(1) > 0 Then
                                        Inv_Material_r.fullfill_grid.Rows(i).Cells(12).Value = "Yes"
                                    Else
                                        Inv_Material_r.fullfill_grid.Rows(i).Cells(12).Value = "no"
                                    End If
                                End While

                            Else
                                Inv_Material_r.fullfill_grid.Rows(i).Cells(12).Value = "no"
                            End If

                            reader5.Close()

                        Next
                        '----------------------------------------------------------------

                        Inv_Material_r.Text = name
                        Inv_Material_r.open_job = ComboBox1.SelectedItem.ToString
                        Inv_Material_r.job_label.Text = ComboBox1.SelectedItem.ToString
                        Me.Visible = False

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try


                Case "mpl_mfg"

                    Try
                        Dim cmd5 As New MySqlCommand
                        cmd5.Parameters.Clear()
                        cmd5.Parameters.AddWithValue("@mr_name", name)
                        cmd5.CommandText = "SELECT shipping_ad from Material_Request.mr where mr_name = @mr_name"
                        cmd5.Connection = Login.Connection
                        Dim reader5 As MySqlDataReader
                        reader5 = cmd5.ExecuteReader

                        If reader5.HasRows Then
                            While reader5.Read
                                MPL_mfg.shipping_l.Text = reader5(0).ToString
                            End While
                        End If

                        reader5.Close()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try


                    MPL_mfg.PR_grid.Rows.Clear()
                    Call My_Material_r.Packing_List(name, MPL_mfg.PR_grid)
                    MPL_mfg.Text = name
                    MPL_mfg.job_label.Text = ComboBox1.SelectedItem.ToString
                    Me.Visible = False

                Case "br_mfg"

                    Try
                        Dim cmd5 As New MySqlCommand
                        cmd5.Parameters.Clear()
                        cmd5.Parameters.AddWithValue("@mr_name", name)
                        cmd5.CommandText = "SELECT shipping_ad from Material_Request.mr where mr_name = @mr_name"
                        cmd5.Connection = Login.Connection
                        Dim reader5 As MySqlDataReader
                        reader5 = cmd5.ExecuteReader

                        If reader5.HasRows Then
                            While reader5.Read
                                Build_mfg.shipping_l.Text = reader5(0).ToString
                            End While
                        End If

                        reader5.Close()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                    Build_mfg.PR_grid.Rows.Clear()
                    Call My_Material_r.Build_Request(name, Build_mfg.PR_grid)
                    Build_mfg.Text = name
                    Build_mfg.job_label.Text = ComboBox1.SelectedItem.ToString
                    Me.Visible = False

            End Select

        End If
    End Sub

    '----- this sub break the assembly an enter the data in a datatable
    Sub break_assembly_inv(assembly As String, qty As Double, datatable As DataTable)

        'Dim Atronix_n As String : Atronix_n = ""

        'Try
        '    Dim cmd_a As New MySqlCommand
        '    cmd_a.Parameters.AddWithValue("@assembly", assembly)
        '    cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @assembly"
        '    cmd_a.Connection = Login.Connection

        '    Dim reader_k As MySqlDataReader
        '    reader_k = cmd_a.ExecuteReader


        '    If reader_k.HasRows Then
        '        While reader_k.Read
        '            Atronix_n = reader_k(0)
        '        End While
        '    End If

        '    reader_k.Close()

        '    Dim cmd_pd As New MySqlCommand
        '    cmd_pd.Parameters.AddWithValue("@adv_n", Atronix_n)
        '    cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, p1.Manufacturer,  p1.Primary_Vendor, adv.Qty, p1.MFG_type, p1.Part_Status, p1.Part_Type from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number =  @adv_n"
        '    cmd_pd.Connection = Login.Connection
        '    Dim readerv As MySqlDataReader
        '    readerv = cmd_pd.ExecuteReader

        '    If readerv.HasRows Then
        '        While readerv.Read
        '            ' datatable.Rows.Add(readerv(0).ToString, readerv(1).ToString, readerv(2).ToString, readerv(5).ToString, "", readerv(3).ToString, "", readerv(4) * qty, reader4(7).ToString, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString)
        '        End While
        '    End If

        '    readerv.Close()

        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try
    End Sub
End Class