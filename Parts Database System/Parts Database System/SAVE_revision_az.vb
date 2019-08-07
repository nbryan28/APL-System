Imports MySql.Data.MySqlClient

Public Class SAVE_revision_az

    Public mode_t As String
    Public mr_name As String
    Public desc_t As String
    Public p_qty As String
    Public panel_n As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Select Case mode_t

            Case "add"
                Call evaluate_add()
            Case "zero"
                Call evaluate_delete()
            Case "change"
                Call evaluate_change()

        End Select

    End Sub



    Sub evaluate_add()


        Dim exist_c As Boolean : exist_c = False

        If String.IsNullOrEmpty(TextBox1.Text) = False And String.Equals("", TextBox1.Text) = False Then

            Try

                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                check_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                check_cmd.CommandText = "select * from Revisions.mr_rev where mr_name = @mr_name and rev_name = @rev_name"
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    exist_c = True
                End If

                reader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            If exist_c = False Then

                Try

                    '-------- get MBOM ------------
                    Dim MB As String : MB = ""

                    Dim cmdm As New MySqlCommand
                    cmdm.Parameters.AddWithValue("@job", My_Material_r.open_job)
                    cmdm.Parameters.AddWithValue("@BOM_type", "MB")
                    cmdm.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type  and job = @job order by release_date desc limit 1"
                    cmdm.Connection = Login.Connection
                    Dim readerm As MySqlDataReader
                    readerm = cmdm.ExecuteReader

                    If readerm.HasRows Then
                        While readerm.Read
                            MB = readerm(0).ToString
                        End While
                    End If

                    readerm.Close()
                    '----------------------------


                    '--- enter data to mr -------
                    Dim main_cmd As New MySqlCommand
                    main_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                    main_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                    main_cmd.Parameters.AddWithValue("@qty_name", Edit_BOM_Panels.qty_b.Text)
                    main_cmd.Parameters.AddWithValue("@desc_name", Edit_BOM_Panels.panel_desc.Text)
                    main_cmd.Parameters.AddWithValue("@panel_name", panel_n)
                    main_cmd.Parameters.AddWithValue("@created_by", current_user)
                    main_cmd.Parameters.AddWithValue("@MB", MB)
                    main_cmd.CommandText = "INSERT INTO Revisions.mr_rev(mr_name, rev_name, created_date , created_by, new_panel, MB, panel_name, desc_name, qty_name) VALUES (@mr_name, @rev_name, now(), @created_by, 'Add', @MB, @panel_name, @desc_name, @qty_name)"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()


                    '-------- enter data to mr_data
                    For i = 0 To Edit_BOM_Panels.Panel_grid.Rows.Count - 1

                        If Edit_BOM_Panels.Panel_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                            Dim Create_cmd6 As New MySqlCommand
                            Create_cmd6.Parameters.Clear()
                            Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name)
                            Create_cmd6.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(0).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(0).Value))
                            Create_cmd6.Parameters.AddWithValue("@description_t", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(1).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(1).Value))
                            Create_cmd6.Parameters.AddWithValue("@ADA_Number", "")
                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(2).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(2).Value))
                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(3).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(3).Value))
                            Create_cmd6.Parameters.AddWithValue("@Price", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(4).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(4).Value.ToString.Replace("$", "")))
                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(7).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(7).Value))
                            Create_cmd6.Parameters.AddWithValue("@current_qty", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(5).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(5).Value))
                            Create_cmd6.Parameters.AddWithValue("@new_qty", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(5).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(5).Value))
                            Create_cmd6.Parameters.AddWithValue("@delta", 0)
                            Create_cmd6.Parameters.AddWithValue("@assembly", "")
                            Create_cmd6.Parameters.AddWithValue("@Notes", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(10).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(10).Value))
                            Create_cmd6.Parameters.AddWithValue("@part_status", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(8).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(8).Value))
                            Create_cmd6.Parameters.AddWithValue("@part_type", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(9).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(9).Value))
                            Create_cmd6.Parameters.AddWithValue("@isitfull", "")
                            Create_cmd6.Parameters.AddWithValue("@need_date", If(Edit_BOM_Panels.Panel_grid.Rows(i).Cells(11).Value Is Nothing, "", Edit_BOM_Panels.Panel_grid.Rows(i).Cells(11).Value))

                            Create_cmd6.CommandText = "INSERT INTO Revisions.mr_rev_data(mr_name, rev_name, Part_No, description_t, ADA_Number, Manufacturer, Vendor, Price, mfg_type, current_qty, new_qty, delta, assembly, Notes, part_status, part_type, isitfull, need_date) VALUES (@mr_name, @rev_name, @Part_No, @description_t, @ADA_Number, @Manufacturer, @Vendor, @Price, @mfg_type, @current_qty, @new_qty, @delta, @assembly, @Notes, @part_status, @part_type, @isitfull, @need_date)"
                            Create_cmd6.Connection = Login.Connection
                            Create_cmd6.ExecuteNonQuery()

                        End If

                    Next

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                Me.Visible = False

            Else
                MessageBox.Show("Revision file already exist!")
            End If

        Else
            MessageBox.Show("Please enter revision name!")
        End If
    End Sub

    Sub evaluate_delete()
        '---- zero out panel -------

        Dim exist_c As Boolean : exist_c = False

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
            check_cmd.CommandText = "select * from Revisions.mr_rev where mr_name = @mr_name and rev_name = @rev_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                exist_c = True
            End If

            reader.Close()



            If exist_c = False Then

                '  -------- get MBOM ------------
                Dim MB As String : MB = ""

                Dim cmdm As New MySqlCommand
                cmdm.Parameters.AddWithValue("@job", My_Material_r.open_job)
                cmdm.Parameters.AddWithValue("@BOM_type", "MB")
                cmdm.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type  and job = @job order by release_date desc limit 1"
                cmdm.Connection = Login.Connection
                Dim readerm As MySqlDataReader
                readerm = cmdm.ExecuteReader

                If readerm.HasRows Then
                    While readerm.Read
                        MB = readerm(0).ToString
                    End While
                End If

                readerm.Close()
                '----------------------------


                '--- enter data to mr -------
                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                main_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                main_cmd.Parameters.AddWithValue("@created_by", current_user)
                main_cmd.Parameters.AddWithValue("@MB", MB)
                main_cmd.Parameters.AddWithValue("@qty_name", 0)
                main_cmd.Parameters.AddWithValue("@panel_name", panel_n)
                main_cmd.CommandText = "INSERT INTO Revisions.mr_rev(mr_name, rev_name, created_date , created_by, new_panel, MB, panel_name, qty_name) VALUES (@mr_name, @rev_name, now(), @created_by, 'Zero Out', @MB, @panel_name, @qty_name)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()


                '-------- enter data to mr_data ----------

                Dim dimen_table = New DataTable
                dimen_table.Columns.Add("Part_No", GetType(String))
                dimen_table.Columns.Add("Description", GetType(String))
                dimen_table.Columns.Add("Manufacturer", GetType(String))
                dimen_table.Columns.Add("Vendor", GetType(String))
                dimen_table.Columns.Add("Price", GetType(String))
                dimen_table.Columns.Add("Qty", GetType(String))
                dimen_table.Columns.Add("Subtotal", GetType(String))
                dimen_table.Columns.Add("mfg_type", GetType(String))
                dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                dimen_table.Columns.Add("qty_needed", GetType(String))
                dimen_table.Columns.Add("Assembly_name", GetType(String))
                dimen_table.Columns.Add("Part_status", GetType(String))
                dimen_table.Columns.Add("Part_Type", GetType(String))
                dimen_table.Columns.Add("Notes", GetType(String))
                dimen_table.Columns.Add("Need_by_date", GetType(String))


                Dim cmd_a As New MySqlCommand
                cmd_a.Parameters.AddWithValue("@mr_name", mr_name)
                cmd_a.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, Subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, Need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                cmd_a.Connection = Login.Connection

                Dim readera As MySqlDataReader
                readera = cmd_a.ExecuteReader

                If readera.HasRows Then
                    While readera.Read
                        dimen_table.Rows.Add(readera(0).ToString, readera(1).ToString, readera(2).ToString, readera(3).ToString, readera(4).ToString, 0, readera(6).ToString, readera(7).ToString, If(readera(8) Is DBNull.Value, "", readera(8)), If(readera(9) Is DBNull.Value, "", readera(9)), If(readera(10) Is DBNull.Value, "", readera(10)), readera(11).ToString, readera(12).ToString, readera(13).ToString, readera(14).ToString)
                    End While
                End If

                readera.Close()

                For i = 0 To dimen_table.Rows.Count - 1

                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name)
                    Create_cmd6.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                    Create_cmd6.Parameters.AddWithValue("@Part_No", dimen_table.Rows(i).Item(0))
                    Create_cmd6.Parameters.AddWithValue("@description_t", dimen_table.Rows(i).Item(1))
                    Create_cmd6.Parameters.AddWithValue("@ADA_Number", "")
                    Create_cmd6.Parameters.AddWithValue("@Manufacturer", dimen_table.Rows(i).Item(2))
                    Create_cmd6.Parameters.AddWithValue("@Vendor", dimen_table.Rows(i).Item(3))
                    Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(4) Is Nothing, "", dimen_table.Rows(i).Item(4).ToString.Replace("$", "")))
                    Create_cmd6.Parameters.AddWithValue("@mfg_type", dimen_table.Rows(i).Item(7))
                    Create_cmd6.Parameters.AddWithValue("@current_qty", 0)
                    Create_cmd6.Parameters.AddWithValue("@new_qty", 0)
                    Create_cmd6.Parameters.AddWithValue("@delta", 0)
                    Create_cmd6.Parameters.AddWithValue("@assembly", "")
                    Create_cmd6.Parameters.AddWithValue("@Notes", dimen_table.Rows(i).Item(13))
                    Create_cmd6.Parameters.AddWithValue("@part_status", dimen_table.Rows(i).Item(11))
                    Create_cmd6.Parameters.AddWithValue("@part_type", dimen_table.Rows(i).Item(12))
                    Create_cmd6.Parameters.AddWithValue("@isitfull", "")
                    Create_cmd6.Parameters.AddWithValue("@need_date", dimen_table.Rows(i).Item(14))

                    Create_cmd6.CommandText = "INSERT INTO Revisions.mr_rev_data(mr_name, rev_name, Part_No, description_t, ADA_Number, Manufacturer, Vendor, Price, mfg_type, current_qty, new_qty, delta, assembly, Notes, part_status, part_type, isitfull, need_date) VALUES (@mr_name, @rev_name, @Part_No, @description_t, @ADA_Number, @Manufacturer, @Vendor, @Price, @mfg_type, @current_qty, @new_qty, @delta, @assembly, @Notes, @part_status, @part_type, @isitfull, @need_date)"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()


                Next

                Me.Visible = False


            Else
                MessageBox.Show("Revision file already exist!")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub evaluate_change()

        '--------------- change panel qty ------------------

        Dim exist_c As Boolean : exist_c = False

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
            check_cmd.CommandText = "select * from Revisions.mr_rev where mr_name = @mr_name and rev_name = @rev_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                exist_c = True
            End If

            reader.Close()



            If exist_c = False Then

                '  -------- get MBOM ------------
                Dim MB As String : MB = ""

                Dim cmdm As New MySqlCommand
                cmdm.Parameters.AddWithValue("@job", My_Material_r.open_job)
                cmdm.Parameters.AddWithValue("@BOM_type", "MB")
                cmdm.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type  and job = @job order by release_date desc limit 1"
                cmdm.Connection = Login.Connection
                Dim readerm As MySqlDataReader
                readerm = cmdm.ExecuteReader

                If readerm.HasRows Then
                    While readerm.Read
                        MB = readerm(0).ToString
                    End While
                End If

                readerm.Close()
                '----------------------------


                '--- enter data to mr -------
                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                main_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                main_cmd.Parameters.AddWithValue("@created_by", current_user)
                main_cmd.Parameters.AddWithValue("@new_panel", "Change Qty to " & p_qty)
                main_cmd.Parameters.AddWithValue("@MB", MB)
                main_cmd.Parameters.AddWithValue("@qty_name", p_qty)
                main_cmd.Parameters.AddWithValue("@panel_name", panel_n)
                main_cmd.CommandText = "INSERT INTO Revisions.mr_rev(mr_name, rev_name, created_date , created_by, new_panel, MB, qty_name, panel_name) VALUES (@mr_name, @rev_name, now(), @created_by, @new_panel, @MB,  @qty_name, @panel_name)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()


                '-------- enter data to mr_data ----------

                Dim dimen_table = New DataTable
                dimen_table.Columns.Add("Part_No", GetType(String))
                dimen_table.Columns.Add("Description", GetType(String))
                dimen_table.Columns.Add("Manufacturer", GetType(String))
                dimen_table.Columns.Add("Vendor", GetType(String))
                dimen_table.Columns.Add("Price", GetType(String))
                dimen_table.Columns.Add("Qty", GetType(String))
                dimen_table.Columns.Add("Subtotal", GetType(String))
                dimen_table.Columns.Add("mfg_type", GetType(String))
                dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                dimen_table.Columns.Add("qty_needed", GetType(String))
                dimen_table.Columns.Add("Assembly_name", GetType(String))
                dimen_table.Columns.Add("Part_status", GetType(String))
                dimen_table.Columns.Add("Part_Type", GetType(String))
                dimen_table.Columns.Add("Notes", GetType(String))
                dimen_table.Columns.Add("Need_by_date", GetType(String))


                Dim cmd_a As New MySqlCommand
                cmd_a.Parameters.AddWithValue("@mr_name", mr_name)
                cmd_a.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, Subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, Need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                cmd_a.Connection = Login.Connection

                Dim readera As MySqlDataReader
                readera = cmd_a.ExecuteReader

                If readera.HasRows Then
                    While readera.Read
                        dimen_table.Rows.Add(readera(0).ToString, readera(1).ToString, readera(2).ToString, readera(3).ToString, readera(4).ToString, 0, readera(6).ToString, readera(7).ToString, If(readera(8) Is DBNull.Value, "", readera(8)), If(readera(9) Is DBNull.Value, "", readera(9)), If(readera(10) Is DBNull.Value, "", readera(10)), readera(11).ToString, readera(12).ToString, readera(13).ToString, readera(14).ToString)
                    End While
                End If

                readera.Close()

                For i = 0 To dimen_table.Rows.Count - 1

                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name)
                    Create_cmd6.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                    Create_cmd6.Parameters.AddWithValue("@Part_No", dimen_table.Rows(i).Item(0))
                    Create_cmd6.Parameters.AddWithValue("@description_t", dimen_table.Rows(i).Item(1))
                    Create_cmd6.Parameters.AddWithValue("@ADA_Number", "")
                    Create_cmd6.Parameters.AddWithValue("@Manufacturer", dimen_table.Rows(i).Item(2))
                    Create_cmd6.Parameters.AddWithValue("@Vendor", dimen_table.Rows(i).Item(3))
                    Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(4) Is Nothing, "", dimen_table.Rows(i).Item(4).ToString.Replace("$", "")))
                    Create_cmd6.Parameters.AddWithValue("@mfg_type", dimen_table.Rows(i).Item(7))
                    Create_cmd6.Parameters.AddWithValue("@current_qty", dimen_table.Rows(i).Item(5))
                    Create_cmd6.Parameters.AddWithValue("@new_qty", dimen_table.Rows(i).Item(5))
                    Create_cmd6.Parameters.AddWithValue("@delta", 0)
                    Create_cmd6.Parameters.AddWithValue("@assembly", "")
                    Create_cmd6.Parameters.AddWithValue("@Notes", dimen_table.Rows(i).Item(13))
                    Create_cmd6.Parameters.AddWithValue("@part_status", dimen_table.Rows(i).Item(11))
                    Create_cmd6.Parameters.AddWithValue("@part_type", dimen_table.Rows(i).Item(12))
                    Create_cmd6.Parameters.AddWithValue("@isitfull", "")
                    Create_cmd6.Parameters.AddWithValue("@need_date", dimen_table.Rows(i).Item(14))

                    Create_cmd6.CommandText = "INSERT INTO Revisions.mr_rev_data(mr_name, rev_name, Part_No, description_t, ADA_Number, Manufacturer, Vendor, Price, mfg_type, current_qty, new_qty, delta, assembly, Notes, part_status, part_type, isitfull, need_date) VALUES (@mr_name, @rev_name, @Part_No, @description_t, @ADA_Number, @Manufacturer, @Vendor, @Price, @mfg_type, @current_qty, @new_qty, @delta, @assembly, @Notes, @part_status, @part_type, @isitfull, @need_date)"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()


                Next


                Me.Visible = False


            Else
                MessageBox.Show("Revision file already exist!")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class