Imports MySql.Data.MySqlClient

Public Class SAVE_revision

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- save data in revision table
        Dim exist_c As Boolean : exist_c = False

        If String.IsNullOrEmpty(TextBox1.Text) = False And String.Equals("", TextBox1.Text) = False Then

            Try
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
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
                    cmdm.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                    cmdm.CommandText = "SELECT MBOM from Material_Request.mr where mr_name = @mr_name  and job = @job"
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
                    main_cmd.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                    main_cmd.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                    main_cmd.Parameters.AddWithValue("@created_by", current_user)
                    main_cmd.Parameters.AddWithValue("@MB", MB)
                    main_cmd.CommandText = "INSERT INTO Revisions.mr_rev(mr_name, rev_name, created_date , created_by, new_panel, MB) VALUES (@mr_name, @rev_name, now(), @created_by, '', @MB)"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()


                    '-------- enter data to mr_data
                    For i = 0 To My_Material_r.rev_grid.Rows.Count - 1

                        If My_Material_r.rev_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                            Dim Create_cmd6 As New MySqlCommand
                            Create_cmd6.Parameters.Clear()
                            Create_cmd6.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                            Create_cmd6.Parameters.AddWithValue("@rev_name", TextBox1.Text)
                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(My_Material_r.rev_grid.Rows(i).Cells(0).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(0).Value))
                            Create_cmd6.Parameters.AddWithValue("@description_t", If(My_Material_r.rev_grid.Rows(i).Cells(1).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(1).Value))
                            Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(My_Material_r.rev_grid.Rows(i).Cells(2).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(2).Value))
                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(My_Material_r.rev_grid.Rows(i).Cells(3).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(3).Value))
                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(My_Material_r.rev_grid.Rows(i).Cells(4).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(4).Value))
                            Create_cmd6.Parameters.AddWithValue("@Price", If(My_Material_r.rev_grid.Rows(i).Cells(5).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")))
                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(My_Material_r.rev_grid.Rows(i).Cells(6).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(6).Value))
                            Create_cmd6.Parameters.AddWithValue("@current_qty", If(My_Material_r.rev_grid.Rows(i).Cells(7).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(7).Value))
                            Create_cmd6.Parameters.AddWithValue("@new_qty", If(My_Material_r.rev_grid.Rows(i).Cells(8).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(8).Value))
                            Create_cmd6.Parameters.AddWithValue("@delta", If(My_Material_r.rev_grid.Rows(i).Cells(9).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(9).Value))
                            Create_cmd6.Parameters.AddWithValue("@assembly", If(My_Material_r.rev_grid.Rows(i).Cells(10).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(10).Value))
                            Create_cmd6.Parameters.AddWithValue("@Notes", If(My_Material_r.rev_grid.Rows(i).Cells(11).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(11).Value))
                            Create_cmd6.Parameters.AddWithValue("@part_status", If(My_Material_r.rev_grid.Rows(i).Cells(12).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(12).Value))
                            Create_cmd6.Parameters.AddWithValue("@part_type", If(My_Material_r.rev_grid.Rows(i).Cells(13).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(13).Value))
                            Create_cmd6.Parameters.AddWithValue("@isitfull", If(My_Material_r.rev_grid.Rows(i).Cells(14).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(14).Value))
                            Create_cmd6.Parameters.AddWithValue("@need_date", If(My_Material_r.rev_grid.Rows(i).Cells(15).Value Is Nothing, "", My_Material_r.rev_grid.Rows(i).Cells(15).Value))

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
End Class