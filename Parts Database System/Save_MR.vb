Imports MySql.Data.MySqlClient

Public Class Save_MR

    Public index As Integer


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.IsNullOrEmpty(TextBox1.Text) = False And String.Equals(TextBox1.Text, "") = False Then

            Dim exist_c As Boolean : exist_c = False


            '------ check if exist --------

            Select Case Purchase_Request.whatipress

                Case 1
                    '///////////////////// MR save /////////////////////////
                    Try
                        Dim check_cmd As New MySqlCommand
                        check_cmd.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                        check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
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

                            '--- enter data to mr -------
                            Dim main_cmd As New MySqlCommand
                            main_cmd.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                            main_cmd.Parameters.AddWithValue("@created_by", current_user)
                            main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, BOM_type) VALUES (@mr_name, now(), @created_by, 'old_BOM')"
                            main_cmd.Connection = Login.Connection
                            main_cmd.ExecuteNonQuery()


                            '-------- enter data to mr_data
                            For i = 0 To Purchase_Request.PR_grid.Rows.Count - 1

                                If IsNumeric(Purchase_Request.PR_grid.Rows(i).Cells(6).Value) = True And Purchase_Request.PR_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                                    Dim Create_cmd6 As New MySqlCommand
                                    Create_cmd6.Parameters.Clear()
                                    Create_cmd6.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                                    Create_cmd6.Parameters.AddWithValue("@Part_No", If(Purchase_Request.PR_grid.Rows(i).Cells(0).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(0).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@Description", If(Purchase_Request.PR_grid.Rows(i).Cells(1).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(1).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(Purchase_Request.PR_grid.Rows(i).Cells(2).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(2).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(Purchase_Request.PR_grid.Rows(i).Cells(3).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(3).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@Vendor", If(Purchase_Request.PR_grid.Rows(i).Cells(4).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(4).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@Price", If(Purchase_Request.PR_grid.Rows(i).Cells(5).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")))
                                    Create_cmd6.Parameters.AddWithValue("@Qty", If(Purchase_Request.PR_grid.Rows(i).Cells(6).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(6).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@subtotal", If(Purchase_Request.PR_grid.Rows(i).Cells(7).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(7).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@mfg_type", If(Purchase_Request.PR_grid.Rows(i).Cells(8).Value Is Nothing, "", Purchase_Request.PR_grid.Rows(i).Cells(8).Value.ToString))
                                    Create_cmd6.Parameters.AddWithValue("@assembly", "")

                                    Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type , Assembly_name) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type , @assembly)"
                                    Create_cmd6.Connection = Login.Connection
                                    Create_cmd6.ExecuteNonQuery()

                                End If

                            Next

                            'Me.Visible = False


                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                        MessageBox.Show("Done")
                        Me.Visible = False


                    Else
                        MessageBox.Show("MR file already exist!")
                    End If

                Case 3

                    '///////////////////// MR save in My Material Request /////////////////////////

                    If My_Material_r.Special_order_check() = True Then

                        If check_need_date() = False Then


                            Try
                                Dim check_cmd As New MySqlCommand
                                check_cmd.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                                check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
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

                                    '--- enter data to mr -------
                                    Dim main_cmd As New MySqlCommand
                                    main_cmd.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                                    main_cmd.Parameters.AddWithValue("@created_by", current_user)
                                    main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, BOM_type) VALUES (@mr_name, now(), @created_by, 'old_BOM')"
                                    main_cmd.Connection = Login.Connection
                                    main_cmd.ExecuteNonQuery()


                                    '-------- enter data to mr_data
                                    For i = 0 To My_Material_r.PR_grid.Rows.Count - 1

                                        If IsNumeric(My_Material_r.PR_grid.Rows(i).Cells(6).Value) = True And My_Material_r.PR_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                                            Dim Create_cmd6 As New MySqlCommand
                                            Create_cmd6.Parameters.Clear()
                                            Create_cmd6.Parameters.AddWithValue("@mr_name", TextBox1.Text)
                                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(My_Material_r.PR_grid.Rows(i).Cells(0).Value Is Nothing, "unknown part", My_Material_r.PR_grid.Rows(i).Cells(0).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Description", If(My_Material_r.PR_grid.Rows(i).Cells(1).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(1).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(My_Material_r.PR_grid.Rows(i).Cells(2).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(2).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(My_Material_r.PR_grid.Rows(i).Cells(3).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(3).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(My_Material_r.PR_grid.Rows(i).Cells(4).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(4).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Price", If(My_Material_r.PR_grid.Rows(i).Cells(5).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")))
                                            Create_cmd6.Parameters.AddWithValue("@Qty", If(My_Material_r.PR_grid.Rows(i).Cells(6).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(6).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@subtotal", If(My_Material_r.PR_grid.Rows(i).Cells(7).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(7).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(My_Material_r.PR_grid.Rows(i).Cells(8).Value Is Nothing, "Panel", My_Material_r.PR_grid.Rows(i).Cells(8).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@assembly", If(My_Material_r.PR_grid.Rows(i).Cells(9).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(9).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Notes", If(My_Material_r.PR_grid.Rows(i).Cells(10).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(10).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Part_status", If(My_Material_r.PR_grid.Rows(i).Cells(11).Value Is Nothing, "Special Order", My_Material_r.PR_grid.Rows(i).Cells(11).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@Part_type", If(My_Material_r.PR_grid.Rows(i).Cells(12).Value Is Nothing, "Other", My_Material_r.PR_grid.Rows(i).Cells(12).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@full_panel", If(My_Material_r.PR_grid.Rows(i).Cells(13).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(13).Value.ToString))
                                            Create_cmd6.Parameters.AddWithValue("@need_by_date", If(My_Material_r.PR_grid.Rows(i).Cells(14).Value Is Nothing, "", My_Material_r.PR_grid.Rows(i).Cells(14).Value.ToString))


                                            Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Assembly_name, Part_status, Part_type, Notes, full_panel, need_by_date) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @assembly, @Part_status, @Part_type, @Notes, @full_panel, @need_by_date)"
                                            Create_cmd6.Connection = Login.Connection
                                            Create_cmd6.ExecuteNonQuery()

                                        End If

                                    Next

                                    ' Call My_Material_r.Generate_MPL(TextBox1.Text, True)

                                    Call Form1.Command_h(current_user, "Bill Of Material Saved " & TextBox1.Text, "None")

                                    MessageBox.Show("Done")
                                    Me.Visible = False
                                    My_Material_r.PR_grid.Rows.Clear()
                                    My_Material_r.Text = "My Material Requests"


                                Catch ex As Exception
                                    MessageBox.Show(ex.ToString)
                                End Try

                            Else
                                MessageBox.Show("File already exist!")
                            End If

                        Else

                            MessageBox.Show("Please Enter a need by date")
                            Me.Visible = False

                        End If
                    Else
                        MessageBox.Show("Invalid Special Order Parts found!")
                        Me.Visible = False
                    End If
            End Select

        Else
            MessageBox.Show("Please enter a name")

        End If
    End Sub

    Private Sub Save_MR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
    End Sub

    Private Sub Save_MR_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        TextBox1.Text = ""
    End Sub

    Function check_need_date() As Boolean

        check_need_date = False

        For i = 0 To My_Material_r.PR_grid.Rows.Count - 1

            If String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(i).Cells(14).Value) = True And String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(i).Cells(0).Value) = False Then
                check_need_date = True
            End If

        Next

    End Function
End Class