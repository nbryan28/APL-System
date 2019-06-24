Imports MySql.Data.MySqlClient

Public Class Enter_data
    Private Sub Enter_data_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = ""

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim quote_n As String
        Dim exist_c As Boolean : exist_c = False

        '--- use in global setup
        Dim v_480 As Double : v_480 = 0
        Dim v_230 As Double : v_230 = 0
        Dim v_575 As Double : v_575 = 0
        Dim PF525_c As Double : PF525_c = 0
        Dim DE1_c As Double : DE1_c = 0
        Dim spare_io As Double
        Dim spare_panel_space As Double
        Dim Apanel As String
        Dim labor_Cost As Double
        Dim include_a As Double : include_a = 0
        Dim t_feet As Double : t_feet = 2000


        If String.IsNullOrEmpty(TextBox1.Text) = False And String.Equals(TextBox1.Text, "") = False Then

            quote_n = TextBox1.Text

            '---- check if quote_n and same job exist
            Try
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@quote_n", quote_n)
                check_cmd.CommandText = "select * from quote_table.my_active_quote_table where quote_name = @quote_n"
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


                '---------- main active_quote table-------------
                Try
                    Dim main_cmd As New MySqlCommand
                    main_cmd.Parameters.AddWithValue("@quote_n", quote_n)
                    main_cmd.Parameters.AddWithValue("@user", current_user)
                    main_cmd.CommandText = "INSERT INTO quote_table.my_active_quote_table(quote_name, created_by, Date_Created) VALUES (@quote_n,@user,now())"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()


                    '------ save global setups --------
                    If myQuote.volts_480.Checked = True Then
                        v_480 = 1
                    ElseIf myQuote.volts_575.Checked = True Then
                        v_575 = 1
                    Else
                        v_230 = 1
                    End If

                    If myQuote.PF525_c.Checked = True Then
                        PF525_c = 1
                    Else
                        DE1_c = 1
                    End If

                    If myQuote.Include_A.Checked = True Then
                        include_a = 1
                    Else
                        include_a = 0
                    End If

                    spare_io = If(IsNumeric(myQuote.percentage_io.Text), myQuote.percentage_io.Text, 0)
                    spare_panel_space = If(IsNumeric(myQuote.percentage_panel.Text), myQuote.percentage_panel.Text, 0)
                    Apanel = myQuote.apanel_box.Text
                    labor_Cost = If(IsNumeric(myQuote.labor_box.Text), myQuote.labor_box.Text, 0)

                    '-- save total feet
                    t_feet = If(IsNumeric(myQuote.t_feet_box.Text), myQuote.t_feet_box.Text, 0)


                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@quote_n", quote_n)
                    Create_cmd.Parameters.AddWithValue("@v_480", CType(v_480, Double))
                    Create_cmd.Parameters.AddWithValue("@v_575", CType(v_575, Double))
                    Create_cmd.Parameters.AddWithValue("@v_230", CType(v_230, Double))
                    Create_cmd.Parameters.AddWithValue("@PF525_c", CType(PF525_c, Double))
                    Create_cmd.Parameters.AddWithValue("@DE1_c", CType(DE1_c, Double))
                    Create_cmd.Parameters.AddWithValue("@spare_io", CType(spare_io, Double))
                    Create_cmd.Parameters.AddWithValue("@spare_panel_space", CType(spare_panel_space, Double))
                    Create_cmd.Parameters.AddWithValue("@apanel", Apanel)
                    Create_cmd.Parameters.AddWithValue("@labor_Cost", CType(labor_Cost, Double))
                    Create_cmd.Parameters.AddWithValue("@include_a", CType(include_a, Double))
                    Create_cmd.Parameters.AddWithValue("@t_feet", CType(t_feet, Double))
                    Create_cmd.CommandText = "INSERT INTO quote_table.g_setup(quote_name, v_480, v_575, v_230, PF525_c, DE1_c, spare_io, spare_panel_space, A_panel ,labor_Cost, include_a, t_feet) VALUES (@quote_n, @v_480,@v_575, @v_230, @PF525_c,@DE1_c, @spare_io,@spare_panel_space, @apanel, @labor_Cost, @include_a, @t_feet)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()


                    '-------------------------- save solutions ------------------------

                    'For i = 0 To myQuote.datatable.Rows.Count - 1

                    '    Dim Create_cmd2 As New MySqlCommand
                    '    Create_cmd2.Parameters.Clear()
                    '    Create_cmd2.Parameters.AddWithValue("@quote_n", quote_n)
                    '    Create_cmd2.Parameters.AddWithValue("@feature_code", myQuote.datatable.Rows(i).Item(0).ToString)
                    '    Create_cmd2.Parameters.AddWithValue("@description", myQuote.datatable.Rows(i).Item(1).ToString)
                    '    Create_cmd2.Parameters.AddWithValue("@solution", myQuote.datatable.Rows(i).Item(2).ToString)
                    '    Create_cmd2.Parameters.AddWithValue("@solution_description", myQuote.datatable.Rows(i).Item(3).ToString)

                    '    Create_cmd2.CommandText = "INSERT INTO quote_table.my_solutions(quote_name,feature_code, description, solution, solution_description) VALUES (@quote_n, @feature_code,@description,@solution,@solution_description)"
                    '    Create_cmd2.Connection = Login.Connection
                    '    Create_cmd2.ExecuteNonQuery()

                    'Next



                    '---------- save inputs ----------

                    '============ Starter Panel =============
                    For i = 1 To myQuote.Panel_grid.Columns.Count - 1
                        For j = 0 To myQuote.Panel_grid.Rows.Count - 1

                            If IsNumeric(myQuote.Panel_grid.Rows(j).Cells(i).Value) = True Then
                                Dim Create_cmd3 As New MySqlCommand
                                Create_cmd3.Parameters.Clear()
                                Create_cmd3.Parameters.AddWithValue("@quote_n", quote_n)
                                Create_cmd3.Parameters.AddWithValue("@set_name", myQuote.Panel_grid.Columns(i).HeaderText.ToString)
                                Create_cmd3.Parameters.AddWithValue("@feature_desc", myQuote.Panel_grid.Rows(j).Cells(0).Value.ToString)
                                Create_cmd3.Parameters.AddWithValue("@qty", CType(myQuote.Panel_grid.Rows(j).Cells(i).Value.ToString, Double))
                                Create_cmd3.Parameters.AddWithValue("@type_p", "Panel")

                                Create_cmd3.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc, qty, type_p) VALUES (@quote_n, @set_name,@feature_desc,@qty,@type_p)"
                                Create_cmd3.Connection = Login.Connection
                                Create_cmd3.ExecuteNonQuery()
                            End If
                        Next
                    Next
                    '================== PLC===============
                    For i = 1 To myQuote.PLC_grid.Columns.Count - 1
                        For j = 0 To myQuote.PLC_grid.Rows.Count - 1

                            If IsNumeric(myQuote.PLC_grid.Rows(j).Cells(i).Value) = True Then
                                Dim Create_cmd2 As New MySqlCommand
                                Create_cmd2.Parameters.Clear()
                                Create_cmd2.Parameters.AddWithValue("@quote_n", quote_n)
                                Create_cmd2.Parameters.AddWithValue("@set_name", myQuote.PLC_grid.Columns(i).HeaderText)
                                Create_cmd2.Parameters.AddWithValue("@feature_desc", myQuote.PLC_grid.Rows(j).Cells(0).Value.ToString)
                                Create_cmd2.Parameters.AddWithValue("@qty", CType(myQuote.PLC_grid.Rows(j).Cells(i).Value.ToString(), Double))
                                Create_cmd2.Parameters.AddWithValue("@type_p", "PLC")

                                Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc,qty, type_p) VALUES (@quote_n,@set_name,@feature_desc,@qty,@type_p)"
                                Create_cmd2.Connection = Login.Connection
                                Create_cmd2.ExecuteNonQuery()
                            End If
                        Next
                    Next

                    '================== Control Panel ===============
                    For i = 1 To myQuote.Control_grid.Columns.Count - 1
                        For j = 0 To myQuote.Control_grid.Rows.Count - 1

                            If IsNumeric(myQuote.Control_grid.Rows(j).Cells(i).Value) = True Then
                                Dim Create_cmd2 As New MySqlCommand
                                Create_cmd2.Parameters.Clear()
                                Create_cmd2.Parameters.AddWithValue("@quote_n", quote_n)
                                Create_cmd2.Parameters.AddWithValue("@set_name", myQuote.Control_grid.Columns(i).HeaderText)
                                Create_cmd2.Parameters.AddWithValue("@feature_desc", myQuote.Control_grid.Rows(j).Cells(0).Value.ToString)
                                Create_cmd2.Parameters.AddWithValue("@qty", CType(myQuote.Control_grid.Rows(j).Cells(i).Value.ToString(), Double))
                                Create_cmd2.Parameters.AddWithValue("@type_p", "Control Panel")

                                Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name, set_name, feature_desc, qty, type_p) VALUES (@quote_n,@set_name,@feature_desc,@qty,@type_p)"
                                Create_cmd2.Connection = Login.Connection
                                Create_cmd2.ExecuteNonQuery()
                            End If
                        Next
                    Next

                    '=================== Field ============
                    For i = 1 To myQuote.Field_grid.Columns.Count - 1
                        For j = 0 To myQuote.Field_grid.Rows.Count - 1

                            If IsNumeric(myQuote.Field_grid.Rows(j).Cells(i).Value) = True Then
                                Dim Create_cmd2 As New MySqlCommand
                                Create_cmd2.Parameters.Clear()
                                Create_cmd2.Parameters.AddWithValue("@quote_n", quote_n)
                                Create_cmd2.Parameters.AddWithValue("@set_name", myQuote.Field_grid.Columns(i).HeaderText)
                                Create_cmd2.Parameters.AddWithValue("@feature_desc", myQuote.Field_grid.Rows(j).Cells(0).Value.ToString)
                                Create_cmd2.Parameters.AddWithValue("@qty", CType(myQuote.Field_grid.Rows(j).Cells(i).Value.ToString, Double))
                                Create_cmd2.Parameters.AddWithValue("@type_p", "Field")

                                Create_cmd2.CommandText = "INSERT INTO quote_table.my_inputs(quote_name,  set_name, feature_desc,qty, type_p) VALUES (@quote_n, @set_name,@feature_desc,@qty,@type_p)"
                                Create_cmd2.Connection = Login.Connection
                                Create_cmd2.ExecuteNonQuery()
                            End If
                        Next
                    Next



                    '---------- save installations -------------
                    For i = 0 To myQuote.Install_grid.Rows.Count - 1

                        Dim Create_cmd2 As New MySqlCommand
                        Create_cmd2.Parameters.Clear()
                        Create_cmd2.Parameters.AddWithValue("@quote_name", quote_n)
                        Create_cmd2.Parameters.AddWithValue("@description", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(0).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(0).Value))
                        Create_cmd2.Parameters.AddWithValue("@qty", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(1).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(1).Value))
                        Create_cmd2.Parameters.AddWithValue("@u_time", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(2).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(2).Value))
                        Create_cmd2.Parameters.AddWithValue("@t_time", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(3).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(3).Value))
                        Create_cmd2.Parameters.AddWithValue("@rate", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(4).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(4).Value))
                        Create_cmd2.Parameters.AddWithValue("@rate_t", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(5).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(5).Value))
                        Create_cmd2.Parameters.AddWithValue("@rate_l", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(6).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(6).Value))
                        Create_cmd2.Parameters.AddWithValue("@rate_tl", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(7).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(7).Value))
                        Create_cmd2.Parameters.AddWithValue("@mat_u", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(8).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(8).Value))
                        Create_cmd2.Parameters.AddWithValue("@mat_t", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(9).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(9).Value))
                        Create_cmd2.Parameters.AddWithValue("@exp_c", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(10).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(10).Value))
                        Create_cmd2.Parameters.AddWithValue("@exp_t", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(11).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(11).Value))
                        Create_cmd2.Parameters.AddWithValue("@total_t", If(String.IsNullOrEmpty(myQuote.Install_grid.Rows(i).Cells(12).Value) = True, "", myQuote.Install_grid.Rows(i).Cells(12).Value))


                        Create_cmd2.CommandText = "INSERT INTO quote_table.install_table(quote_name, description, qty, u_time, t_time, rate, rate_t, rate_l, rate_tl, mat_u, mat_t, exp_c, exp_t, total_t) VALUES (@quote_name, @description, @qty, @u_time, @t_time, @rate, @rate_t, @rate_l, @rate_tl, @mat_u, @mat_t, @exp_c, @exp_t, @total_t)"
                        Create_cmd2.Connection = Login.Connection
                        Create_cmd2.ExecuteNonQuery()

                    Next


                    '---------- save Quote totals --------
                    For i = 0 To myQuote.Quote_grid.Rows.Count - 1

                        Dim Create_cmd2 As New MySqlCommand
                        Create_cmd2.Parameters.Clear()
                        Create_cmd2.Parameters.AddWithValue("@quote_name", quote_n)
                        Create_cmd2.Parameters.AddWithValue("@blank", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(0).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(0).Value))
                        Create_cmd2.Parameters.AddWithValue("@description", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(1).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(1).Value))
                        Create_cmd2.Parameters.AddWithValue("@qty", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(2).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(2).Value))
                        Create_cmd2.Parameters.AddWithValue("@risk", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(3).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(3).Value))
                        Create_cmd2.Parameters.AddWithValue("@qty_w", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(4).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(4).Value))
                        Create_cmd2.Parameters.AddWithValue("@unit_c", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(5).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(5).Value))
                        Create_cmd2.Parameters.AddWithValue("@total_c", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(6).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(6).Value))
                        Create_cmd2.Parameters.AddWithValue("@markup", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(7).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(7).Value))
                        Create_cmd2.Parameters.AddWithValue("@u_price", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(8).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(8).Value))
                        Create_cmd2.Parameters.AddWithValue("@f_price", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(9).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(9).Value))
                        Create_cmd2.Parameters.AddWithValue("@margin", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(10).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(10).Value))
                        Create_cmd2.Parameters.AddWithValue("@type_p", If(String.IsNullOrEmpty(myQuote.Quote_grid.Rows(i).Cells(11).Value) = True, "", myQuote.Quote_grid.Rows(i).Cells(11).Value))
                        Create_cmd2.Parameters.AddWithValue("@id", i + 1)


                        Create_cmd2.CommandText = "INSERT INTO quote_table.total_q_table(quote_name, blank, description, qty, risk, qty_w, unit_c, total_c, markup, u_price, f_price, margin, type_p, id ) VALUES (@quote_name, @blank, @description, @qty, @risk, @qty_w, @unit_c, @total_c, @markup, @u_price, @f_price, @margin, @type_p, @id)"
                        Create_cmd2.Connection = Login.Connection
                        Create_cmd2.ExecuteNonQuery()

                    Next
                    '======================================

                    Me.Visible = False
                    myQuote.Text = quote_n

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Else
                MessageBox.Show("Quote number already exist")
            End If


        Else
            MessageBox.Show("Please enter a Quote number")
        End If

    End Sub
End Class