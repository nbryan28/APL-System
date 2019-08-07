Imports MySql.Data.MySqlClient

Public Class Quote_open
    Private Sub Quote_open_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        open_grid.Rows.Clear()

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select quote_name, created_by, Date_Created, quote_probability from quote_table.my_active_quote_table"
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
                    open_grid.Rows(i).Cells(2).Value = reader(2)
                    open_grid.Rows(i).Cells(3).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        Dim quote_n As String : quote_n = ""
        Dim set_list = New List(Of String)()
        ProgressBar1.Visible = True
        Dim index_k = open_grid.CurrentCell.RowIndex

        Dim result As DialogResult = MessageBox.Show("Current Quote changes will be deleted. Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then


            myQuote.go_on = True
            myQuote.change_sol = True

            'remove all ada columns
            For i = myQuote.Panel_grid.Columns.Count - 1 To 1 Step -1
                myQuote.Panel_grid.Columns.RemoveAt(i)
                myQuote.PLC_grid.Columns.RemoveAt(i)
                myQuote.Field_grid.Columns.RemoveAt(i)
                myQuote.Control_grid.Columns.RemoveAt(i)
            Next

            myQuote.counter = 1
            myQuote.my_Set_info.Clear()

            If Allocation_parts.Visible = True Then
                Allocation_parts.alloc_grid.Rows.Clear()
            End If

            '--reset install and quote table

            For i = 1 To myQuote.Install_grid.Rows.Count - 1
                For j = 1 To myQuote.Install_grid.Columns.Count - 1
                    If i <> 33 And i <> 36 And i <> 42 And j <> 2 Then
                        myQuote.Install_grid.Rows(i).Cells(j).Value = 0
                    End If
                Next
            Next

            Call myQuote.Fill_installation()
            '  cal_ins = True
            '----------------------------------------------
            myQuote.Quote_grid.Rows.Clear()
            Call myQuote.load_total_quote()

            quote_n = open_grid.Rows(index_k).Cells(0).value

            Try
                '-------------- 
                '---------------- Global settings ------------------
                Dim cmd6 As New MySqlCommand
                cmd6.Parameters.AddWithValue("@quote", quote_n)
                cmd6.CommandText = "SELECT v_480, v_575, v_230, PF525_c, DE1_c, spare_io, spare_panel_space, A_panel, labor_Cost, include_a, t_feet from quote_table.g_setup where quote_name = @quote"
                cmd6.Connection = Login.Connection
                Dim reader6 As MySqlDataReader
                reader6 = cmd6.ExecuteReader

                If reader6.HasRows Then
                    While reader6.Read

                        If String.Equals(reader6(0).ToString, "1") = True Then
                            myQuote.volts_480.Checked = True
                            myQuote.volts_575.Checked = False
                            myQuote.volts_230.Checked = False

                        ElseIf String.Equals(reader6(1).ToString, "1") = True Then
                            myQuote.volts_575.Checked = True
                            myQuote.volts_480.Checked = False
                            myQuote.volts_230.Checked = False

                        ElseIf String.Equals(reader6(2).ToString, "1") = True Then
                            myQuote.volts_230.Checked = True
                            myQuote.volts_480.Checked = False
                            myQuote.volts_575.Checked = False
                        End If

                        If String.Equals(reader6(3).ToString, "1") = True Then
                            myQuote.PF525_c.Checked = True
                            myQuote.DE1_c.Checked = False

                        ElseIf String.Equals(reader6(4).ToString, "1") = True Then
                            myQuote.DE1_c.Checked = True
                            myQuote.PF525_c.Checked = False
                        End If

                        If String.Equals(reader6(9).ToString, "1") = True Then
                            myQuote.Include_A.Checked = True
                        Else
                            myQuote.Include_A.Checked = False
                        End If

                        myQuote.percentage_io.Text = reader6(5).ToString
                        myQuote.percentage_panel.Text = reader6(6).ToString
                        myQuote.sol_label.Text = reader6(7).ToString  'main control solution label
                        myQuote.labor_box.Text = reader6(8).ToString
                        myQuote.t_feet_box.Text = reader6(10).ToString

                    End While
                End If

                reader6.Close()

                myQuote.apanel_box.Text = myQuote.sol_label.Text

                ProgressBar1.Value = 10

                Dim cmd2 As New MySqlCommand
                cmd2.Parameters.AddWithValue("@quote", quote_n)
                cmd2.CommandText = "SELECT distinct set_name from quote_table.my_inputs where quote_name = @quote"
                cmd2.Connection = Login.Connection
                Dim reader2 As MySqlDataReader
                reader2 = cmd2.ExecuteReader

                If reader2.HasRows Then
                    While reader2.Read
                        set_list.Add(reader2(0))
                    End While
                End If

                reader2.Close()

                '---- add columns ---
                For i = 0 To set_list.Count - 1

                    myQuote.Panel_grid.Columns.Add(set_list.Item(i), set_list.Item(i))
                    myQuote.PLC_grid.Columns.Add(set_list.Item(i), set_list.Item(i))
                    myQuote.Field_grid.Columns.Add(set_list.Item(i), set_list.Item(i))
                    myQuote.Control_grid.Columns.Add(set_list.Item(i), set_list.Item(i))

                    myQuote.my_Set_info.Add(set_list.Item(i), "")
                    myQuote.counter = myQuote.counter + 1

                Next



                For i = 0 To myQuote.Panel_grid.Columns.Count - 1
                    myQuote.Panel_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.PLC_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.Field_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.Control_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next

                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@quote", quote_n)
                cmd3.CommandText = "SELECT feature_desc,qty,type_p,set_name from quote_table.my_inputs where quote_name = @quote"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        '--- logic
                        If String.Equals(reader3(2).ToString, "Panel") = True Then

                            For j = 0 To myQuote.Panel_grid.Rows.Count - 1
                                If String.Equals(reader3(0).ToString, myQuote.Panel_grid.Rows(j).Cells(0).Value) = True Then
                                    For z = 1 To myQuote.Panel_grid.Columns.Count - 1
                                        If String.Equals(reader3(3).ToString, myQuote.Panel_grid.Columns(z).HeaderText) = True Then
                                            myQuote.Panel_grid.Rows(j).Cells(z).Value = reader3(1).ToString
                                            j = myQuote.Panel_grid.Rows.Count
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next


                        ElseIf String.Equals(reader3(2).ToString, "PLC") = True Then

                            For j = 0 To myQuote.PLC_grid.Rows.Count - 1
                                If String.Equals(reader3(0).ToString, myQuote.PLC_grid.Rows(j).Cells(0).Value) = True Then
                                    For z = 1 To myQuote.PLC_grid.Columns.Count - 1
                                        If String.Equals(reader3(3).ToString, myQuote.PLC_grid.Columns(z).HeaderText) = True Then
                                            myQuote.PLC_grid.Rows(j).Cells(z).Value = reader3(1).ToString
                                            j = myQuote.PLC_grid.Rows.Count
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next


                        ElseIf String.Equals(reader3(2).ToString, "Control Panel") = True Then

                            For j = 0 To myQuote.Control_grid.Rows.Count - 1
                                If String.Equals(reader3(0).ToString, myQuote.Control_grid.Rows(j).Cells(0).Value) = True Then
                                    For z = 1 To myQuote.Control_grid.Columns.Count - 1
                                        If String.Equals(reader3(3).ToString, myQuote.Control_grid.Columns(z).HeaderText) = True Then
                                            myQuote.Control_grid.Rows(j).Cells(z).Value = reader3(1).ToString
                                            j = myQuote.Control_grid.Rows.Count
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next

                        ElseIf String.Equals(reader3(2).ToString, "Field") = True Then

                            For j = 0 To myQuote.Field_grid.Rows.Count - 1
                                If String.Equals(reader3(0).ToString, myQuote.Field_grid.Rows(j).Cells(0).Value) = True Then
                                    For z = 1 To myQuote.Field_grid.Columns.Count - 1
                                        If String.Equals(reader3(3).ToString, myQuote.Field_grid.Columns(z).HeaderText) = True Then
                                            myQuote.Field_grid.Rows(j).Cells(z).Value = reader3(1).ToString
                                            j = myQuote.Field_grid.Rows.Count
                                            Exit For
                                        End If
                                    Next
                                End If
                            Next

                        End If

                    End While
                End If

                reader3.Close()

                ProgressBar1.Value = 60

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            myQuote.go_on = False
            myQuote.change_sol = False
            myQuote.recal_enable.Checked = True

            '--------- Enter install values -----------
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("desc", GetType(String))
            dimen_table.Columns.Add("qty", GetType(String))
            dimen_table.Columns.Add("u_time", GetType(String))
            dimen_table.Columns.Add("t_time", GetType(String))
            dimen_table.Columns.Add("rate", GetType(String))
            dimen_table.Columns.Add("rate_t", GetType(String))
            dimen_table.Columns.Add("rate_l", GetType(String))
            dimen_table.Columns.Add("rate_tl", GetType(String))
            dimen_table.Columns.Add("mat_u", GetType(String))
            dimen_table.Columns.Add("mat_t", GetType(String))
            dimen_table.Columns.Add("exp_c", GetType(String))
            dimen_table.Columns.Add("exp_t", GetType(String))
            dimen_table.Columns.Add("total_t", GetType(String))

            Dim cmd31 As New MySqlCommand
            cmd31.Parameters.AddWithValue("@quote_name", quote_n)
            cmd31.CommandText = "SELECT * from quote_table.install_table where quote_name = @quote_name"
            cmd31.Connection = Login.Connection
            Dim reader31 As MySqlDataReader
            reader31 = cmd31.ExecuteReader


            If reader31.HasRows Then
                While reader31.Read
                    dimen_table.Rows.Add(reader31(1), reader31(2), reader31(3), reader31(4), reader31(5), reader31(6), reader31(7), reader31(8), reader31(9), reader31(10), reader31(11), reader31(12), reader31(13))
                End While
            End If

            reader31.Close()

            For i = 0 To dimen_table.Rows.Count - 1
                For j = 0 To myQuote.Install_grid.Rows.Count - 1
                    If String.Equals(dimen_table.Rows(i).Item(0), myQuote.Install_grid.Rows(j).Cells(0).Value) = True Then
                        myQuote.Install_grid.Rows(j).Cells(1).Value = dimen_table.Rows(i).Item(1)
                        myQuote.Install_grid.Rows(j).Cells(2).Value = dimen_table.Rows(i).Item(2)
                        myQuote.Install_grid.Rows(j).Cells(3).Value = dimen_table.Rows(i).Item(3)
                        myQuote.Install_grid.Rows(j).Cells(4).Value = dimen_table.Rows(i).Item(4)
                        myQuote.Install_grid.Rows(j).Cells(5).Value = dimen_table.Rows(i).Item(5)
                        myQuote.Install_grid.Rows(j).Cells(6).Value = dimen_table.Rows(i).Item(6)
                        myQuote.Install_grid.Rows(j).Cells(7).Value = dimen_table.Rows(i).Item(7)
                        myQuote.Install_grid.Rows(j).Cells(8).Value = dimen_table.Rows(i).Item(8)
                        myQuote.Install_grid.Rows(j).Cells(9).Value = dimen_table.Rows(i).Item(9)
                        myQuote.Install_grid.Rows(j).Cells(10).Value = dimen_table.Rows(i).Item(10)
                        myQuote.Install_grid.Rows(j).Cells(11).Value = dimen_table.Rows(i).Item(11)
                        myQuote.Install_grid.Rows(j).Cells(12).Value = dimen_table.Rows(i).Item(12)

                        Exit For
                    End If
                Next
            Next

            ProgressBar1.Value = 75

            '----------------------------------------
            '--- enter totals quote table data
            Dim cmd32 As New MySqlCommand
            cmd32.Parameters.AddWithValue("@quote_name", quote_n)
            cmd32.CommandText = "SELECT * from quote_table.total_q_table where quote_name = @quote_name order by id asc"
            cmd32.Connection = Login.Connection
            Dim reader32 As MySqlDataReader
            reader32 = cmd32.ExecuteReader
            Dim x As Integer : x = 0

            If reader32.HasRows Then
                While reader32.Read
                    If x < myQuote.Quote_grid.Rows.Count Then
                        myQuote.Quote_grid.Rows(x).Cells(0).Value = reader32(1)
                        myQuote.Quote_grid.Rows(x).Cells(1).Value = reader32(2)
                        myQuote.Quote_grid.Rows(x).Cells(2).Value = reader32(3)
                        myQuote.Quote_grid.Rows(x).Cells(3).Value = reader32(4)
                        myQuote.Quote_grid.Rows(x).Cells(4).Value = reader32(5)
                        myQuote.Quote_grid.Rows(x).Cells(5).Value = reader32(6)
                        myQuote.Quote_grid.Rows(x).Cells(6).Value = reader32(7)
                        myQuote.Quote_grid.Rows(x).Cells(7).Value = reader32(8)
                        myQuote.Quote_grid.Rows(x).Cells(8).Value = reader32(9)
                        myQuote.Quote_grid.Rows(x).Cells(9).Value = reader32(10)
                        myQuote.Quote_grid.Rows(x).Cells(10).Value = reader32(11)
                        myQuote.Quote_grid.Rows(x).Cells(11).Value = reader32(12)
                        x = x + 1
                    End If
                End While
            End If

            reader32.Close()

            ProgressBar1.Value = 88

            '---------- enter quote information -------
            Call myQuote.General_calculation()  'Call general calculation after enable it
            myQuote.Text = quote_n

            Call myQuote.totals_bt()  'calculate totals in total excel

        End If

        ProgressBar1.Visible = False
        Me.Visible = False

    End Sub

    Private Sub Quote_open_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class