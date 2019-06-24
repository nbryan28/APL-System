Imports MySql.Data.MySqlClient

Public Class Open_Revision
    Private Sub Open_Revision_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        open_grid.Rows.Clear()
        Me.Text = "Available saved revision for " & My_Material_r.Text

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
            cmd4.CommandText = "SELECT rev_name, created_date, created_by from Revisions.mr_rev where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(i).Cells(1).Value = reader4(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader4(1).ToString
                    open_grid.Rows(i).Cells(3).Value = reader4(2).ToString
                    i = i + 1
                End While
            End If

            reader4.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        '--- open revision selected

        If My_Material_r.isitreleased = True Then

            If My_Material_r.rev_mode = False Then

                '-- get revision name ----
                If open_grid.Rows.Count > 0 Then

                    Dim index_k = open_grid.CurrentCell.RowIndex

                    Dim rev_name As String : rev_name = open_grid.Rows(index_k).Cells(1).Value 'name of rev

                    Inventory_manage.part_sel = "revision"
                    My_Material_r.rev_mode = True
                    My_Material_r.rev_grid.Rows.Clear()


                    Try
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                        cmd4.Parameters.AddWithValue("@rev_name", rev_name)
                        cmd4.CommandText = "SELECT Part_No, description_t, ADA_Number, Manufacturer, Vendor, Price, mfg_type, current_qty, new_qty, delta, assembly, Notes, part_status, part_Type, isitfull, need_date from Revisions.mr_rev_data where mr_name = @mr_name and rev_name = @rev_name"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then
                            Dim i As Integer : i = 0
                            While reader4.Read
                                My_Material_r.rev_grid.Rows.Add(New String() {})
                                My_Material_r.rev_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(14).Value = reader4(14).ToString
                                My_Material_r.rev_grid.Rows(i).Cells(15).Value = reader4(15).ToString

                                i = i + 1
                            End While

                        End If

                        reader4.Close()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try


                    My_Material_r.TabControl1.TabPages.Insert(2, My_Material_r.TabPage2)
                    My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage2
                    Me.Visible = False


                End If
            Else
                        MessageBox.Show("Close Revision Mode first before opening a Revision file")
            End If
        Else
            MessageBox.Show("This Material Request has not been released")
        End If
    End Sub
End Class