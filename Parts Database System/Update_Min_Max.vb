
Imports MySql.Data.MySqlClient

Public Class Update_Min_Max

    Private Sub Update_Min_Max_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            part_assembly.Rows.Clear()
            Dim cmd_pd As New MySqlCommand
            cmd_pd.CommandText = "SELECT DEVICE_Name, Description from devices"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read
                    part_assembly.Rows.Add(readerv(0).ToString, readerv(1).ToString)  'add a new row
                End While
            End If

            readerv.Close()

            For i = 0 To part_assembly.Rows.Count - 1

                Dim cmd_pd2 As New MySqlCommand
                cmd_pd2.Parameters.AddWithValue("@part", part_assembly.Rows(i).Cells(0).Value)
                cmd_pd2.CommandText = "SELECT min_qty, max_qty from inventory.inventory_qty where part_name = @part"
                cmd_pd2.Connection = Login.Connection
                Dim readerv2 As MySqlDataReader
                readerv2 = cmd_pd2.ExecuteReader

                If readerv2.HasRows Then
                    While readerv2.Read
                        part_assembly.Rows(i).Cells(2).Value = readerv2(0).ToString
                        part_assembly.Rows(i).Cells(3).Value = readerv2(1).ToString
                        part_assembly.Rows(i).Cells(4).Value = readerv2(0).ToString
                        part_assembly.Rows(i).Cells(5).Value = readerv2(1).ToString
                    End While
                End If

                readerv2.Close()
            Next



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub part_assembly_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles part_assembly.CellContentClick
        comp_grid.Rows.Clear()

        If part_assembly.Rows.Count > 0 Then

            Dim index_k = part_assembly.CurrentCell.RowIndex


            Try
                Dim cmd_pd As New MySqlCommand
                cmd_pd.Parameters.AddWithValue("@lega", part_assembly.Rows(index_k).Cells(0).Value)
                cmd_pd.CommandText = "SELECT Part_Name from adv where Legacy_ADA = @lega"
                cmd_pd.Connection = Login.Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_pd.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        comp_grid.Rows.Add(readerv(0).ToString)  'add a new row
                    End While
                End If

                readerv.Close()

                '--------------------------------------
                For i = 0 To comp_grid.Rows.Count - 1

                    Dim cmd_pd2 As New MySqlCommand
                    cmd_pd2.Parameters.AddWithValue("@part", comp_grid.Rows(i).Cells(0).Value)
                    cmd_pd2.CommandText = "SELECT min_qty, max_qty from inventory.inventory_qty where part_name = @part"
                    cmd_pd2.Connection = Login.Connection
                    Dim readerv2 As MySqlDataReader
                    readerv2 = cmd_pd2.ExecuteReader

                    If readerv2.HasRows Then
                        While readerv2.Read
                            comp_grid.Rows(i).Cells(1).Value = readerv2(0).ToString
                            comp_grid.Rows(i).Cells(2).Value = readerv2(1).ToString
                        End While
                    End If

                    readerv2.Close()
                Next
                '--------------------------------------

                Me.Text = part_assembly.Rows(index_k).Cells(0).Value

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Update_Min_Max_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '---------- update assemblies -------------

        Try

            Label4.Visible = True
            Application.DoEvents()

            '----- clear all min and max of all the assemblies --
            For i = 0 To part_assembly.Rows.Count - 1
                Call zero_min_max(part_assembly.Rows(i).Cells(0).Value)
            Next

            '/////////////////////////////////////////////////


            For i = 0 To part_assembly.Rows.Count - 1

                Dim DataTable = New DataTable 'this table will contain parts of assembly
                DataTable.Columns.Add("part", GetType(String))
                DataTable.Columns.Add("qty", GetType(String))
                DataTable.Columns.Add("min", GetType(String))
                DataTable.Columns.Add("max", GetType(String))

                If IsNumeric(part_assembly.Rows(i).Cells(4).Value) = True And IsNumeric(part_assembly.Rows(i).Cells(5).Value) = True Then

                    ' If part_assembly.Rows(i).Cells(4).Value <= part_assembly.Rows(i).Cells(5).Value Then


                    '----- update max and min -------
                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.Clear()
                    Create_cmd.Parameters.AddWithValue("@part_name", part_assembly.Rows(i).Cells(0).Value)
                    Create_cmd.Parameters.AddWithValue("@min_qty", If(part_assembly.Rows(i).Cells(4).Value < 0, 0, part_assembly.Rows(i).Cells(4).Value))
                    Create_cmd.Parameters.AddWithValue("@max_qty", If(part_assembly.Rows(i).Cells(5).Value < 0, 0, part_assembly.Rows(i).Cells(5).Value))

                    Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET min_qty = @min_qty, max_qty = @max_qty where part_name = @part_name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    '-----//// enter members of the specified assembly in datatable///////// ---------

                    Dim cmd_pd As New MySqlCommand
                    cmd_pd.Parameters.Clear()
                    cmd_pd.Parameters.AddWithValue("@lega", part_assembly.Rows(i).Cells(0).Value)
                    cmd_pd.CommandText = "SELECT Part_Name, Qty from adv where Legacy_ADA = @lega"
                    cmd_pd.Connection = Login.Connection
                    Dim readerv As MySqlDataReader
                    readerv = cmd_pd.ExecuteReader

                    If readerv.HasRows Then
                        While readerv.Read
                            DataTable.Rows.Add(readerv(0).ToString, readerv(1).ToString, 0, 0) 'add a new row
                        End While
                    End If

                    readerv.Close()

                    '------  add min and max as well -------
                    For z = 0 To DataTable.Rows.Count - 1
                        Dim cmd_pd3 As New MySqlCommand
                        cmd_pd3.Parameters.Clear()
                        cmd_pd3.Parameters.AddWithValue("@part", DataTable.Rows(z).Item(0).ToString)
                        cmd_pd3.CommandText = "SELECT min_qty, max_qty from inventory.inventory_qty where part_name = @part"
                        cmd_pd3.Connection = Login.Connection
                        Dim readerv3 As MySqlDataReader
                        readerv3 = cmd_pd3.ExecuteReader

                        If readerv3.HasRows Then
                            While readerv3.Read
                                If readerv3.IsDBNull(0) = False Then
                                    DataTable.Rows(z).Item(2) = readerv3(0)
                                End If

                                If readerv3.IsDBNull(1) = False Then
                                    DataTable.Rows(z).Item(3) = readerv3(1)
                                End If
                            End While
                        End If

                        readerv3.Close()
                    Next

                    '/////////////////////////////---------------------------------////////////////////////

                    '---- adjust min and max of members ----------
                    For j = 0 To DataTable.Rows.Count - 1

                        Dim cmd_pd2 As New MySqlCommand
                        cmd_pd2.Parameters.Clear()
                        cmd_pd2.Parameters.AddWithValue("@part", DataTable.Rows(j).Item(0).ToString)
                        cmd_pd2.Parameters.AddWithValue("@min", If(IsNumeric(DataTable.Rows(j).Item(2)), DataTable.Rows(j).Item(2), 0) + If(IsNumeric(DataTable.Rows(j).Item(1)), DataTable.Rows(j).Item(1), 0) * If(part_assembly.Rows(i).Cells(4).Value < 0, 0, part_assembly.Rows(i).Cells(4).Value))
                        cmd_pd2.Parameters.AddWithValue("@max", If(IsNumeric(DataTable.Rows(j).Item(3)), DataTable.Rows(j).Item(3), 0) + If(IsNumeric(DataTable.Rows(j).Item(1)), DataTable.Rows(j).Item(1), 0) * If(part_assembly.Rows(i).Cells(5).Value < 0, 0, part_assembly.Rows(i).Cells(5).Value))
                        cmd_pd2.CommandText = "UPDATE inventory.inventory_qty SET min_qty = @min, max_qty = @max where part_name = @part"
                        cmd_pd2.Connection = Login.Connection
                        cmd_pd2.ExecuteNonQuery()

                    Next


                    '-----------------------------------------------------------------


                    ' End If

                End If

            Next

            Call Inventory_manage.General_inv_cal()  'recal inventory values
            Call reload_tables()
            Label4.Visible = False
            MessageBox.Show("Done")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Sub reload_tables()
        Try
            part_assembly.Rows.Clear()
            Dim cmd_pd As New MySqlCommand
            cmd_pd.CommandText = "SELECT DEVICE_Name, Description from devices"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read
                    part_assembly.Rows.Add(readerv(0).ToString, readerv(1).ToString)  'add a new row
                End While
            End If

            readerv.Close()

            For i = 0 To part_assembly.Rows.Count - 1

                Dim cmd_pd2 As New MySqlCommand
                cmd_pd2.Parameters.AddWithValue("@part", part_assembly.Rows(i).Cells(0).Value)
                cmd_pd2.CommandText = "SELECT min_qty, max_qty from inventory.inventory_qty where part_name = @part"
                cmd_pd2.Connection = Login.Connection
                Dim readerv2 As MySqlDataReader
                readerv2 = cmd_pd2.ExecuteReader

                If readerv2.HasRows Then
                    While readerv2.Read
                        part_assembly.Rows(i).Cells(2).Value = readerv2(0).ToString
                        part_assembly.Rows(i).Cells(3).Value = readerv2(1).ToString
                        part_assembly.Rows(i).Cells(4).Value = readerv2(0).ToString
                        part_assembly.Rows(i).Cells(5).Value = readerv2(1).ToString
                    End While
                End If

                readerv2.Close()
            Next



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub zero_min_max(assem As String)

        '-zero out min and maxs

        Dim DataTable = New DataTable 'this table will contain parts of assembly

        DataTable.Columns.Add("part", GetType(String))

        Dim cmd_pd As New MySqlCommand
        cmd_pd.Parameters.Clear()
        cmd_pd.Parameters.AddWithValue("@lega", assem)
        cmd_pd.CommandText = "SELECT Part_Name from adv where Legacy_ADA = @lega"
        cmd_pd.Connection = Login.Connection
        Dim readerv As MySqlDataReader
        readerv = cmd_pd.ExecuteReader

        If readerv.HasRows Then
            While readerv.Read
                DataTable.Rows.Add(readerv(0).ToString) 'add a new row
            End While
        End If

        readerv.Close()


        '/////////////////////////////---------------------------------////////////////////////

        '---- adjust min and max of members ----------
        For j = 0 To DataTable.Rows.Count - 1

            Dim cmd_pd2 As New MySqlCommand
            cmd_pd2.Parameters.Clear()
            cmd_pd2.Parameters.AddWithValue("@part", DataTable.Rows(j).Item(0).ToString)
            cmd_pd2.CommandText = "UPDATE inventory.inventory_qty SET min_qty = 0, max_qty = 0 where part_name = @part"
            cmd_pd2.Connection = Login.Connection
            cmd_pd2.ExecuteNonQuery()

        Next

    End Sub

End Class