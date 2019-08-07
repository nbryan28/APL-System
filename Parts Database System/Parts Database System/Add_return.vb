Imports MySql.Data.MySqlClient

Public Class Add_return

    Public load_m As Boolean

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Add_return_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null order by job"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    ComboBox1.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        'Try
        '    With DirectCast(DataGridView1.Columns(0), DataGridViewComboBoxColumn)

        '        Dim cmd_v As New MySqlCommand
        '        cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null order by job"

        '        cmd_v.Connection = Login.Connection
        '        Dim readerv As MySqlDataReader
        '        readerv = cmd_v.ExecuteReader

        '        If readerv.HasRows Then
        '            While readerv.Read
        '                .Items.Add(readerv(0))
        '            End While
        '        End If

        '        readerv.Close()

        '    End With

        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try


    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles addr_grid.CellValueChanged
        '-- when value change update BOM name

        'If load_m = True Then

        '    Try
        '        Dim i As Integer : i = DataGridView1.CurrentRow.Index
        '        Dim job As String : job = DataGridView1.Rows(i).Cells(0).Value



        '        Dim dgvcc1 As DataGridViewComboBoxCell
        '        dgvcc1 = DataGridView1.Rows(i).Cells(1)
        '        dgvcc1.Items.Clear()

        '        Dim id As Integer : id = 1
        '        Dim cmd_v As New MySqlCommand
        '        cmd_v.Parameters.AddWithValue("@job", job)
        '        cmd_v.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

        '        cmd_v.Connection = Login.Connection
        '        Dim readerv As MySqlDataReader
        '        readerv = cmd_v.ExecuteReader

        '        If readerv.HasRows Then
        '            While readerv.Read
        '                id = readerv(0)
        '            End While
        '        End If

        '        readerv.Close()

        '        '///////////////////////////

        '        For i = 1 To id

        '            Dim cmd_v2 As New MySqlCommand
        '            cmd_v2.Parameters.AddWithValue("@job", job)
        '            cmd_v2.Parameters.AddWithValue("@id_bom", i)
        '            cmd_v2.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

        '            cmd_v2.Connection = Login.Connection

        '            Dim readerv2 As MySqlDataReader
        '            readerv2 = cmd_v2.ExecuteReader

        '            If readerv2.HasRows Then
        '                While readerv2.Read
        '                    ' dgvcc1 = DataGridView1.Rows(i).Cells(1)
        '                    dgvcc1.Items.Add(readerv2(0))
        '                End While
        '            End If

        '            readerv2.Close()
        '        Next

        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)
        '    End Try

        'End If
    End Sub

    Private Sub DataGridView1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles addr_grid.CellValidating
        'If (e.ColumnIndex = comboBoxColumn.Index) Then

        '    Dim eFV As Object : eFV = e.FormattedValue
        '    Try

        '        If (comboBoxColumn.Items.Contains(eFV) = False) Then

        '            comboBoxColumn.Items.Add(eFV)
        '            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = eFV


        '        End If
        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)
        '    End Try

        'End If
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles addr_grid.DataError
        'If (e.Exception.Message = "DataGridViewComboBoxCell value is not valid.") Then
        '    e.ThrowException = False
        'End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Add_return_sel.ShowDialog()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If addr_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In addr_grid.SelectedRows
                Try
                    addr_grid.Rows.Remove(r)
                Catch ex As Exception
                    MessageBox.Show("This row cannot be deleted")
                End Try
            Next
        Else
            MessageBox.Show("Select the row you want to delete")
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        '--- filter by job

        Dim selected_job As String : selected_job = ""

        If Not ComboBox1.SelectedItem Is Nothing Then
            selected_job = ComboBox1.SelectedItem.ToString
        Else
            selected_job = ""
        End If

        total_grid.Rows.Clear()

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", selected_job)
            cmd4.CommandText = "SELECT job, task, qty, Part_No, Notes, need_by_date, Cost from Tracking_Reports.add_return where job = @job"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    total_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, "", reader4(6).ToString})
                End While

            End If

            reader4.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '-- show all jobs
        total_grid.Rows.Clear()

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT job, task, qty, Part_No, Notes, need_by_date, Cost from Tracking_Reports.add_return"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    total_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, "", reader4(6).ToString})
                End While

            End If

            reader4.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '---- analyze the rows entered
        Try

            Dim goahead As Boolean : goahead = False

            For i = 0 To addr_grid.Rows.Count - 1

                goahead = False

                If String.IsNullOrEmpty(addr_grid.Rows(i).Cells(1).Value) = False And String.IsNullOrEmpty(addr_grid.Rows(i).Cells(3).Value) = False And String.IsNullOrEmpty(addr_grid.Rows(i).Cells(0).Value) = False Then  'no add or return assign, part name and job

                    '--- find project --------
                    Dim cmd44 As New MySqlCommand
                    cmd44.Parameters.AddWithValue("@job", addr_grid.Rows(i).Cells(0).Value)
                    cmd44.CommandText = "SELECT * from management.projects where Job_number = @job"
                    cmd44.Connection = Login.Connection
                    Dim reader44 As MySqlDataReader
                    reader44 = cmd44.ExecuteReader

                    If reader44.HasRows Then
                        While reader44.Read
                            goahead = True
                        End While
                    End If

                    reader44.Close()


                    If goahead = True Then

                        If IsNumeric(addr_grid.Rows(i).Cells(2).Value) = True Then 'if qty is a number
                            If addr_grid.Rows(i).Cells(2).Value > 0 Then

                                Dim main_cmd As New MySqlCommand
                                main_cmd.Parameters.AddWithValue("@job", addr_grid.Rows(i).Cells(0).Value)
                                main_cmd.Parameters.AddWithValue("@task", addr_grid.Rows(i).Cells(1).Value)
                                main_cmd.Parameters.AddWithValue("@qty", addr_grid.Rows(i).Cells(2).Value)
                                main_cmd.Parameters.AddWithValue("@Part_No", addr_grid.Rows(i).Cells(3).Value)
                                main_cmd.Parameters.AddWithValue("@Notes", addr_grid.Rows(i).Cells(4).Value)
                                main_cmd.Parameters.AddWithValue("@need_by_date", addr_grid.Rows(i).Cells(5).Value)
                                main_cmd.Parameters.AddWithValue("@Cost", addr_grid.Rows(i).Cells(6).Value)
                                main_cmd.CommandText = "INSERT INTO Tracking_Reports.add_return(job, task, qty, Part_No, Notes, need_by_date, Cost) VALUES (@job, @task, @qty, @Part_No, @Notes, @need_By_date, @Cost)"
                                main_cmd.Connection = Login.Connection
                                main_cmd.ExecuteNonQuery()



                                '--- if its mark as add
                                If String.Equals(addr_grid.Rows(i).Cells(1).Value, "Add") = True Then

                                    Dim old_inv_qty As Double : old_inv_qty = 0

                                    Dim cmd4 As New MySqlCommand
                                    cmd4.Parameters.AddWithValue("@part_name", addr_grid.Rows(i).Cells(3).Value)
                                    cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                                    cmd4.Connection = Login.Connection
                                    Dim reader4 As MySqlDataReader
                                    reader4 = cmd4.ExecuteReader

                                    If reader4.HasRows Then
                                        While reader4.Read
                                            old_inv_qty = reader4(0).ToString
                                        End While
                                    End If

                                    reader4.Close()

                                    '-------- update ---------
                                    Dim rem_qty As Double : rem_qty = 0

                                    If IsNumeric(addr_grid.Rows(i).Cells(2).Value) = True Then
                                        If addr_grid.Rows(i).Cells(2).Value >= 0 Then
                                            rem_qty = addr_grid.Rows(i).Cells(2).Value
                                        End If

                                    End If

                                    Dim Create_cmd As New MySqlCommand
                                    Create_cmd.Parameters.AddWithValue("@part_name", addr_grid.Rows(i).Cells(3).Value)
                                    Create_cmd.Parameters.AddWithValue("@current_qty", If(old_inv_qty - rem_qty < 0, 0, old_inv_qty - rem_qty))

                                    Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                                    Create_cmd.Connection = Login.Connection
                                    Create_cmd.ExecuteNonQuery()

                                    Call Inventory_manage.refresh_table()

                                ElseIf String.Equals(addr_grid.Rows(i).Cells(1).Value, "Return") = True Then 'else return

                                    Dim old_inv_qty As Double : old_inv_qty = 0

                                    Dim cmd4 As New MySqlCommand
                                    cmd4.Parameters.AddWithValue("@part_name", addr_grid.Rows(i).Cells(3).Value)
                                    cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                                    cmd4.Connection = Login.Connection
                                    Dim reader4 As MySqlDataReader
                                    reader4 = cmd4.ExecuteReader

                                    If reader4.HasRows Then
                                        While reader4.Read
                                            old_inv_qty = reader4(0).ToString
                                        End While
                                    End If

                                    reader4.Close()

                                    '-------- update ---------
                                    Dim rem_qty As Double : rem_qty = 0

                                    If IsNumeric(addr_grid.Rows(i).Cells(2).Value) = True Then
                                        If addr_grid.Rows(i).Cells(2).Value >= 0 Then
                                            rem_qty = addr_grid.Rows(i).Cells(2).Value
                                        End If

                                    End If

                                    Dim Create_cmd As New MySqlCommand
                                    Create_cmd.Parameters.AddWithValue("@part_name", addr_grid.Rows(i).Cells(3).Value)
                                    Create_cmd.Parameters.AddWithValue("@current_qty", old_inv_qty + rem_qty)

                                    Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                                    Create_cmd.Connection = Login.Connection
                                    Create_cmd.ExecuteNonQuery()

                                    Call Inventory_manage.refresh_table()

                                End If

                            End If
                        End If
                    End If
                End If

            Next

            addr_grid.Rows.Clear()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        For i = 0 To total_grid.Rows.Count - 1

            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(3).Value)
            Create_cmd.Parameters.AddWithValue("@job", total_grid.Rows(i).Cells(0).Value)
            Create_cmd.Parameters.AddWithValue("@cost", total_grid.Rows(i).Cells(7).Value)

            Create_cmd.CommandText = "UPDATE Tracking_Reports.add_return SET Cost = @cost where Part_No = @Part_No and job = @job"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()

        Next

        MessageBox.Show("Done")

    End Sub
End Class