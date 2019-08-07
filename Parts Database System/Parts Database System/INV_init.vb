Imports MySql.Data.MySqlClient

Public Class INV_init

    Public counter_i As Integer


    Private Sub INV_init_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        open_grid.Columns(4).Visible = False
        wait_la.Visible = True
        Application.DoEvents()

        Try
            job_box.Items.Clear()
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    job_box.Items.Add(reader2(0))
                End While
            End If

            reader2.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Call todo_list()
        wait_la.Visible = False

    End Sub

    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = Color.Sienna
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.SaddleBrown
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        'home
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        open_grid.Columns(2).Visible = True
        open_grid.Columns(3).Visible = True
        open_grid.Columns(4).Visible = True

        Label7.ForeColor = Color.WhiteSmoke
        Label4.ForeColor = Color.Orange
        job_box.Visible = True


        open_grid.Rows.Clear()

        Dim job_list = New List(Of String)()
        counter_i = 0

        Try
            '---------  get jobs ---------
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    If job_list.Contains(reader2(0).ToString) = False Then
                        job_list.Add(reader2(0).ToString)
                    End If
                End While
            End If

            reader2.Close()
            '-------------------------------
            For i = 0 To job_list.Count - 1

                Dim bom_list = New List(Of String)()

                '------- get BOM ------
                '------ get all BOM of the job 
                Dim n_bom As Double : n_bom = 0
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", job_list.Item(i))
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

                For j = 1 To n_bom

                    Dim check_cmd1 As New MySqlCommand
                    check_cmd1.Parameters.Clear()
                    check_cmd1.Parameters.AddWithValue("@job", job_list.Item(i))
                    check_cmd1.Parameters.AddWithValue("@id_bom", j)
                    check_cmd1.CommandText = "select mr_name, Date_Created, created_by, last_modified , release_date, released_by, job from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            open_grid.Rows.Add(New String() {})
                            open_grid.Rows(counter_i).Cells(1).Value = reader1(0).ToString
                            open_grid.Rows(counter_i).Cells(2).Value = reader1(1).ToString
                            open_grid.Rows(counter_i).Cells(3).Value = reader1(2).ToString
                            open_grid.Rows(counter_i).Cells(4).Value = reader1(3).ToString
                            open_grid.Rows(counter_i).Cells(5).Value = reader1(4).ToString
                            open_grid.Rows(counter_i).Cells(6).Value = reader1(5).ToString
                            open_grid.Rows(counter_i).Cells(7).Value = reader1(6).ToString

                            counter_i = counter_i + 1
                        End While
                    End If

                    reader1.Close()
                Next
            Next


            counter_i = 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub job_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles job_box.SelectedValueChanged
        open_grid.Rows.Clear()

        Try

            '------ get all BOM of the job Note: unelegant way of do this. Try to find a better Query
            Dim n_bom As Double : n_bom = 0
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", job_box.SelectedItem.ToString)
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
                check_cmd1.Parameters.AddWithValue("@job", job_box.SelectedItem.ToString)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select mr_name,  Date_Created, created_by, last_modified , release_date, released_by, job from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    ' Dim j As Integer : j = 0


                    While reader1.Read
                        open_grid.Rows.Add(New String() {})
                        open_grid.Rows(i - 1).Cells(1).Value = reader1(0).ToString
                        open_grid.Rows(i - 1).Cells(2).Value = reader1(1).ToString
                        open_grid.Rows(i - 1).Cells(3).Value = reader1(2).ToString
                        open_grid.Rows(i - 1).Cells(4).Value = reader1(3).ToString
                        open_grid.Rows(i - 1).Cells(5).Value = reader1(4).ToString
                        open_grid.Rows(i - 1).Cells(6).Value = reader1(5).ToString
                        open_grid.Rows(i - 1).Cells(7).Value = reader1(6).ToString

                    End While
                End If

                reader1.Close()
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        wait_la.Visible = True
        Application.DoEvents()
        Call todo_list()
        wait_la.Visible = False


    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick

        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex

            Dim name As String = ""


            If open_grid.Rows.Count > 0 Then
                name = If(String.IsNullOrEmpty(open_grid.Rows(index_k).Cells(1).Value) = True, "", open_grid.Rows(index_k).Cells(1).Value)
            End If

            If String.Equals(name, "") = False Then
                '----------- Open a release IR ----------
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
                    cmd4.Parameters.AddWithValue("@mr_name", Name)
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

                            Inv_Material_r.fullfill_grid.Rows(i).Cells(12).Value = "No"
                        End If

                        reader5.Close()

                    Next
                    '----------------------------------------------------------------

                    Inv_Material_r.Text = name
                    Inv_Material_r.open_job = open_grid.Rows(index_k).Cells(7).Value
                    Inv_Material_r.job_label.Text = open_grid.Rows(index_k).Cells(7).Value

                    Me.Visible = False
                    Inv_Material_r.Visible = True
                    Inv_Material_r.fullfill_grid.Rows(0).Cells(9).Value = 1
                    Inv_Material_r.fullfill_grid.Rows(0).Cells(9).Value = 0

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If

        End If

    End Sub

    Function action_take_inv(mr_name As String) As Boolean
        action_take_inv = False

        '--- this function return true if there is a part to be fullfill and there are enough parts in current inv
        Try
            '--- get all parts in datatable

            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Part_No", GetType(String))
            dimen_table.Columns.Add("qty", GetType(String))
            dimen_table.Columns.Add("fulfilled", GetType(String))

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT Part_No, Qty, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString)
                End While
            End If

            reader4.Close()

            '-------------------------------------------
            For i = 0 To dimen_table.Rows.Count - 1

                Dim qty As Double : qty = If(IsNumeric(dimen_table.Rows(i).Item(1)) = True, dimen_table.Rows(i).Item(1), 0)
                Dim full_qty As Double : full_qty = If(IsNumeric(dimen_table.Rows(i).Item(2)) = True, dimen_table.Rows(i).Item(2), 0)
                Dim current_inv As Double : current_inv = 0

                '--- get current qty
                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0))
                cmd5.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        current_inv = reader5(0).ToString
                    End While
                End If

                If qty - full_qty > 0 Then
                    If (qty - full_qty) <= current_inv Then
                        action_take_inv = True
                        i = dimen_table.Rows.Count + 3
                    End If
                End If


                reader5.Close()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    Sub todo_list()
        '-- todo list
        Label7.ForeColor = Color.Orange
        Label4.ForeColor = Color.WhiteSmoke
        job_box.Visible = False


        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        open_grid.Columns(4).Visible = False

        open_grid.Rows.Clear()

        Try
            '--- get all MR (latest revision) put them on datatable
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Filename", GetType(String))
            dimen_table.Columns.Add("released_date", GetType(String))
            dimen_table.Columns.Add("released_by", GetType(String))
            dimen_table.Columns.Add("job", GetType(String))

            ' dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reade
            Dim job_list = New List(Of String)()
            counter_i = 0


            '---------  get jobs ---------
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr where released = 'Y' and finished is null order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    If job_list.Contains(reader2(0).ToString) = False Then
                        job_list.Add(reader2(0).ToString)
                    End If
                End While
            End If

            reader2.Close()
            '-------------------------------
            For i = 0 To job_list.Count - 1

                Dim bom_list = New List(Of String)()

                '------- get BOM ------
                '------ get all BOM of the job 
                Dim n_bom As Double : n_bom = 0
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", job_list.Item(i))
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

                For j = 1 To n_bom

                    Dim check_cmd1 As New MySqlCommand
                    check_cmd1.Parameters.Clear()
                    check_cmd1.Parameters.AddWithValue("@job", job_list.Item(i))
                    check_cmd1.Parameters.AddWithValue("@id_bom", j)
                    check_cmd1.CommandText = "select mr_name, release_date, released_by, job from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            dimen_table.Rows.Add(reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString)
                        End While
                    End If

                    reader1.Close()
                Next
            Next

            '--- see which mr need to be taken action

            For i = 0 To dimen_table.Rows.Count - 1
                If action_take_inv(dimen_table.Rows(i).Item(0)) = True Then
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(counter_i).Cells(1).Value = dimen_table.Rows(i).Item(0)
                    open_grid.Rows(counter_i).Cells(5).Value = dimen_table.Rows(i).Item(1)
                    open_grid.Rows(counter_i).Cells(6).Value = dimen_table.Rows(i).Item(2)
                    open_grid.Rows(counter_i).Cells(7).Value = dimen_table.Rows(i).Item(3)

                    counter_i = counter_i + 1
                End If


            Next

            counter_i = 0

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub INV_init_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class