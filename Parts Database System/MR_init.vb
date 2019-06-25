Imports MySql.Data.MySqlClient

Public Class MR_init

    Public counter_i As Integer


    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.DarkSlateGray
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

    Private Sub MR_init_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        wait_la.Visible = True
        Application.DoEvents()
        'open_grid.Columns(4).Visible = False
        open_grid.Columns(8).Visible = False

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

        open_grid.Columns(8).Visible = True
        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        '  open_grid.Columns(4).Visible = False

        open_grid.Rows.Clear()

        '  Call todo_list()
        ' Call todo_list2()
        wait_la.Visible = False

        'Try
        '    '--- get all MR (latest revision) put them on datatable
        '    Dim dimen_table = New DataTable
        '    dimen_table.Columns.Add("Filename", GetType(String))
        '    dimen_table.Columns.Add("released_date", GetType(String))
        '    dimen_table.Columns.Add("released_by", GetType(String))
        '    dimen_table.Columns.Add("job", GetType(String))
        '    dimen_table.Columns.Add("need_by_date", GetType(String))

        '    ' dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reade
        '    Dim job_list = New List(Of String)()
        '    counter_i = 0


        '    '---------  get jobs ---------
        '    Dim check_cmd2 As New MySqlCommand
        '    check_cmd2.CommandText = "select distinct job from Material_Request.mr where released = 'Y' and finished is null order by job"
        '    check_cmd2.Connection = Login.Connection
        '    check_cmd2.ExecuteNonQuery()

        '    Dim reader2 As MySqlDataReader
        '    reader2 = check_cmd2.ExecuteReader

        '    If reader2.HasRows Then
        '        While reader2.Read
        '            If job_list.Contains(reader2(0).ToString) = False Then
        '                job_list.Add(reader2(0).ToString)
        '            End If
        '        End While
        '    End If

        '    reader2.Close()
        '    '-------------------------------
        '    For i = 0 To job_list.Count - 1

        '        Dim bom_list = New List(Of String)()

        '        '------- get BOM ------
        '        '------ get all BOM of the job 
        '        Dim n_bom As Double : n_bom = 0
        '        Dim check_cmd As New MySqlCommand
        '        check_cmd.Parameters.AddWithValue("@job", job_list.Item(i))
        '        check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

        '        check_cmd.Connection = Login.Connection
        '        check_cmd.ExecuteNonQuery()

        '        Dim reader As MySqlDataReader
        '        reader = check_cmd.ExecuteReader

        '        If reader.HasRows Then
        '            While reader.Read
        '                n_bom = reader(0)
        '            End While
        '        End If

        '        reader.Close()

        '        For j = 1 To n_bom

        '            Dim check_cmd1 As New MySqlCommand
        '            check_cmd1.Parameters.Clear()
        '            check_cmd1.Parameters.AddWithValue("@job", job_list.Item(i))
        '            check_cmd1.Parameters.AddWithValue("@id_bom", j)
        '            check_cmd1.CommandText = "select mr_name, release_date, released_by, job, need_date from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB') order by release_date desc limit 1"

        '            check_cmd1.Connection = Login.Connection
        '            check_cmd1.ExecuteNonQuery()

        '            Dim reader1 As MySqlDataReader
        '            reader1 = check_cmd1.ExecuteReader

        '            If reader1.HasRows Then

        '                While reader1.Read
        '                    dimen_table.Rows.Add(reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString)
        '                    ' counter_i = counter_i + 1
        '                End While
        '            End If

        '            reader1.Close()
        '        Next
        '    Next


        '    '--- once data is in datatable, loop and add to the opengrid only the ones that are not 100% fulfilled
        '    For i = 0 To dimen_table.Rows.Count - 1

        '        '---------- calculate percentage ---------
        '        Dim complete_mr As Double : complete_mr = 0
        '        Dim total_qty As Double : total_qty = 0
        '        Dim fullf As Double : fullf = 0


        '        Dim cmd3 As New MySqlCommand
        '        cmd3.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
        '        cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
        '        cmd3.Connection = Login.Connection
        '        Dim reader3 As MySqlDataReader
        '        reader3 = cmd3.ExecuteReader

        '        If reader3.HasRows Then
        '            While reader3.Read
        '                If IsDBNull(reader3(0)) Then
        '                    fullf = 0
        '                Else
        '                    fullf = CType(reader3(0), Double)
        '                End If
        '            End While
        '        End If

        '        reader3.Close()

        '        Dim cmdx As New MySqlCommand
        '        cmdx.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
        '        cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
        '        cmdx.Connection = Login.Connection
        '        Dim readerx As MySqlDataReader
        '        readerx = cmdx.ExecuteReader

        '        If readerx.HasRows Then
        '            While readerx.Read
        '                If IsDBNull(readerx(0)) Then
        '                    total_qty = 0
        '                Else
        '                    total_qty = CType(readerx(0), Double)
        '                End If
        '            End While
        '        End If

        '        readerx.Close()

        '        If total_qty > 0 Then
        '            complete_mr = Math.Round((fullf / total_qty) * 100)
        '        End If

        '        If complete_mr < 100 Then
        '            open_grid.Rows.Add(New String() {})
        '            open_grid.Rows(counter_i).Cells(1).Value = dimen_table.Rows(i).Item(0)
        '            open_grid.Rows(counter_i).Cells(5).Value = dimen_table.Rows(i).Item(1)
        '            open_grid.Rows(counter_i).Cells(6).Value = dimen_table.Rows(i).Item(2)
        '            open_grid.Rows(counter_i).Cells(7).Value = dimen_table.Rows(i).Item(3)
        '            open_grid.Rows(counter_i).Cells(4).Value = dimen_table.Rows(i).Item(4) 'just added
        '            open_grid.Rows(counter_i).Cells(8).Value = complete_mr & "%"

        '            counter_i = counter_i + 1
        '        End If


        '    Next


        '    counter_i = 0

        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try

        wait_la.Visible = False
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        Cursor.Current = Cursors.WaitCursor

        'MR list
        Label7.ForeColor = Color.WhiteSmoke
        Label4.ForeColor = Color.DarkSeaGreen
        Label3.ForeColor = Color.WhiteSmoke

        job_box.Visible = True
        open_grid.Columns(8).Visible = False
        open_grid.Columns(2).Visible = True
        open_grid.Columns(3).Visible = True


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
                    check_cmd1.CommandText = "select mr_name, Date_Created, created_by, need_date , release_date, released_by, job from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_Type = 'ASM') order by release_date desc limit 1"

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
                            open_grid.Rows(counter_i).Cells(4).Value = Convert.ToDateTime(reader1(3))  'need date delete to string
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
            Cursor.Current = Cursors.Default
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
                check_cmd1.CommandText = "select mr_name,  Date_Created, created_by, need_date , release_date, released_by, job from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_Type = 'ASM') order by release_date desc limit 1"

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
                        open_grid.Rows(i - 1).Cells(4).Value = Convert.ToDateTime(reader1(3))  'need date
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

        Cursor.Current = Cursors.WaitCursor

        'pending MRs
        Label3.ForeColor = Color.WhiteSmoke
        Label4.ForeColor = Color.WhiteSmoke
        Label7.ForeColor = Color.DarkSeaGreen

        job_box.Visible = False
        wait_la.Visible = True
        Application.DoEvents()

        open_grid.Columns(8).Visible = True
        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        ' open_grid.Columns(4).Visible = False

        open_grid.Rows.Clear()

        Try
            '--- get all MR (latest revision) put them on datatable
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Filename", GetType(String))
            dimen_table.Columns.Add("released_date", GetType(String))
            dimen_table.Columns.Add("released_by", GetType(String))
            dimen_table.Columns.Add("job", GetType(String))
            dimen_table.Columns.Add("need_date", GetType(DateTime))

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
                    check_cmd1.CommandText = "select mr_name, release_date, released_by, job , need_date from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            dimen_table.Rows.Add(reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, Convert.ToDateTime(reader1(4)))  'change need date
                            ' counter_i = counter_i + 1
                        End While
                    End If

                    reader1.Close()
                Next
            Next


            '--- once data is in datatable, loop and add to the opengrid only the ones that are not 100% fulfilled
            For i = 0 To dimen_table.Rows.Count - 1

                '---------- calculate percentage ---------
                Dim complete_mr As Double : complete_mr = 0
                Dim total_qty As Double : total_qty = 0
                Dim fullf As Double : fullf = 0


                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
                cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        If IsDBNull(reader3(0)) Then
                            fullf = 0
                        Else
                            fullf = CType(reader3(0), Double)
                        End If
                    End While
                End If

                reader3.Close()

                Dim cmdx As New MySqlCommand
                cmdx.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
                cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
                cmdx.Connection = Login.Connection
                Dim readerx As MySqlDataReader
                readerx = cmdx.ExecuteReader

                If readerx.HasRows Then
                    While readerx.Read
                        If IsDBNull(readerx(0)) Then
                            total_qty = 0
                        Else
                            total_qty = CType(readerx(0), Double)
                        End If
                    End While
                End If

                readerx.Close()

                Dim complete_noround As Double : complete_noround = 0

                If total_qty > 0 Then
                    complete_mr = Math.Floor((fullf / total_qty) * 100)
                    complete_noround = (fullf / total_qty) * 100
                End If

                If complete_noround < 100 Then
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(counter_i).Cells(1).Value = dimen_table.Rows(i).Item(0)
                    open_grid.Rows(counter_i).Cells(5).Value = dimen_table.Rows(i).Item(1)
                    open_grid.Rows(counter_i).Cells(6).Value = dimen_table.Rows(i).Item(2)
                    open_grid.Rows(counter_i).Cells(7).Value = dimen_table.Rows(i).Item(3)
                    open_grid.Rows(counter_i).Cells(4).Value = Convert.ToDateTime(dimen_table.Rows(i).Item(4))
                    open_grid.Rows(counter_i).Cells(8).Value = complete_mr & "%"

                    counter_i = counter_i + 1
                End If


            Next


            counter_i = 0
            Cursor.Current = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        wait_la.Visible = False
    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick

        If open_grid.Rows.Count > 0 Then

            Cursor.Current = Cursors.WaitCursor
            wait_la.Visible = True
            wait_la.Text = "Preparing Material Request ...."
            Application.DoEvents()

            Dim index_k = open_grid.CurrentCell.RowIndex
            Dim name As String = ""


            If open_grid.Rows.Count > 0 Then
                name = If(String.IsNullOrEmpty(open_grid.Rows(index_k).Cells(1).Value) = True, "", open_grid.Rows(index_k).Cells(1).Value)
            End If

            If String.Equals(name, "") = False Then

                Dim ifisASM As Boolean : ifisASM = False

                Try
                    Dim cmdx As New MySqlCommand
                    cmdx.Parameters.AddWithValue("@mr_name", name)
                    cmdx.CommandText = "select BOM_type from Material_Request.mr where mr_name = @mr_name and (BOM_type = 'ASM' or BOM_type = 'old_BOM');"
                    cmdx.Connection = Login.Connection
                    Dim readerx As MySqlDataReader
                    readerx = cmdx.ExecuteReader

                    If readerx.HasRows Then
                        ifisASM = True
                    End If

                    readerx.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                If ifisASM = False Then

                    'if is not ASM BOM then
                    Procurement_Overview.Visible = True
                    Call Procurement_Overview.open_my_job(name)
                    Call Procurement_Overview.Color_Module()
                    Me.Visible = False
                    Procurement_Overview.ASM_mode = False
                    'Procurement_Overview.Visible = True
                    Procurement_Overview.total_grid.Columns(17).DisplayIndex = 17

                Else

                    Procurement_Overview.Visible = True
                    Call Procurement_Overview.open_my_job(name)
                    Call Procurement_Overview.Color_Module()
                    Procurement_Overview.total_grid.Columns(11).Visible = True
                    Me.Visible = False
                    Procurement_Overview.ASM_mode = True
                    Procurement_Overview.total_grid.Columns(17).DisplayIndex = 12

                    'Call Procurement_Overview.open_my_job(name)
                    'Call Procurement_Overview.Color_Module()
                    'Procurement_Overview.Visible = True
                    'Procurement_Overview.total_grid.Columns(11).Visible = True
                    'Procurement_Overview.ASM_mode = True
                    'Me.Visible = False

                End If



            End If

            wait_la.Visible = False
            wait_la.Text = "Loading Data, Please wait ....."
            Cursor.Current = Cursors.Default

        End If
    End Sub

    Private Sub MR_init_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Label7.ForeColor = Color.WhiteSmoke
        Label4.ForeColor = Color.WhiteSmoke

        wait_la.Visible = True
        Application.DoEvents()
        Call todo_list()
        wait_la.Visible = False

    End Sub

    Sub Pending_mr()

        '---------Pending MR display

        wait_la.Visible = True
        Application.DoEvents()
        'open_grid.Columns(4).Visible = False
        open_grid.Columns(8).Visible = False

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

        open_grid.Columns(8).Visible = True
        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        '  open_grid.Columns(4).Visible = False

        open_grid.Rows.Clear()

        Try
            '--- get all MR (latest revision) put them on datatable
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Filename", GetType(String))
            dimen_table.Columns.Add("released_date", GetType(String))
            dimen_table.Columns.Add("released_by", GetType(String))
            dimen_table.Columns.Add("job", GetType(String))
            dimen_table.Columns.Add("need_by_date", GetType(String))

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
                    check_cmd1.CommandText = "select mr_name, release_date, released_by, job, need_date from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_Type = 'ASM') order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            dimen_table.Rows.Add(reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, Convert.ToDateTime(reader1(4)))  'change need date
                            ' counter_i = counter_i + 1
                        End While
                    End If

                    reader1.Close()
                Next
            Next


            '--- once data is in datatable, loop and add to the opengrid only the ones that are not 100% fulfilled
            For i = 0 To dimen_table.Rows.Count - 1

                '---------- calculate percentage ---------
                Dim complete_mr As Double : complete_mr = 0
                Dim total_qty As Double : total_qty = 0
                Dim fullf As Double : fullf = 0


                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
                cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        If IsDBNull(reader3(0)) Then
                            fullf = 0
                        Else
                            fullf = CType(reader3(0), Double)
                        End If
                    End While
                End If

                reader3.Close()

                Dim cmdx As New MySqlCommand
                cmdx.Parameters.AddWithValue("@mr_name", dimen_table.Rows(i).Item(0))
                cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
                cmdx.Connection = Login.Connection
                Dim readerx As MySqlDataReader
                readerx = cmdx.ExecuteReader

                If readerx.HasRows Then
                    While readerx.Read
                        If IsDBNull(readerx(0)) Then
                            total_qty = 0
                        Else
                            total_qty = CType(readerx(0), Double)
                        End If
                    End While
                End If

                readerx.Close()

                If total_qty > 0 Then
                    complete_mr = Math.Round((fullf / total_qty) * 100)
                End If

                If complete_mr < 100 Then
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(counter_i).Cells(1).Value = dimen_table.Rows(i).Item(0)
                    open_grid.Rows(counter_i).Cells(5).Value = dimen_table.Rows(i).Item(1)
                    open_grid.Rows(counter_i).Cells(6).Value = dimen_table.Rows(i).Item(2)
                    open_grid.Rows(counter_i).Cells(7).Value = dimen_table.Rows(i).Item(3)
                    open_grid.Rows(counter_i).Cells(4).Value = dimen_table.Rows(i).Item(4) 'just added
                    open_grid.Rows(counter_i).Cells(8).Value = complete_mr & "%"

                    counter_i = counter_i + 1
                End If


            Next


            counter_i = 0

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        wait_la.Visible = False
    End Sub

    Sub todo_list()
        '-- todo list
        Label3.ForeColor = Color.DarkSeaGreen
        Label4.ForeColor = Color.WhiteSmoke
        Label7.ForeColor = Color.WhiteSmoke

        job_box.Visible = False

        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        open_grid.Columns(8).Visible = False

        open_grid.Rows.Clear()

        Try
            '--- get all MR (latest revision) put them on datatable
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Filename", GetType(String))
            dimen_table.Columns.Add("released_date", GetType(String))
            dimen_table.Columns.Add("released_by", GetType(String))
            dimen_table.Columns.Add("job", GetType(String))
            dimen_table.Columns.Add("need_date", GetType(DateTime))

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
                    check_cmd1.CommandText = "select mr_name, release_date, released_by, job, need_date from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            dimen_table.Rows.Add(reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, Convert.ToDateTime(reader1(4)))
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
                    open_grid.Rows(counter_i).Cells(4).Value = Convert.ToDateTime(dimen_table.Rows(i).Item(4))

                    counter_i = counter_i + 1
                End If


            Next

            counter_i = 0

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Sub todo_list2()
        'improved subroutine

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Filename", GetType(String))
        dimen_table.Columns.Add("released_date", GetType(String))
        dimen_table.Columns.Add("released_by", GetType(String))
        dimen_table.Columns.Add("job", GetType(String))

        Label3.ForeColor = Color.DarkSeaGreen
        Label4.ForeColor = Color.WhiteSmoke
        Label7.ForeColor = Color.WhiteSmoke

        job_box.Visible = False

        open_grid.Columns(2).Visible = False
        open_grid.Columns(3).Visible = False
        open_grid.Columns(8).Visible = False

        open_grid.Rows.Clear()

        Dim mr_list = New List(Of String)()
        counter_i = 0

        Dim check_cmd2 As New MySqlCommand
        check_cmd2.CommandText = "SELECT distinct p1.mr_name from Material_Request.mr as p1 INNER JOIN Material_Request.mr_data ON p1.mr_name = mr_data.mr_name where mr_data.latest_r = 'x' and p1.released = 'Y' and (p1.BOM_type = 'ASM' or p1.BOM_type = 'old_BOM' or p1.BOM_type = 'MB' )"
        check_cmd2.Connection = Login.Connection
        check_cmd2.ExecuteNonQuery()

        Dim reader2 As MySqlDataReader
        reader2 = check_cmd2.ExecuteReader

        If reader2.HasRows Then
            While reader2.Read
                If mr_list.Contains(reader2(0).ToString) = False Then
                    mr_list.Add(reader2(0).ToString)
                End If
            End While
        End If

        reader2.Close()

        '-------- fill table ----
        For i = 0 To mr_list.Count - 1

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.Clear()
            check_cmd1.Parameters.AddWithValue("@mr_name", mr_list.Item(i))
            check_cmd1.CommandText = "select mr_name, release_date, released_by, job from Material_Request.mr where mr_name = @mr_name order by release_date"

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
        '----------------------


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
End Class