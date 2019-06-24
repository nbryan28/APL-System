Imports MySql.Data.MySqlClient


Public Class BOM_init

    Public counter_i As Integer
    Public mr_mode_r As Boolean


    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.FromArgb(61, 60, 78)
    End Sub

    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter
        Label3.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave
        Label3.BackColor = Color.FromArgb(61, 60, 78)
    End Sub



    Private Sub BOM_init_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        mr_mode_r = False
        open_grid.Columns(7).Visible = False

        Try

            Dim bmp1 As New System.Drawing.Bitmap(My.Resources.myImages.ada_package)
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select mr_name, Date_Created, created_by, last_modified, BOM_type from Material_Request.mr  where released is null and (BOM_type = 'old_BOM' or BOM_Type = 'MB') order by mr_name"
            '  check_cmd.CommandText = "select mr_name, Date_Created, created_by, last_modified from Material_Request.mr  where released is null order by mr_name "
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    open_grid.Rows.Add(New String() {})

                    '-- select right image --
                    If String.Equals(reader(4).ToString, "MB") = True Then
                        open_grid.Rows(i).Cells(0).Value = bmp1
                    End If

                    open_grid.Rows(i).Cells(1).Value = reader(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader(1)
                    open_grid.Rows(i).Cells(3).Value = reader(2).ToString
                    open_grid.Rows(i).Cells(4).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()

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
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

        mr_mode_r = False
        Label7.ForeColor = Color.CadetBlue
        Label4.ForeColor = Color.WhiteSmoke
        job_box.Visible = False
        open_grid.Columns(7).Visible = False
        open_grid.Rows.Clear()

        Try
            Dim bmp1 As New System.Drawing.Bitmap(My.Resources.myImages.ada_package)
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select mr_name, Date_Created, created_by, last_modified, BOM_type from Material_Request.mr  where released is null and (BOM_type = 'old_BOM' or BOM_Type = 'MB' ) order by mr_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    open_grid.Rows.Add(New String() {})

                    If String.Equals(reader(4).ToString, "MB") = True Then
                        open_grid.Rows(i).Cells(0).Value = bmp1
                    End If

                    open_grid.Rows(i).Cells(1).Value = reader(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader(1)
                    open_grid.Rows(i).Cells(3).Value = reader(2).ToString
                    open_grid.Rows(i).Cells(4).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        'released

        mr_mode_r = True
        Label7.ForeColor = Color.WhiteSmoke
        Label4.ForeColor = Color.CadetBlue
        job_box.Visible = True
        open_grid.Columns(7).Visible = True

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

                    Dim bmp1 As New System.Drawing.Bitmap(My.Resources.myImages.ada_package)
                    Dim bmp2 As New System.Drawing.Bitmap(My.Resources.abom)

                    check_cmd1.Parameters.Clear()
                    check_cmd1.Parameters.AddWithValue("@job", job_list.Item(i))
                    check_cmd1.Parameters.AddWithValue("@id_bom", j)
                    check_cmd1.CommandText = "select mr_name, Date_Created, created_by, last_modified , release_date, released_by, job, BOM_type from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            open_grid.Rows.Add(New String() {})

                            If String.Equals(reader1(7).ToString, "MB") = True Then
                                open_grid.Rows(counter_i).Cells(0).Value = bmp1
                            ElseIf String.Equals(reader1(7).ToString, "ASM") = True Then
                                open_grid.Rows(counter_i).Cells(0).Value = bmp2
                            End If
                            open_grid.Rows(counter_i).Cells(1).Value = reader1(0).ToString
                            open_grid.Rows(counter_i).Cells(2).Value = reader1(1)
                            open_grid.Rows(counter_i).Cells(3).Value = reader1(2).ToString
                            open_grid.Rows(counter_i).Cells(4).Value = reader1(3).ToString
                            open_grid.Rows(counter_i).Cells(5).Value = reader1(4)
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
            Dim bmp1 As New System.Drawing.Bitmap(My.Resources.myImages.ada_package)
            Dim bmp2 As New System.Drawing.Bitmap(My.Resources.abom)

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
                check_cmd1.CommandText = "select mr_name,  Date_Created, created_by, last_modified , release_date, released_by, job, BOM_type from Material_Request.mr where job = @job and id_bom = @id_bom and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    ' Dim j As Integer : j = 0


                    While reader1.Read
                        open_grid.Rows.Add(New String() {})

                        If String.Equals(reader1(7).ToString, "MB") = True Then
                            open_grid.Rows(i - 1).Cells(0).Value = bmp1
                        ElseIf String.Equals(reader1(7).ToString, "ASM") = True Then
                            open_grid.Rows(i - 1).Cells(0).Value = bmp2
                        End If

                        open_grid.Rows(i - 1).Cells(1).Value = reader1(0).ToString
                        open_grid.Rows(i - 1).Cells(2).Value = reader1(1)
                        open_grid.Rows(i - 1).Cells(3).Value = reader1(2).ToString
                        open_grid.Rows(i - 1).Cells(4).Value = reader1(3).ToString
                        open_grid.Rows(i - 1).Cells(5).Value = reader1(4)
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

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        'home
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        '--- new mr form

        BOM_types.Visible = True
        Me.Visible = False
    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        '------------ select MR and open it

        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex

            If index_k >= 0 And index_k < open_grid.Rows.Count Then

                Dim name As String = ""

                If open_grid.Rows.Count > 0 Then
                    name = If(String.IsNullOrEmpty(open_grid.Rows(index_k).Cells(1).Value) = True, "", open_grid.Rows(index_k).Cells(1).Value)
                End If

                If String.Equals(name, "") = False Then

                    '------- check if its a MB type ------
                    Dim isitMB As Boolean : isitMB = False
                    Dim MB_job As String : MB_job = "" 'MB unreleased BOMs have job numbers, old BOMs don't

                    Try
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@mr_name", name)
                        cmd4.CommandText = "SELECT BOM_type, job from Material_Request.mr where mr_name = @mr_name"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then
                            While reader4.Read
                                MB_job = reader4(1).ToString

                                If String.Equals(reader4(0).ToString, "MB") = True Then
                                    isitMB = True
                                End If
                            End While
                        End If

                        reader4.Close()
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try


                    Select Case mr_mode_r  'deal with unreleased BOM

                        Case False

                            '-------- Open BOM MB package ---

                            If isitMB = True Then

                                BOM_types.Visible = True
                                Call BOM_types.Open_BOM_package(name, MB_job)
                                Me.Visible = False

                                '-- if is an old_BOM type then
                            Else

                                My_Material_r.PR_grid.Rows.Clear()

                                Try
                                    Dim cmd4 As New MySqlCommand
                                    cmd4.Parameters.AddWithValue("@mr_name", name)
                                    cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name, Notes, Part_status, Part_type, full_panel, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                                    cmd4.Connection = Login.Connection
                                    Dim reader4 As MySqlDataReader
                                    reader4 = cmd4.ExecuteReader

                                    If reader4.HasRows Then
                                        Dim i As Integer : i = 0
                                        While reader4.Read
                                            My_Material_r.PR_grid.Rows.Add(New String() {})
                                            My_Material_r.PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                                            My_Material_r.PR_grid.Rows(i).Cells(14).Value = reader4(14).ToString

                                            i = i + 1
                                        End While

                                    End If

                                    reader4.Close()
                                    My_Material_r.Text = name
                                    My_Material_r.single_grid.Rows.Clear()
                                    My_Material_r.job_label.Text = "Open Project: "

                                    'Me.Visible = False
                                    My_Material_r.Visible = True
                                    My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage1

                                Catch ex As Exception
                                    MessageBox.Show(ex.ToString)
                                End Try

                                For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
                                    My_Material_r.PR_grid.Rows(i).ReadOnly = False
                                Next
                                My_Material_r.isitreleased = False
                                My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
                                Inventory_manage.part_sel = ""
                                My_Material_r.rev_mode = False

                                Me.Visible = False

                            End If

                        Case True
                            '----------- Open a release MR ----------
                            My_Material_r.PR_grid.Rows.Clear()
                            My_Material_r.mrbox1.Items.Clear()
                            My_Material_r.mrbox2.Items.Clear()

                            Try
                                Dim cmd4 As New MySqlCommand
                                cmd4.Parameters.AddWithValue("@mr_name", name)
                                cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name ,Notes, Part_status, Part_type, full_panel, need_by_date, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
                                cmd4.Connection = Login.Connection
                                Dim reader4 As MySqlDataReader
                                reader4 = cmd4.ExecuteReader

                                If reader4.HasRows Then
                                    Dim i As Integer : i = 0
                                    While reader4.Read
                                        My_Material_r.PR_grid.Rows.Add(New String() {})
                                        My_Material_r.PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(14).Value = reader4(14).ToString
                                        My_Material_r.PR_grid.Rows(i).Cells(15).Value = reader4(15).ToString

                                        i = i + 1
                                    End While

                                End If

                                reader4.Close()
                                My_Material_r.Text = name
                                My_Material_r.single_grid.Rows.Clear()
                                ' Me.Visible = False
                                My_Material_r.Visible = True
                                My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage1

                                '---------- fill combobox with all revisions -------
                                My_Material_r.ComboBox2.Items.Clear()
                                Dim check_cmd As New MySqlCommand
                                check_cmd.Parameters.AddWithValue("@job", open_grid.Rows(index_k).Cells(7).Value)
                                check_cmd.CommandText = "select distinct mr_name from Material_Request.mr where released = 'Y' and job = @job"
                                check_cmd.Connection = Login.Connection
                                check_cmd.ExecuteNonQuery()

                                Dim reader As MySqlDataReader
                                reader = check_cmd.ExecuteReader

                                If reader.HasRows Then
                                    While reader.Read
                                        My_Material_r.ComboBox2.Items.Add(reader(0))
                                        My_Material_r.mrbox1.Items.Add(reader(0))
                                        My_Material_r.mrbox2.Items.Add(reader(0))
                                    End While
                                End If

                                reader.Close()

                            Catch ex As Exception
                                MessageBox.Show(ex.ToString)
                            End Try

                            For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
                                My_Material_r.PR_grid.Rows(i).ReadOnly = True
                            Next

                            My_Material_r.isitreleased = True
                            My_Material_r.open_job = open_grid.Rows(index_k).Cells(7).Value
                            My_Material_r.job_label.Text = "Open Project: " & open_grid.Rows(index_k).Cells(7).Value
                            My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
                            Inventory_manage.part_sel = ""
                            My_Material_r.rev_mode = False
                            Me.Visible = False

                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub BOM_init_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub Label7_MouseEnter(sender As Object, e As EventArgs) Handles Label7.MouseEnter
        Label7.BackColor = Color.DimGray
    End Sub

    Private Sub Label7_MouseLeave(sender As Object, e As EventArgs) Handles Label7.MouseLeave
        Label7.BackColor = Color.FromArgb(61, 60, 78)
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.BackColor = Color.DimGray
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.BackColor = Color.FromArgb(61, 60, 78)
    End Sub

End Class