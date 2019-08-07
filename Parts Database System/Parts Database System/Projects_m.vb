Imports MySql.Data.MySqlClient
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class Projects_m

    Public temp_name As String
    Public temp_title As String
    Public temp_email As String
    Public temp_phone As String
    Public Smtp_Server As New SmtpClient


    Private Sub Projects_m_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TabControl1.TabPages.Remove(TabPage2)
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)


        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"

        Try
            '----- fill username ------
            Dim cmd_u As New MySqlCommand
            cmd_u.CommandText = "SELECT distinct username from users  order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'Fill projects table
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT * from management.projects order by Job_number desc", Login.Connection)

            adapter.Fill(table)     'DataViewGrid1 fill
            projects_grid.DataSource = table

            projects_grid.Columns(0).Width = 240
            projects_grid.Columns(1).Width = 1240
            projects_grid.Columns(2).Width = 640
            projects_grid.Columns(3).Width = 640
            projects_grid.Columns(4).Width = 640
            projects_grid.Columns(5).Width = 640


            ''Fill assignments table
            'Dim table_assi As New DataTable
            'Dim adapter_assi As New MySqlDataAdapter("SELECT p1.Employee_Name, p1.Job_number, projects.Job_description from management.assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            'adapter_assi.Fill(table_assi)
            'Assig_grid.DataSource = table_assi


            'Assig_grid.Columns(0).Width = 260
            'Assig_grid.Columns(1).Width = 200
            'Assig_grid.Columns(2).Width = 560


            ''Fill employees table
            'Dim table_em As New DataTable
            'Dim adapter_em As New MySqlDataAdapter("SELECT * from management.employees order by Employee_name ", Login.Connection)

            'adapter_em.Fill(table_em)
            'employee_grid.DataSource = table_em

            'employee_grid.Columns(0).Width = 260
            'employee_grid.Columns(1).Width = 260
            'employee_grid.Columns(2).Width = 660
            'employee_grid.Columns(3).Width = 660


            ''Fill employee combobox
            'Dim cmd_v As New MySqlCommand
            'cmd_v.CommandText = "SELECT distinct Employee_name from management.employees order by Employee_name"

            'cmd_v.Connection = Login.Connection
            'Dim readerv As MySqlDataReader
            'readerv = cmd_v.ExecuteReader

            'If readerv.HasRows Then
            '    While readerv.Read
            '        employee_com.Items.Add(readerv(0))
            '    End While
            'End If

            'readerv.Close()

            'Fill username
            'Dim cmd_u As New MySqlCommand
            'cmd_u.CommandText = "SELECT distinct username from users order by username"

            'cmd_u.Connection = Login.Connection
            'Dim readeru As MySqlDataReader
            'readeru = cmd_u.ExecuteReader

            'If readeru.HasRows Then
            '    While readeru.Read
            '        user_box.Items.Add(readeru(0))
            '    End While
            'End If

            'readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_com.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()


            '--------- fill close projects

            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT distinct p1.job, p1.finished , management.projects.Job_description from Material_Request.mr as p1 INNER JOIN management.projects ON p1.job = management.projects.Job_number where p1.job is not null "
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    close_grid.Rows.Add(New String() {})
                    close_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    close_grid.Rows(i).Cells(1).Value = reader4(2).ToString
                    close_grid.Rows(i).Cells(2).Value = If(reader4.IsDBNull(1) = True, False, True)
                    i = i + 1
                End While

            End If

            reader4.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '----------------------------------- ENTER A NEW ASSIGNMENT ------------------------------------------


        '---------MAKE SURE TEXTBOXES ARE FILL UP (EMPLOYEE AND JOB) ------------------------
        If String.Equals(employee_com.Text, "") = True Or String.Equals(job_com.Text, "") = True Then
            MessageBox.Show("Please Select a Job and Employee")
        Else

            '----------- MAKE SURE THE ASSIGNMENT DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Employee_Name", employee_com.Text)
                cmd.Parameters.AddWithValue("@job", job_com.Text)
                cmd.CommandText = "SELECT * from management.assignments where Employee_Name = @Employee_Name and Job_number = @job"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader



                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@Employee_Name", employee_com.Text)
                    Create_cmd.Parameters.AddWithValue("@Job", job_com.Text)

                    Create_cmd.CommandText = "INSERT INTO management.assignments VALUES (@Employee_Name, @Job)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Assignment added succesfully!")

                Else

                    MessageBox.Show("Duplicate detected!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub

    Sub Refresh_tables(c As MySqlConnection)

        '---------- REFRESH DATA IN GRIDs ----------------------
        Try
            'REFRESH ASSIGNMENTS TABLE 
            'Dim tabl As New DataTable
            'Dim ad As New MySqlDataAdapter("SELECT p1.Employee_Name, p1.Job_number, projects.Job_description from management.assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number", c)
            'ad.Fill(tabl)
            'assig_grid.DataSource = tabl

            'assig_grid.Columns(0).Width = 260
            'assig_grid.Columns(1).Width = 200
            'assig_grid.Columns(2).Width = 560

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT * from management.projects order by Job_number desc", Login.Connection)

            adapter.Fill(table)     'DataViewGrid1 fill
            projects_grid.DataSource = table

            projects_grid.Columns(0).Width = 240
            projects_grid.Columns(1).Width = 1240
            projects_grid.Columns(2).Width = 640
            projects_grid.Columns(3).Width = 640
            projects_grid.Columns(4).Width = 640
            projects_grid.Columns(5).Width = 640

            'Dim table_em As New DataTable
            'Dim adapter_em As New MySqlDataAdapter("SELECT * from management.employees order by Employee_name ", Login.Connection)

            'adapter_em.Fill(table_em)
            'employee_grid.DataSource = table_em

            'employee_grid.Columns(0).Width = 260
            'employee_grid.Columns(1).Width = 260
            'employee_grid.Columns(2).Width = 660
            'employee_grid.Columns(3).Width = 660

        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '============================ Delete assignment ==================================

        Dim index_grid As Integer : index_grid = assig_grid.CurrentCell.RowIndex

        Dim dlgR As DialogResult
        dlgR = MessageBox.Show("Are you sure you want to delete this assignment?", "Attention!", MessageBoxButtons.YesNo)

        If dlgR = DialogResult.Yes Then

            Dim Create_d As New MySqlCommand
            Create_d.Parameters.AddWithValue("@Employee_Name", assig_grid.Rows(index_grid).Cells(0).Value)
            Create_d.Parameters.AddWithValue("@Job", assig_grid.Rows(index_grid).Cells(1).Value)
            Create_d.CommandText = "DELETE FROM management.assignments where Employee_Name = @Employee_Name and Job_number = @Job"
            Create_d.Connection = Login.Connection
            Create_d.ExecuteNonQuery()

            MessageBox.Show("Assignment deleted succesfully")
            Call Refresh_tables(Login.Connection)

        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '--------------SEARCHING JOB STRING ------------------
        Dim found_job As Boolean : found_job = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In projects_grid.Rows
            If String.Compare(row.Cells.Item("Job_number").Value.ToString, Project_textbox.Text) = 0 Then
                rowindex = row.Index
                projects_grid.CurrentCell = projects_grid.Rows(rowindex).Cells(0)
                found_job = True
                Exit For
            End If
        Next
        If found_job = False Then
            MsgBox("Job not found")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '============================ Delete jobs in project table ==================================

        Dim index_grid As Integer : index_grid = projects_grid.CurrentCell.RowIndex

        Dim dlgR As DialogResult
        dlgR = MessageBox.Show("Are you sure you want to delete Project: " & projects_grid.Rows(index_grid).Cells(0).Value & " ?", "Attention!", MessageBoxButtons.YesNo)

        If dlgR = DialogResult.Yes Then

            Try

                '------- check if the job is used in mr table -------
                Dim exist_c As Boolean : exist_c = False


                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@job", projects_grid.Rows(index_grid).Cells(0).Value)
                cmd4.CommandText = "SELECT * from Material_Request.mr where job = @job"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then

                    While reader4.Read
                        exist_c = True
                    End While

                End If

                reader4.Close()

                '----------------------------------------------------
                If exist_c = False Then

                    Dim Create_p As New MySqlCommand
                    Create_p.Parameters.AddWithValue("@Job_n", projects_grid.Rows(index_grid).Cells(0).Value)
                    Create_p.Parameters.AddWithValue("@Job_d", projects_grid.Rows(index_grid).Cells(1).Value)
                    Create_p.CommandText = "DELETE FROM management.projects where Job_number = @Job_n and Job_description = @Job_d"
                    Create_p.Connection = Login.Connection
                    Create_p.ExecuteNonQuery()

                    MessageBox.Show("Project deleted succesfully")
                    Call Refresh_tables(Login.Connection)

                    job_name.Text = ""
                    desc_job.Text = ""
                    ship_a.Text = ""
                    a_ship.Text = ""
                    onsite.Text = ""
                    CPM.Text = ""

                Else
                    MessageBox.Show("This Project is currently being used!")
                End If


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        '==================== ADD JOB =========================


        '---------MAKE SURE TEXTBOXES ARE FILL UP (JOB AND DESCRIPTION) ------------------------
        If String.Equals(job_name.Text, "") = True Or String.Equals(CPM.Text, "") = True Or String.Equals(desc_job.Text, "") = True Or String.Equals(ship_a.Text, "") = True Then


            If String.Equals(job_name.Text, "") = True Then
                job_name.BackColor = Color.Red
            Else
                job_name.BackColor = Color.White
            End If

            If String.Equals(CPM.Text, "") = True Then
                CPM.BackColor = Color.Red
            Else
                CPM.BackColor = Color.White
            End If

            If String.Equals(desc_job.Text, "") = True Then
                desc_job.BackColor = Color.Red
            Else
                desc_job.BackColor = Color.White
            End If

            If String.Equals(ship_a.Text, "") = True Then
                ship_a.BackColor = Color.Red
            Else
                ship_a.BackColor = Color.White
            End If

            MessageBox.Show("Please fill the red textboxes")
        Else

            '----------- MAKE SURE THE PROJECT DOES NOT EXIST ------------

            '--reset color
            job_name.BackColor = Color.White
            CPM.BackColor = Color.White
            desc_job.BackColor = Color.White
            ship_a.BackColor = Color.White
            '--------------

            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@job_number", job_name.Text)
                cmd.Parameters.AddWithValue("@job_desc", desc_job.Text)
                cmd.CommandText = "SELECT * from management.projects where Job_number = @job_number"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@job_number", job_name.Text)
                    Create_cmd.Parameters.AddWithValue("@job_desc", desc_job.Text)
                    Create_cmd.Parameters.AddWithValue("@Shipping_address", ship_a.Text)
                    Create_cmd.Parameters.AddWithValue("@alt_Shipping_address", a_ship.Text)
                    Create_cmd.Parameters.AddWithValue("@onsite_contact", onsite.Text)
                    Create_cmd.Parameters.AddWithValue("@Customer_project_pm", CPM.Text)
                    Create_cmd.CommandText = "INSERT INTO management.projects VALUES (@job_number, @job_desc, @Shipping_address, @alt_Shipping_address, @onsite_contact, @Customer_project_pm)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Job added succesfully!")

                    job_name.Text = ""
                    desc_job.Text = ""
                    ship_a.Text = ""
                    a_ship.Text = ""
                    onsite.Text = ""
                    CPM.Text = ""

                Else

                    MessageBox.Show("Job already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub

    Private Sub employee_grid_DoubleClick(sender As Object, e As EventArgs) Handles employee_grid.DoubleClick

        Dim employee_t As Integer : employee_t = employee_grid.CurrentCell.RowIndex


        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@employee", employee_grid.Rows(employee_t).Cells(0).Value)
            cmd.CommandText = "SELECT * from management.employees where Employee_name = @employee"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read

                    name_box.Text = reader(0).ToString
                    temp_name = name_box.Text

                    Title_box.Text = reader(1).ToString
                    temp_title = Title_box.Text

                    email_box.Text = reader(2).ToString
                    temp_email = email_box.Text

                    phone_box.Text = reader(3).ToString
                    temp_phone = phone_box.Text

                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        '--------------------------- Enter employee --------------------------
        '---------MAKE SURE TEXTBOXES ARE FILL UP (name AND title) ------------------------

        If String.Equals(name_box.Text, "") = True Or String.Equals(Title_box.Text, "") = True Then
            MessageBox.Show("Please Enter Employee name and Title")
        Else

            '----------- MAKE SURE THE EMPLOYEE DOES NOT EXIST ------------

            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand

                cmd.Parameters.AddWithValue("@employee_name", name_box.Text)

                cmd.CommandText = "SELECT * from management.employees where Employee_name = @employee_name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@employee_name", name_box.Text)
                    Create_cmd.Parameters.AddWithValue("@title", Title_box.Text)
                    Create_cmd.Parameters.AddWithValue("@email", email_box.Text)
                    Create_cmd.Parameters.AddWithValue("@phone", phone_box.Text)

                    Create_cmd.CommandText = "INSERT INTO management.employees VALUES (@employee_name, @title, @email , @phone)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Employee added succesfully!")

                Else

                    MessageBox.Show("Employee already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '-----------DELETE EMPLOYEE --------------------

        If String.Equals(name_box.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name, Part Status and Part Type")
        Else
            Try

                Dim Create_cmd As New MySqlCommand

                Dim dlgR As DialogResult
                dlgR = MessageBox.Show("Are you sure you want to delete " & name_box.Text & "?", "Attention!", MessageBoxButtons.YesNo)

                If dlgR = DialogResult.Yes Then

                    Create_cmd.Parameters.AddWithValue("@Employee_Name", name_box.Text)

                    Create_cmd.CommandText = "DELETE FROM management.employees where Employee_name = @Employee_Name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    MessageBox.Show(name_box.Text & "  deleted succesfully")
                    Call Refresh_tables(Login.Connection)

                End If


            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If


    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim update_flag As Boolean : update_flag = True

        If String.Equals(name_box.Text, "") = True Or String.Equals(Title_box.Text, "") = True And String.Equals(temp_name, "") = True Then
            MessageBox.Show("Please Fill Employee Name and title")
        Else


            Try
                Dim cmd As New MySqlCommand
                cmd.CommandText = "SELECT Employee_name from management.employees where Employee_name = """ & temp_name & """"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                'MAKE SURE THE PART EXIST SO IT CAN BE UPDATED
                If Not reader.HasRows Then
                    MessageBox.Show("Employee does not exist! ... Update incomplete")
                    update_flag = False
                End If

                reader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try


            '--------------------------------------------------------------------------------------------------------------------------------------

            If update_flag = True Then

                Try
                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@employee_name", name_box.Text)
                    Create_cmd.Parameters.AddWithValue("@title", Title_box.Text)
                    Create_cmd.Parameters.AddWithValue("@email", email_box.Text)
                    Create_cmd.Parameters.AddWithValue("@phone", phone_box.Text)

                    'temp variables
                    Create_cmd.Parameters.AddWithValue("@t_employee_name", temp_name)
                    Create_cmd.Parameters.AddWithValue("@t_title", temp_title)
                    Create_cmd.Parameters.AddWithValue("@t_email", temp_email)
                    Create_cmd.Parameters.AddWithValue("@t_phone", temp_phone)


                    Create_cmd.CommandText = "UPDATE management.employees  SET  Employee_name  = @employee_name, Title = @title, Email_address = @email, phone_n = @phone where Employee_name = @t_employee_name and Title = @t_title and Email_address = @t_email and phone_n =  @t_phone "
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Employee updated succesfully")

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If
        End If
    End Sub

    Private Sub username_b_SelectedValueChanged(sender As Object, e As EventArgs) Handles username_b.SelectedValueChanged
        Try
            history_grid.Rows.Clear()
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@username", username_b.Text)
            cmd4.CommandText = "SELECT job, action_m, role , date_m from management.history where username = @username"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    history_grid.Rows.Add(New String() {})
                    history_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    history_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    history_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    history_grid.Rows(i).Cells(3).Value = reader4(3).ToString


                    i = i + 1
                End While

            End If

            reader4.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '------ mananage close projects --------
        Try

            Dim billing_message As String : billing_message = ""

            For i = 0 To close_grid.Rows.Count - 1

                Dim fini As String : fini = ""

                If Convert.ToBoolean(close_grid.Rows(i).Cells(2).Value) = True Then


                    '---- check if job was closed before -----
                    Dim cmd21 As New MySqlCommand
                    cmd21.Parameters.AddWithValue("@job", close_grid.Rows(i).Cells(0).Value)
                    cmd21.CommandText = "SELECT finished from Material_Request.mr where job = @job"
                    cmd21.Connection = Login.Connection
                    Dim reader21 As MySqlDataReader
                    reader21 = cmd21.ExecuteReader

                    If reader21.HasRows Then
                        While reader21.Read
                            fini = If(reader21(0) Is DBNull.Value, "", reader21(0))
                        End While
                    End If

                    reader21.Close()

                    '--------------------------------
                    'Dim Create_cmd As New MySqlCommand
                    'Create_cmd.Parameters.AddWithValue("@job", close_grid.Rows(i).Cells(0).Value)

                    'Create_cmd.CommandText = "UPDATE Material_Request.mr SET finished = 'Y' where job = @job"
                    'Create_cmd.Connection = Login.Connection
                    'Create_cmd.ExecuteNonQuery()

                    If String.Equals(fini, "") = True Then
                        '--check costs
                        If check_totals(close_grid.Rows(i).Cells(0).Value) = True Then

                            Dim result As DialogResult = MessageBox.Show("Parts with zero cost have been found, Do you still want to close Project " & close_grid.Rows(i).Cells(0).Value & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                            If (result = DialogResult.Yes) Then

                                Dim Create_cmd As New MySqlCommand
                                Create_cmd.Parameters.AddWithValue("@job", close_grid.Rows(i).Cells(0).Value)

                                Create_cmd.CommandText = "UPDATE Material_Request.mr SET finished = 'Y' where job = @job"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()

                                Call close_project_data(close_grid.Rows(i).Cells(0).Value)

                                billing_message = billing_message & close_grid.Rows(i).Cells(0).Value & vbCrLf

                            End If

                        Else

                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@job", close_grid.Rows(i).Cells(0).Value)

                            Create_cmd.CommandText = "UPDATE Material_Request.mr SET finished = 'Y' where job = @job"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()

                            Call close_project_data(close_grid.Rows(i).Cells(0).Value)

                            billing_message = billing_message & close_grid.Rows(i).Cells(0).Value & vbCrLf

                        End If

                    End If


                Else

                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@job", close_grid.Rows(i).Cells(0).Value)

                    Create_cmd.CommandText = "UPDATE Material_Request.mr  SET finished = NULL where job = @job"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                End If
            Next

            If String.Equals(billing_message, "") = False Then
                billing_message = "The Following Projects Total Cost are ready" & vbCrLf & vbCrLf & billing_message

                If enable_mess = True Then
                    'send email
                End If

            End If

            MessageBox.Show("Done")

        Catch ex As Exception
            MessageBox.Show(ex, ToString)
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If String.Equals(job_name.Text, "") = False And String.IsNullOrEmpty(job_name.Text) = False Then
            Try
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@job_number", job_name.Text)
                Create_cmd.Parameters.AddWithValue("@job_desc", desc_job.Text)
                Create_cmd.Parameters.AddWithValue("@Shipping_address", ship_a.Text)
                Create_cmd.Parameters.AddWithValue("@alt_Shipping_address", a_ship.Text)
                Create_cmd.Parameters.AddWithValue("@onsite_contact", onsite.Text)
                Create_cmd.Parameters.AddWithValue("@Customer_project_pm", CPM.Text)

                Create_cmd.CommandText = "UPDATE management.projects  SET  Job_description = @job_desc, Shipping_address = @Shipping_address, alt_Shipping_address =  @alt_Shipping_address,  onsite_contact = @onsite_contact, Customer_project_pm = @Customer_project_pm where Job_Number = @job_number"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Call Refresh_tables(Login.Connection)
                MessageBox.Show("Job updated succesfully")
                job_name.Text = ""
                desc_job.Text = ""
                ship_a.Text = ""
                a_ship.Text = ""
                onsite.Text = ""
                CPM.Text = ""

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub projects_grid_DoubleClick(sender As Object, e As EventArgs) Handles projects_grid.DoubleClick
        '-- fill textboxes
        Dim job As String = projects_grid.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Job_number", job)
            cmd.CommandText = "SELECT * from management.projects where Job_number = @Job_number"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    job_name.Text = reader(0).ToString
                    desc_job.Text = reader(1).ToString
                    ship_a.Text = reader(2).ToString
                    a_ship.Text = reader(3).ToString
                    onsite.Text = reader(4).ToString
                    CPM.Text = reader(5).ToString
                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Function check_totals(job As String) As Boolean

        '--check zero or non numeric qtys in tables and give the user an option to create an excel report---
        '--- table general inv ---------
        Dim dimen_table = New DataTable

        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("Min", GetType(String))


        '--------- table general inv zero---------
        Dim dimen_table2 = New DataTable

        dimen_table2.Columns.Add("Part_No", GetType(String))
        dimen_table2.Columns.Add("Price", GetType(String))
        dimen_table2.Columns.Add("Min", GetType(String))

        '--- table project specific ---------
        Dim ps_table = New DataTable

        ps_table.Columns.Add("Part_No", GetType(String))
        ps_table.Columns.Add("Cost", GetType(String))


        '--- table project specific ---------
        Dim ps_table2 = New DataTable

        ps_table2.Columns.Add("Part_No", GetType(String))
        ps_table2.Columns.Add("Cost", GetType(String))

        '--- table add-return ---------
        Dim ar_table = New DataTable

        ar_table.Columns.Add("Part_No", GetType(String))
        ar_table.Columns.Add("Cost", GetType(String))

        '--- table add-return ---------
        Dim ar_table2 = New DataTable

        ar_table2.Columns.Add("Part_No", GetType(String))
        ar_table2.Columns.Add("Cost", GetType(String))



        '----------------------------------------
        Dim errors_f As Boolean : errors_f = False
        Dim no_tracking_report As Boolean : no_tracking_report = False

        check_totals = False

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            Try
                '--------//////////////// check if inventory only cost is numeric and non zero -------


                Dim mr_name As String : mr_name = "not_found"
                Dim cmd21 As New MySqlCommand
                cmd21.Parameters.AddWithValue("@job", job)
                cmd21.CommandText = "select mr_name from Material_Request.mr where job = @job and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"
                cmd21.Connection = Login.Connection
                Dim reader21 As MySqlDataReader
                reader21 = cmd21.ExecuteReader

                If reader21.HasRows Then
                    While reader21.Read
                        mr_name = reader21(0).ToString
                    End While
                End If

                reader21.Close()

                mr_name = Procurement_Overview.get_last_revision(mr_name)


                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@mr_name", mr_name)
                cmd41.CommandText = "SELECT  Part_No, Price from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader


                If reader41.HasRows Then
                    While reader41.Read
                        dimen_table.Rows.Add(reader41(0).ToString, reader41(1), 0)
                    End While
                End If

                reader41.Close()

                '-----------------

                For i = 0 To dimen_table.Rows.Count - 1

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0).ToString)
                    cmd5.CommandText = "SELECT min_qty from inventory.inventory_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    Dim reader5 As MySqlDataReader
                    reader5 = cmd5.ExecuteReader

                    If reader5.HasRows Then
                        While reader5.Read
                            dimen_table.Rows(i).Item(2) = reader5(0)
                        End While
                    End If

                    reader5.Close()

                    If dimen_table.Rows(i).Item(2) > 0 Then

                        If IsNumeric(dimen_table.Rows(i).Item(1)) = False Then
                            errors_f = True
                            dimen_table2.Rows.Add(dimen_table.Rows(i).Item(0), dimen_table.Rows(i).Item(1))
                        Else
                            If dimen_table.Rows(i).Item(1) <= 0 Then
                                dimen_table2.Rows.Add(dimen_table.Rows(i).Item(0), dimen_table.Rows(i).Item(1))
                                errors_f = True
                            End If
                        End If

                    End If

                    '-------------------
                Next



                '-------//////////////////// check if project specific (tracking reports exist and all cost are numeric

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@job", job)
                cmd4.CommandText = "SELECT Part_No, cost from Tracking_Reports.my_tracking_reports where job = @job"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader


                If reader4.HasRows Then
                    no_tracking_report = False

                    While reader4.Read
                        ps_table.Rows.Add(reader4(0).ToString, reader4(1).ToString)
                    End While

                Else
                    no_tracking_report = True
                End If

                reader4.Close()

                For i = 0 To ps_table.Rows.Count - 1
                    If IsNumeric(ps_table.Rows(i).Item(1)) = False Then
                        ps_table2.Rows.Add(ps_table.Rows(i).Item(0), ps_table.Rows(i).Item(1))
                        errors_f = True
                    Else
                        If ps_table.Rows(i).Item(1) <= 0 Then
                            ps_table2.Rows.Add(ps_table.Rows(i).Item(0), ps_table.Rows(i).Item(1))
                            errors_f = True
                        End If
                    End If
                Next


                '---------//////////////////////// check if add/return cost are numeric and non zero
                Dim cmd46 As New MySqlCommand
                cmd46.Parameters.AddWithValue("@job", job)
                cmd46.CommandText = "SELECT Part_No, Cost from Tracking_Reports.add_return where job = @job"
                cmd46.Connection = Login.Connection
                Dim reader46 As MySqlDataReader
                reader46 = cmd46.ExecuteReader


                If reader46.HasRows Then
                    While reader46.Read
                        ar_table.Rows.Add(reader46(0).ToString, reader46(1).ToString)
                    End While
                End If

                reader46.Close()

                For i = 0 To ar_table.Rows.Count - 1
                    If IsNumeric(ar_table.Rows(i).Item(1)) = False Then
                        ar_table2.Rows.Add(ar_table.Rows(i).Item(0), ar_table.Rows(i).Item(1))
                        errors_f = True
                    Else
                        If ar_table.Rows(i).Item(1) <= 0 Then
                            ar_table2.Rows.Add(ar_table.Rows(i).Item(0), ar_table.Rows(i).Item(1))
                            errors_f = True
                        End If
                    End If
                Next


                '----------------------------------------------------------

                If errors_f = True Then

                    check_totals = True

                    Dim result As DialogResult = MessageBox.Show("Would you like to create a report of the parts with zero cost for project :" & job & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result = DialogResult.Yes) Then

                        Dim xlWorkBook As Excel.Workbook
                        Dim xlWorkSheet As Excel.Worksheet
                        Dim misValue As Object = System.Reflection.Missing.Value
                        xlWorkBook = xlApp.Workbooks.Add(misValue)
                        xlWorkSheet = xlWorkBook.Sheets("sheet1")
                        xlWorkSheet.Range("A:B").ColumnWidth = 40
                        xlWorkSheet.Range("C:C").ColumnWidth = 30
                        xlWorkSheet.Range("D:E").ColumnWidth = 20
                        xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter

                        Dim index As Integer : index = 5
                        xlWorkSheet.Cells(1, 1) = "Parts with zero or no Price for Project " & job
                        xlWorkSheet.Cells(3, 1) = "Parts from inventory"
                        xlWorkSheet.Cells(3, 1).interior.color = Color.Yellow

                        For i = 0 To dimen_table2.Rows.Count - 1
                            xlWorkSheet.Cells(index, 1) = dimen_table2.Rows(i).Item(0).ToString
                            index = index + 1
                        Next

                        index = index + 1
                        xlWorkSheet.Cells(index, 1) = "Project Specific Parts"
                        xlWorkSheet.Cells(index, 1).interior.color = Color.Yellow
                        index = index + 1

                        If no_tracking_report = True Then
                            xlWorkSheet.Cells(index, 1) = "Project Specific Parts data not found"
                            index = index + 1
                        Else

                            For i = 0 To ps_table2.Rows.Count - 1
                                xlWorkSheet.Cells(index, 1) = ps_table2.Rows(i).Item(0).ToString
                                index = index + 1
                            Next

                        End If

                        xlWorkSheet.Cells(index, 1) = "Add / return Parts"
                        xlWorkSheet.Cells(index, 1).interior.color = Color.Yellow
                        index = index + 1

                        For i = 0 To ar_table2.Rows.Count - 1
                            xlWorkSheet.Cells(index, 1) = ar_table2.Rows(i).Item(0).ToString
                            index = index + 1
                        Next

                        SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                            xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                        End If

                        xlWorkBook.Close(False)
                        xlApp.Quit()


                        Marshal.ReleaseComObject(xlApp)

                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If



    End Function

    Sub close_project_data(job As String)

        '--enter all data to total cost table -----
        '----------- Cal totals of job ---------------

        Dim Inv_c As Double : Inv_c = 0
        Dim PS_c As Double : PS_c = 0
        Dim AR_c As Double : AR_c = 0
        Dim bulk As Double : bulk = 0   'bulk
        Dim ship As Double : ship = 0   'ship


        '--------- table general inv ---------
        Dim dimen_table = New DataTable

        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Min", GetType(String))


        '--- table project specific ---------
        Dim ps_table = New DataTable

        ps_table.Columns.Add("Part_No", GetType(String))
        ps_table.Columns.Add("Qty", GetType(String))
        ps_table.Columns.Add("Cost", GetType(String))


        '--- table add-return ---------
        Dim ar_table = New DataTable

        ar_table.Columns.Add("Part_No", GetType(String))
        ar_table.Columns.Add("Qty", GetType(String))
        ar_table.Columns.Add("Cost", GetType(String))
        ar_table.Columns.Add("task", GetType(String))


        '----------------------------------------


        '--------------- fill tables ---------------
        Try

            '-----------/////////////// fill only inventory table ////////////-------------------

            Dim mr_name As String : mr_name = "not_found"
            Dim cmd21 As New MySqlCommand
            cmd21.Parameters.AddWithValue("@job", job)
            cmd21.CommandText = "select mr_name from Material_Request.mr where job = @job and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"
            cmd21.Connection = Login.Connection
            Dim reader21 As MySqlDataReader
            reader21 = cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    mr_name = reader21(0).ToString
                End While
            End If

            reader21.Close()

            mr_name = Procurement_Overview.get_last_revision(mr_name)

            '--/////////////////////

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT  Part_No, Price,  Qty from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, 0)
                End While
            End If

            reader4.Close()

            '-----------------

            For i = 0 To dimen_table.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0).ToString)
                cmd5.CommandText = "SELECT min_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        dimen_table.Rows(i).Item(3) = reader5(0)
                    End While
                End If

                reader5.Close()

                '-- get totals --

                If CType(dimen_table.Rows(i).Item(3), Double) > 0 Then
                    Inv_c = Inv_c + (If(IsNumeric(dimen_table.Rows(i).Item(1)), dimen_table.Rows(i).Item(1), 0) * If(IsNumeric(dimen_table.Rows(i).Item(1)), dimen_table.Rows(i).Item(1), 0))
                End If

                '-------------------

            Next

            '-----------/////////////// fill Project specific ////////////-------------------
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "SELECT Part_No, qty_purchased, cost from Tracking_Reports.my_tracking_reports where job = @job"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader


            If reader41.HasRows Then
                While reader41.Read
                    ps_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString)
                End While
            End If

            reader41.Close()

            For i = 0 To ps_table.Rows.Count - 1
                PS_c = PS_c + (If(IsNumeric(ps_table.Rows(i).Item(1)), ps_table.Rows(i).Item(1), 0) * If(IsNumeric(ps_table.Rows(i).Item(2)), ps_table.Rows(i).Item(2), 0))
            Next

            '------------------------------------------------------------

            '-----------/////////////// fill add return  ////////////-------------------

            Dim cmd42 As New MySqlCommand
            cmd42.Parameters.AddWithValue("@job", job)
            cmd42.CommandText = "SELECT Part_No, qty, Cost, task from Tracking_Reports.add_return where job = @job"
            cmd42.Connection = Login.Connection
            Dim reader42 As MySqlDataReader
            reader42 = cmd42.ExecuteReader


            If reader42.HasRows Then
                While reader42.Read
                    ar_table.Rows.Add(reader42(0).ToString, reader42(1).ToString, reader42(2).ToString, reader42(3).ToString)
                End While
            End If

            reader42.Close()

            For i = 0 To ar_table.Rows.Count - 1

                If String.Equals(dimen_table.Rows(i).Item(3), "Add") Then
                    AR_c = AR_c + (If(IsNumeric(ar_table.Rows(i).Item(1)), ar_table.Rows(i).Item(1), 0) * If(IsNumeric(ar_table.Rows(i).Item(2)), ar_table.Rows(i).Item(2), 0))
                Else
                    AR_c = AR_c + (If(IsNumeric(ar_table.Rows(i).Item(1)), ar_table.Rows(i).Item(1), 0) * If(IsNumeric(ar_table.Rows(i).Item(2)), ar_table.Rows(i).Item(2), 0) * -1)
                End If
            Next

            '---------- shipping and bulk ------
            Dim cmd211 As New MySqlCommand
            cmd211.Parameters.AddWithValue("@job", job)
            cmd211.CommandText = "SELECT bulk_cost, shipping_cost from management.total_cost where job = @job"
            cmd211.Connection = Login.Connection
            Dim reader211 As MySqlDataReader
            reader211 = cmd211.ExecuteReader

            If reader211.HasRows Then
                While reader211.Read
                    bulk = If(reader211(0) Is DBNull.Value, 0, Decimal.Round(reader211(0).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    ship = If(reader211(1) Is DBNull.Value, 0, Decimal.Round(reader211(1).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                End While
            End If

            reader211.Close()
            '-----------------------------------

            '----------- put all data in textboxes --------------
            Inv_c = Decimal.Round(Inv_c, 2, MidpointRounding.AwayFromZero).ToString("N")
            PS_c = Decimal.Round(PS_c, 2, MidpointRounding.AwayFromZero).ToString("N")
            AR_c = Decimal.Round(AR_c, 2, MidpointRounding.AwayFromZero).ToString("N")

            '----------------------------------------------------

            '-- if record exist update else create it
            Dim exist_c As Boolean : exist_c = False

            Dim cmd54 As New MySqlCommand
            cmd54.Parameters.Clear()
            cmd54.Parameters.AddWithValue("@job", job)
            cmd54.CommandText = "SELECT * from management.total_cost where job = @job"
            cmd54.Connection = Login.Connection
            Dim reader54 As MySqlDataReader
            reader54 = cmd54.ExecuteReader

            If reader54.HasRows Then
                While reader54.Read
                    exist_c = True
                End While
            End If

            reader54.Close()

            If exist_c = True Then

                '-- update it
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@job", job)
                Create_cmd.Parameters.AddWithValue("@gi_cost", If(IsNumeric(Inv_c), Inv_c, 0))
                Create_cmd.Parameters.AddWithValue("@ps_cost", If(IsNumeric(PS_c), PS_c, 0))
                Create_cmd.Parameters.AddWithValue("@ar_cost", If(IsNumeric(AR_c), AR_c, 0))
                Create_cmd.Parameters.AddWithValue("@bulk_cost", If(IsNumeric(bulk), bulk, 0))
                Create_cmd.Parameters.AddWithValue("@shipping_cost", If(IsNumeric(ship), ship, 0))


                Create_cmd.CommandText = "UPDATE management.total_cost  SET gi_cost = @gi_cost, ps_cost = @ps_cost, ar_cost = @ar_cost, shipping_cost = @shipping_cost, bulk_cost = @bulk_cost  where job = @job"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Else

                '------ insert
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@job", job)
                Create_cmd.Parameters.AddWithValue("@gi_cost", If(IsNumeric(Inv_c), Inv_c, 0))
                Create_cmd.Parameters.AddWithValue("@ps_cost", If(IsNumeric(PS_c), PS_c, 0))
                Create_cmd.Parameters.AddWithValue("@ar_cost", If(IsNumeric(AR_c), AR_c, 0))
                Create_cmd.Parameters.AddWithValue("@bulk_cost", If(IsNumeric(bulk), bulk, 0))
                Create_cmd.Parameters.AddWithValue("@shipping_cost", If(IsNumeric(ship), ship, 0))

                Create_cmd.CommandText = "INSERT INTO management.total_cost(job, gi_cost, ps_cost, ar_cost, bulk_cost, shipping_cost) VALUES (@job, @gi_cost, @ps_cost, @ar_cost, @bulk_cost, @shipping_cost)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

End Class