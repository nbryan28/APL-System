Imports MySql.Data.MySqlClient


Public Class Assign_Projects
    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Assign_Projects_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            project_box.Items.Clear()
            user_box.Items.Clear()

            '------------ adding projects -----------

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr  order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    project_box.Items.Add(reader2(0))
                End While
            End If

            reader2.Close()

            '--------- adding username -----------

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.CommandText = "select distinct username from users order by username"
            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then
                While reader1.Read
                    user_box.Items.Add(reader1(0))
                End While
            End If

            reader1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub project_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles project_box.SelectedValueChanged

        '--- reload assigmnet grid and description ----

        assi_grid.Rows.Clear()

        Try

            '------ set job description -----

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@job", project_box.Text)
            check_cmd2.CommandText = "select Job_description from management.projects where Job_number = @job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    job_desc.Text = reader2(0).ToString
                End While
            End If

            reader2.Close()



            '--- fill assignemnts table -------
            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@Job_Number", project_box.Text)
            check_cmd1.CommandText = "select Employee_Name from management.assignments where Job_Number = @Job_Number"
            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then
                While reader1.Read
                    assi_grid.Rows.Add(New String() {reader1(0).ToString})
                End While
            End If

            reader1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub user_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles user_box.SelectedValueChanged
        '---reload username and fill the boxes ---------

        u_box.Text = ""
        e_box.Text = ""
        r_box.Text = ""

        Dim check_cmd2 As New MySqlCommand
        check_cmd2.Parameters.AddWithValue("@username", user_box.Text)
        check_cmd2.CommandText = "select username, Role, email from users where username = @username"
        check_cmd2.Connection = Login.Connection
        check_cmd2.ExecuteNonQuery()

        Dim reader2 As MySqlDataReader
        reader2 = check_cmd2.ExecuteReader

        If reader2.HasRows Then
            While reader2.Read
                u_box.Text = reader2(0).ToString
                e_box.Text = reader2(1).ToString
                r_box.Text = reader2(2).ToString
            End While
        End If

        reader2.Close()



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- add user to project

        If String.Equals(job_desc.Text, "Project Description") = False And String.IsNullOrEmpty(u_box.Text) = False Then

            Dim exist_c As Boolean : exist_c = False
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@Employee_Name", u_box.Text)
            cmd4.Parameters.AddWithValue("@Job_number", project_box.Text)

            cmd4.CommandText = "Select * from management.assignments where Employee_Name = @Employee_Name and Job_number = @Job_number"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    exist_c = True
                End While
            End If

            reader4.Close()

            If exist_c = False Then

                Dim insert_cmd As New MySqlCommand
                insert_cmd.Parameters.AddWithValue("@Employee_Name", u_box.Text)
                insert_cmd.Parameters.AddWithValue("@Job_number", project_box.Text)

                insert_cmd.CommandText = "INSERT INTO management.assignments VALUES (@Employee_Name, @Job_number)"
                insert_cmd.Connection = Login.Connection
                insert_cmd.ExecuteNonQuery()
                MessageBox.Show(u_box.Text & " has been assigned to " & job_desc.Text & " successfully!")
                Call refresh_Table()
            Else
                MessageBox.Show(u_box.Text & " already assigned to this project")

            End If


        End If




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '--removing user from project ---

        If assi_grid.Rows.Count > 0 Then

            Dim index_k = assi_grid.CurrentCell.RowIndex
            Dim user_n As String : user_n = assi_grid.Rows(index_k).Cells(0).Value

            Dim result As DialogResult = MessageBox.Show("Are you sure, you want to remove this user from Project " & project_box.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                Dim check_cmd22 As New MySqlCommand
                check_cmd22.Parameters.AddWithValue("@Employee_Name", user_n)
                check_cmd22.Parameters.AddWithValue("@Job_Number", project_box.Text)
                check_cmd22.CommandText = "delete from management.assignments where Employee_Name = @Employee_Name and Job_number = @Job_Number"
                check_cmd22.Connection = Login.Connection
                check_cmd22.ExecuteNonQuery()

                MessageBox.Show("User removed succesfully!")

                Call refresh_table()

            End If

        End If

    End Sub

    Sub refresh_Table()
        assi_grid.Rows.Clear()

        Try
            '--- fill assignemnts table -------
            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@Job_Number", project_box.Text)
            check_cmd1.CommandText = "select Employee_Name from management.assignments where Job_Number = @Job_Number"
            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then
                While reader1.Read
                    assi_grid.Rows.Add(New String() {reader1(0).ToString})
                End While
            End If

            reader1.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class