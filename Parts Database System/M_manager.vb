Imports MySql.Data.MySqlClient

Public Class M_manager
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub M_manager_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If General_management = True Then

            mode_c.Visible = True
            'Fill assignments table
            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.Project_Engineer, p1.Job_number, projects.Job_description from management.proj_assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Engineer Management")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        ElseIf Engineer_management = True Then

            'Fill assignments table
            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.Project_Engineer, p1.Job_number, projects.Job_description from management.proj_assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Engineer")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.Parameters.AddWithValue("@current_user", current_user)
            cmd_j.CommandText = "SELECT distinct Job_number from management.proj_assignments where Project_Engineer = @current_user order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        ElseIf Procurement_management = True Then

            'Fill assignments table
            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.specialist, p1.Job_number, projects.Job_description from management.proc_assig as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Procurement")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.proj_assignments order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        End If
    End Sub

    Private Sub mode_c_SelectedValueChanged(sender As Object, e As EventArgs) Handles mode_c.SelectedValueChanged

        If String.Equals(mode_c.SelectedItem.ToString, "General") = True Then


            'Fill assignments table

            job_b.Items.Clear()
            username_b.Items.Clear()

            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.Project_Engineer, p1.Job_number, projects.Job_description from management.proj_assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Engineer Management")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        ElseIf String.Equals(mode_c.SelectedItem.ToString, "Engineer") = True Then

            'Fill assignments table

            job_b.Items.Clear()
            username_b.Items.Clear()

            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.Project_Engineer, p1.Job_number, projects.Job_description from management.proj_assignments as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Engineer")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.Parameters.AddWithValue("@current_user", current_user)
            cmd_j.CommandText = "SELECT distinct Job_number from management.proj_assignments where Project_Engineer = @current_user order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        ElseIf String.Equals(mode_c.SelectedItem.ToString, "Procurement") = True Then

            'Fill assignments table

            job_b.Items.Clear()
            username_b.Items.Clear()

            Dim table_assi As New DataTable
            Dim adapter_assi As New MySqlDataAdapter("SELECT p1.specialist, p1.Job_number, projects.Job_description from management.proc_assig as p1 INNER JOIN management.projects ON p1.Job_number = projects.Job_number ", Login.Connection)

            adapter_assi.Fill(table_assi)
            Assig_grid.DataSource = table_assi


            Assig_grid.Columns(0).Width = 260
            Assig_grid.Columns(1).Width = 200
            Assig_grid.Columns(2).Width = 460

            'Fill username
            Dim cmd_u As New MySqlCommand
            cmd_u.Parameters.AddWithValue("@role", "Procurement")
            cmd_u.CommandText = "SELECT distinct username from users where Role = @role order by username"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    username_b.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.proj_assignments order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_b.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        End If
    End Sub
End Class