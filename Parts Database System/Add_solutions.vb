Imports MySql.Data.MySqlClient

Public Class Add_solutions
    Private Sub Add_solutions_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim table_kits As New DataTable
        Dim adapter_kits As New MySqlDataAdapter("SELECT * from quote_table.feature_solutions", Login.Connection)

        adapter_kits.Fill(table_kits)     'kits grid fill
        sol_table.DataSource = table_kits


        'Setting Columns size for KITS Datagrid
        For i = 0 To sol_table.ColumnCount - 1
            With sol_table.Columns(i)
                .Width = 550
            End With
        Next i
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        '---------MAKE SURE TEXTBOXES ARE FILL UP  ------------------------
        If String.Equals(sol_name.Text, "") = True Or String.Equals(desc_name.Text, "") = True Then
            MessageBox.Show("Please Fill Solution name and description")
        Else

            '----------- MAKE SURE THE FEATURE_CODE DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@solution", sol_name.Text)
                cmd.Parameters.AddWithValue("@description", desc_name.Text)
                cmd.CommandText = "SELECT * from quote_table.feature_solutions where solution_name = @solution and sol_description = @description"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@solution", sol_name.Text)
                    Create_cmd.Parameters.AddWithValue("@description", desc_name.Text)
                    Create_cmd.CommandText = "INSERT INTO quote_table.feature_solutions(solution_name, sol_description) VALUES (@solution, @description)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    MessageBox.Show("Solution added succesfully!")
                    Me.Visible = False

                Else
                    MessageBox.Show("Solution already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If


            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub
End Class