Imports MySql.Data.MySqlClient
Public Class Job_assign
    Private Sub Job_assign_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Clear()

        Try
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    ComboBox1.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        If IsNothing(ComboBox1.SelectedItem) = False Then

            Dim not_applic As Boolean : not_applic = False

            '-- check if job is in used
            Try
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", ComboBox1.Text)
                check_cmd.CommandText = "select * from Material_Request.mr where job = @job"
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    not_applic = True
                End If

                reader.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            If not_applic = False Then

                '---get shipping address

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.AddWithValue("@job", ComboBox1.Text)
                    check_cmd.CommandText = "select Shipping_address from management.projects where Job_number = @job"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            BOM_types.shipping_ad = reader(0).ToString
                        End While
                    End If

                    reader.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try



                BOM_types.job_Selected = ComboBox1.Text
                BOM_types.job_label.Text = "Current Job:  " & ComboBox1.Text
                Me.Visible = False

            Else
                MessageBox.Show("There is a Master BOM already associated with this Project!")

            End If
        Else
                MessageBox.Show("Please select a Job Number")
        End If
    End Sub

    Private Sub Job_assign_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If IsNothing(ComboBox1.SelectedItem) = True Then
            MessageBox.Show("Changes won't be saved until you select a Job number")
        End If
    End Sub

    Private Sub Job_assign_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class