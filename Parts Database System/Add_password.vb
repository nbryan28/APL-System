Imports MySql.Data.MySqlClient
Public Class Add_password
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.Equals(passbox.Text, "") = False And String.IsNullOrEmpty(passbox.Text) = False Then

            Select Case User_form.mode
                Case "add"

                    Call User_form.add_username()

                Case "update"

                    Call User_form.update_username()

                Case Else
                    Try
                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@username", Login_rights.TextBox1.Text)
                        Create_cmd.Parameters.AddWithValue("@password", passbox.Text)
                        Create_cmd.CommandText = "UPDATE users SET password = @password, reset_t = '' where username = @username"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

            End Select

            Me.Visible = False
            Me.Close()

        Else

            MessageBox.Show("Password cannot be blank")

        End If

    End Sub

    Private Sub Add_password_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        passbox.Text = ""
    End Sub
End Class