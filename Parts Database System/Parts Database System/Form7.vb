Imports MySql.Data.MySqlClient

Public Class Login_rights
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Login_rights_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Log_out = False
        Engineer = False
        Engineer_management = False
        Procurement = False
        Procurement_management = False

        Try

            '-------- see if pass need to be reseted
            Dim reset_m As Boolean = False
            Dim cmdr As New MySqlCommand
            cmdr.Parameters.AddWithValue("@username", TextBox1.Text)
            cmdr.CommandText = "SELECT reset_t from users where username = @username"
            cmdr.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmdr.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read()
                    If String.Equals(reader2(0).ToString, "APL") = True Then
                        reset_m = True
                    End If
                End While
            End If

            reader2.Close()

            '-----------------------------------

            If reset_m = False Then

                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@username", TextBox1.Text)
                cmd.Parameters.AddWithValue("@password", TextBox2.Text)
                cmd.CommandText = "SELECT * from users where username = @username and password = @password"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    Log_out = True
                    Form1.Login_l.Text = "User Log Out"
                    Form1.Text = "APL (" & TextBox1.Text & "'s Session)"
                    current_user = TextBox1.Text

                    While reader.Read()

                        Select Case True

                            Case String.Equals(reader(3).ToString, "Engineer")
                                Engineer = True
                            Case String.Equals(reader(3).ToString, "Engineer Management")
                                Engineer_management = True
                            Case String.Equals(reader(3).ToString, "Procurement")
                                Procurement = True
                            Case String.Equals(reader(3).ToString, "Procurement Management")
                                Procurement_management = True
                            Case String.Equals(reader(3).ToString, "General Management")
                                General_management = True
                            Case String.Equals(reader(3).ToString, "Inventory")
                                Inventory_m = True
                            Case String.Equals(reader(3).ToString, "Manufacturing")
                                Manufacturing = True
                        End Select

                    End While

                    reader.Close()

                    'update status table
                    Dim status_cmd As New MySqlCommand
                    status_cmd.Parameters.AddWithValue("@User", TextBox1.Text)
                    status_cmd.CommandText = "UPDATE management.status_user SET Status = 'Active', Log_in_time = now() where User = @User"
                    status_cmd.Connection = Login.Connection
                    status_cmd.ExecuteNonQuery()

                    Form1.TabControl1.TabPages.Insert(5, Form1.TabPage10)


                Else
                    MessageBox.Show("Incorrect Username or Password!")
                    reader.Close()
                End If
            Else

                Me.Visible = False
                Add_password.ShowDialog()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Me.Visible = False
    End Sub

    Private Sub TextBox1_VisibleChanged(sender As Object, e As EventArgs) Handles TextBox1.VisibleChanged
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
End Class