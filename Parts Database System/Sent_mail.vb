Imports MySql.Data.MySqlClient

Public Class Sent_mail

    Public role_me As String
    Public role_rec As String

    Private Sub Sent_mail_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Try
            Dim cmd_v As New MySqlCommand
            Dim cmd_v2 As New MySqlCommand
            cmd_v.CommandText = "SELECT username from users"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    user_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

            cmd_v2.Parameters.AddWithValue("@user", current_user)
            cmd_v2.CommandText = "SELECT Role from users where username = @user"

            cmd_v2.Connection = Login.Connection
            Dim readerv2 As MySqlDataReader
            readerv2 = cmd_v2.ExecuteReader

            If readerv2.HasRows Then
                While readerv2.Read
                    role_me = readerv2(0).ToString
                End While
            End If

            readerv2.Close()



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        'send email ------------------ get role of receiver
        Dim min_flag As Boolean : min_flag = True

        Try
            Dim cmd_v2 As New MySqlCommand
            cmd_v2.Parameters.AddWithValue("@user", user_box.Text)
            cmd_v2.CommandText = "SELECT Role from users where username = @user"
            cmd_v2.Connection = Login.Connection
            Dim readerv2 As MySqlDataReader
            readerv2 = cmd_v2.ExecuteReader

            If readerv2.HasRows Then
                While readerv2.Read
                    role_rec = readerv2(0).ToString
                End While
            End If

            readerv2.Close()



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        '---------MAKE SURE TEXTBOXES ARE FILL UP (PART, STATUS AND TYPE) ------------------------
        If String.Equals(mail_t.Text, "") = True Or String.Equals(title_t.Text, "") = True Then
            MessageBox.Show("Please Fill eMail and Title Area")
        Else


            Try
                Dim Create_cmd As New MySqlCommand

                Create_cmd.Parameters.AddWithValue("@Sender", current_user)
                Create_cmd.Parameters.AddWithValue("@role_s", role_me)
                Create_cmd.Parameters.AddWithValue("@Receiver", user_box.Text)
                Create_cmd.Parameters.AddWithValue("@role_r", role_rec)
                Create_cmd.Parameters.AddWithValue("@priority", "n")
                Create_cmd.Parameters.AddWithValue("@date_s", Now)
                Create_cmd.Parameters.AddWithValue("@read_m", "n")
                Create_cmd.Parameters.AddWithValue("@Mail", mail_t.Text)
                Create_cmd.Parameters.AddWithValue("@title", title_t.Text)

                Create_cmd.CommandText = "INSERT INTO management.Dropbox(Sender, role_s, Receiver, role_r, priority, date_s, read_m, Mail, title) VALUES (@Sender,@role_s,@Receiver,@role_r,@priority,@date_s,@read_m,@Mail,@title)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()


                MessageBox.Show("Message sent!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try

        End If

        Me.Visible = False
    End Sub

    Private Sub Sent_mail_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        title_t.Text = ""
        mail_t.Text = ""
    End Sub

    Sub Sent_multiple_emails(team As String, title As String, mail As String)
        'This function sent an email to all the members of a team
        Dim name_of_employee As New List(Of String)()

        Try
            Dim assi_list As New MySqlCommand
            assi_list.Parameters.AddWithValue("@role", team)
            assi_list.CommandText = "SELECT distinct username from users where Role = @role"
            assi_list.Connection = Login.Connection

            Dim reader_l As MySqlDataReader
            reader_l = assi_list.ExecuteReader

            If reader_l.HasRows Then
                While reader_l.Read
                    name_of_employee.Add(reader_l(0))
                End While
            End If

            reader_l.Close()

            'start sending notification

            For i = 0 To name_of_employee.Count - 1
                Dim create_cmd As New MySqlCommand
                create_cmd.Parameters.Clear()
                create_cmd.Parameters.AddWithValue("@Sender", "APL System")
                create_cmd.Parameters.AddWithValue("@role_s", "General Management")
                create_cmd.Parameters.AddWithValue("@Receiver", name_of_employee.Item(i))
                create_cmd.Parameters.AddWithValue("@role_r", team)
                create_cmd.Parameters.AddWithValue("@priority", "Y")
                create_cmd.Parameters.AddWithValue("@date_s", Now)
                create_cmd.Parameters.AddWithValue("@read_m", "n")
                create_cmd.Parameters.AddWithValue("@Mail", mail)
                create_cmd.Parameters.AddWithValue("@title", title)

                create_cmd.CommandText = "INSERT INTO management.Dropbox(Sender, role_s, Receiver, role_r, priority, date_s, read_m, Mail, title) VALUES (@Sender,@role_s,@Receiver,@role_r,@priority,@date_s,@read_m,@Mail,@title)"
                create_cmd.Connection = Login.Connection
                create_cmd.ExecuteNonQuery()

            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub
End Class