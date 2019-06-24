Imports MySql.Data.MySqlClient

Public Class User_form

    Public mode As String
    Public temp_user As String


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub User_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        mode = ""
        temp_user = "xxxxxxxa"
        '-------- LOAD USER FORM -----------
        Try
            Dim tabl As New DataTable
            Dim ad As New MySqlDataAdapter("Select username, Role, email, reg_date from users", Login.Connection)
            ad.Fill(tabl)
            User_grid.DataSource = tabl

            'Setting Columns size 
            For i = 0 To User_grid.ColumnCount - 1
                With User_grid.Columns(i)
                    .Width = 580
                End With
            Next i



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Refresh_data(c As MySqlConnection)

        '---------- REFRESH DATA IN GRID ----------------------
        Try
            Dim tabl2 As New DataTable
            Dim ad2 As New MySqlDataAdapter("Select username, Role, email, reg_date from users", c)
            ad2.Fill(tabl2)
            User_grid.DataSource = tabl2
        Catch ex As Exception
        End Try

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Create_b.Click

        '--------------ADD USER -----------------

        ''---------MAKE SURE TEXTBOXES ARE FILL------------------------
        'If String.Equals(User_c.Text, "") = True Or String.Equals(Roles_c.Text, "") = True Then
        '    MessageBox.Show("Please Fill Username and Roles fields")
        'Else

        '    '----------- MAKE SURE THE RECORD DOES NOT EXIST ------------
        '    Try
        '        Dim cmd As New MySqlCommand
        '        Dim Create_cmd As New MySqlCommand
        '        Dim active_user As New MySqlCommand

        '        cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
        '        cmd.CommandText = "SELECT username from users where username = @User_Name"
        '        cmd.Connection = Login.Connection
        '        Dim reader As MySqlDataReader
        '        reader = cmd.ExecuteReader

        '        If Not reader.HasRows Then

        '            'IF IT DOES NOT EXIST THEN
        '            reader.Close()
        '            Create_cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
        '            ' Create_cmd.Parameters.AddWithValue("@Password", Pass_c.Text)
        '            Create_cmd.Parameters.AddWithValue("@Role", Roles_c.Text)
        '            Create_cmd.Parameters.AddWithValue("@Email", email_c.Text)

        '            active_user.Parameters.AddWithValue("@User_Name", User_c.Text)
        '            active_user.Parameters.AddWithValue("@Role", Roles_c.Text)

        '            Create_cmd.CommandText = "INSERT INTO users(username, Role, email, reset_t) VALUES (@User_Name,  @Role , @Email, '' )"
        '            Create_cmd.Connection = Login.Connection
        '            Create_cmd.ExecuteNonQuery()

        '            'insert info in status_user table
        '            active_user.CommandText = "INSERT INTO management.status_user(User, Role, Status , Log_in_time) VALUES (@User_Name, @Role ,'Not Active' , now() )"
        '            active_user.Connection = Login.Connection
        '            active_user.ExecuteNonQuery()

        '            MessageBox.Show("Record entered succesfully")
        '            Call Refresh_data(Login.Connection)

        '        Else
        '            MessageBox.Show("User already exist!")
        '            'DO NOT FORGET TO CLOSE THE READER
        '            reader.Close()
        '        End If

        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)

        '    End Try
        'End If
        mode = "add"
        Add_password.ShowDialog()
        '  Call add_username()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '--------------SEARCHING USER STRING------------------
        Dim found_user As Boolean : found_user = False
        Dim rowindex As Integer
        For Each row As DataGridViewRow In User_grid.Rows
            If String.Compare(row.Cells.Item("username").Value.ToString, TextBox1.Text) = 0 Then
                rowindex = row.Index
                User_grid.CurrentCell = User_grid.Rows(rowindex).Cells(0)
                found_user = True
                Exit For
            End If
        Next
        If found_user = False Then
            MsgBox("User not found")
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '-----------UPDATE USER TABLE --------------------

        If ok_press = True Then
            Try
                Dim selectedItem As Object
                Dim query_t As String
                Dim Type_name As String
                selectedItem = ComboBox1.SelectedItem
                Type_name = selectedItem.ToString

                If String.Equals(Type_name, "All Roles") = True Then
                    query_t = ""
                Else
                    query_t = "and Role = '" & Type_name & "'"
                End If

                Dim ta As New DataTable
                Dim adapt As New MySqlDataAdapter("SELECT username, Role, email, reg_date from users where 1=1 " & query_t, Login.Connection)
                adapt.Fill(ta)
                User_grid.DataSource = ta

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub


    Private Sub User_grid_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles User_grid.CellContentDoubleClick

        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW

        Dim username_t = user_grid.CurrentCell.Value.ToString
        temp_user = user_grid.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@User_Name", username_t)
            ' cmd.CommandText = "SELECT * from users where username = """ & username_t & """"
            cmd.CommandText = "SELECT username, Role, email from users where username = @User_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    User_c.Text = reader(0).ToString
                    ' Pass_c.Text = reader(2).ToString
                    Roles_c.Text = reader(1).ToString
                    email_c.Text = reader(2).ToString

                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '------------ UPDATE RECORD----------------


        If String.Equals(User_c.Text, "") = True Or String.Equals(Roles_c.Text, "") = True Then
            MessageBox.Show("Please Fill Username and Roles fields")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim active_user As New MySqlCommand

                cmd.Parameters.AddWithValue("@User_Name", temp_user)
                cmd.CommandText = "SELECT username from users where username = @User_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                'MAKE SURE THE RECORD EXIST SO IT CAN BE UPDATED
                If reader.HasRows Then

                    reader.Close()
                    Create_cmd.Parameters.AddWithValue("@User_Name", temp_user)
                    Create_cmd.Parameters.AddWithValue("@n_User_Name", User_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Role", Roles_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Email", email_c.Text)
                    Create_cmd.CommandText = "UPDATE users  SET username = @n_User_Name, Role = @Role, email = @Email where username = @User_Name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    'update status_user table

                    active_user.Parameters.AddWithValue("@User_Name", temp_user)
                    active_user.Parameters.AddWithValue("@nUser_Name", User_c.Text)
                    active_user.Parameters.AddWithValue("@Role", Roles_c.Text)
                    active_user.CommandText = "UPDATE management.status_user  SET User = @nUser_Name,  Role = @Role where User = @User_Name"
                    active_user.Connection = Login.Connection
                    active_user.ExecuteNonQuery()

                    MessageBox.Show("Record updated succesfully")
                    Call Refresh_data(Login.Connection)

                Else
                    MessageBox.Show("User does not exist! ... Update incomplete")
        reader.Close()
        End If

        Catch ex As Exception
        MessageBox.Show(ex.ToString)

        End Try



        End If
        '  mode = "update"
        ' Add_password.ShowDialog()
        ' Call update_user()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Call delete_username()

    End Sub


    Private Sub User_form_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        Dim ctrl As Control
        For Each ctrl In Me.GroupBox1.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next

        For Each ctrl In Me.GroupBox2.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next
    End Sub


    Sub add_username()
        '--------------ADD USER -----------------

        '---------MAKE SURE TEXTBOXES ARE FILL------------------------
        If String.Equals(User_c.Text, "") = True Or String.Equals(Roles_c.Text, "") = True Then
            MessageBox.Show("Please Fill Username and Roles fields")
        Else

            '----------- MAKE SURE THE RECORD DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim active_user As New MySqlCommand

                cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                cmd.CommandText = "SELECT username from users where username = @User_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN
                    reader.Close()
                    Create_cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Password", Add_password.passbox.Text)
                    Create_cmd.Parameters.AddWithValue("@Role", Roles_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Email", email_c.Text)

                    active_user.Parameters.AddWithValue("@User_Name", User_c.Text)
                    active_user.Parameters.AddWithValue("@Role", Roles_c.Text)

                    Create_cmd.CommandText = "INSERT INTO users(username, password, Role, email, reset_t) VALUES (@User_Name, @Password, @Role , @Email, '' )"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    'insert info in status_user table
                    active_user.CommandText = "INSERT INTO management.status_user(User, Role, Status , Log_in_time) VALUES (@User_Name, @Role ,'Not Active' , now() )"
                    active_user.Connection = Login.Connection
                    active_user.ExecuteNonQuery()

                    MessageBox.Show("Record entered succesfully")
                    Call Refresh_data(Login.Connection)

                Else
                    MessageBox.Show("User already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Sub update_username()
        '------------ UPDATE RECORD----------------


        If String.Equals(User_c.Text, "") = True Or String.Equals(Roles_c.Text, "") = True Then
            MessageBox.Show("Please Fill Username and Roles fields")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim active_user As New MySqlCommand

                cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                cmd.CommandText = "SELECT username from users where username = @User_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                'MAKE SURE THE RECORD EXIST SO IT CAN BE UPDATED
                If reader.HasRows Then

                    reader.Close()
                    Create_cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Password", Add_password.passbox.Text)
                    Create_cmd.Parameters.AddWithValue("@Role", Roles_c.Text)
                    Create_cmd.Parameters.AddWithValue("@Email", email_c.Text)
                    Create_cmd.CommandText = "UPDATE users  SET username = @User_Name, password = @Password, Role = @Role, email = @Email where username = @User_Name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    'update status_user table

                    active_user.Parameters.AddWithValue("@User_Name", User_c.Text)
                    active_user.Parameters.AddWithValue("@Role", Roles_c.Text)
                    active_user.CommandText = "UPDATE management.status_user  SET User = @User_Name,  Role = @Role where User = @User_Name"
                    active_user.Connection = Login.Connection
                    active_user.ExecuteNonQuery()

                    MessageBox.Show("Record updated succesfully")
                    Call Refresh_data(Login.Connection)

                Else
                    MessageBox.Show("User does not exist! ... Update incomplete")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try



        End If
    End Sub

    Sub delete_username()

        '-----------DELETE RECORD --------------------
        If String.Equals(User_c.Text, "") = True Or String.Equals(Roles_c.Text, "") = True Then
            MessageBox.Show("Please Fill Username, and Roles fields")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim active_user As New MySqlCommand

                cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                cmd.CommandText = "SELECT username from users where username = @User_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    Dim dlgR As DialogResult
                    dlgR = MessageBox.Show("Are you sure you want to delete this record?", "Attention!", MessageBoxButtons.YesNo)
                    If dlgR = DialogResult.Yes Then

                        reader.Close()
                        Create_cmd.Parameters.AddWithValue("@User_Name", User_c.Text)
                        Create_cmd.CommandText = "DELETE FROM users where username = @User_Name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                        active_user.Parameters.AddWithValue("@User_Name", User_c.Text)
                        active_user.CommandText = "DELETE FROM management.status_user where user = @User_Name"
                        active_user.Connection = Login.Connection
                        active_user.ExecuteNonQuery()


                        MessageBox.Show("Record deleted succesfully")
                        Call Refresh_data(Login.Connection)
                    Else
                        reader.Close()

                    End If
                Else
                    MessageBox.Show("We could not find any record to be removed")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        '-- reset pass
        Try
            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@username", User_c.Text)
            Create_cmd.CommandText = "UPDATE users SET reset_t = 'APL' where username = @username"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()
            MessageBox.Show("Password was reset")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class