Imports MySql.Data.MySqlClient

Public Class MailBox

    Public toggle As Boolean  'turn on/off timer


    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter
        Label3.BackColor = Color.Teal
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave
        Label3.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.BackColor = Color.Teal
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles Label5.MouseEnter
        Label5.BackColor = Color.Teal
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles Label5.MouseLeave
        Label5.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub MailBox_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        toggle = False
        Timer1.Interval = 3000
        Timer1.Start()

        'load messages
        Dim table As New DataTable
        Dim adapter As New MySqlDataAdapter("SELECT Sender, role_s as role, title, read_m, date_s from management.Dropbox where receiver = '" & current_user & "' order by date_s desc", Login.Connection)

        Try
            adapter.Fill(table)     'DataViewGrid1 fill
            dropbox_grid.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill


            dropbox_grid.Columns(0).Width = 200
            dropbox_grid.Columns(1).Width = 200
            dropbox_grid.Columns(2).Width = 780
            dropbox_grid.Columns(3).Width = 80
            dropbox_grid.Columns(4).Width = 300

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Sent_mail.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If Me.Visible = True And toggle = True Then
            Try
                Dim table As New DataTable
                Dim adapter As New MySqlDataAdapter("SELECT Sender, role_s as role, title, read_m, date_s from management.Dropbox where receiver = '" & current_user & "' order by date_s desc", Login.Connection)


                adapter.Fill(table)     'DataViewGrid1 fill
                dropbox_grid.DataSource = table
                '   Form1.Specs_Table.DataSource = table   'Specifications table fill


                dropbox_grid.Columns(0).Width = 200
                dropbox_grid.Columns(1).Width = 200
                dropbox_grid.Columns(2).Width = 780
                dropbox_grid.Columns(3).Width = 80
                dropbox_grid.Columns(4).Width = 300

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub dropbox_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dropbox_grid.RowsAdded
        For i = 0 To dropbox_grid.Rows.Count - 1

            If String.Equals(dropbox_grid.Rows(i).Cells(3).Value, "n") = True Then
                dropbox_grid.Rows(i).DefaultCellStyle.BackColor = Color.Peru
            Else
                dropbox_grid.Rows(i).DefaultCellStyle.BackColor = Color.WhiteSmoke
            End If
        Next
    End Sub

    Private Sub dropbox_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles dropbox_grid.RowsRemoved
        For i = 0 To dropbox_grid.Rows.Count - 1

            If String.Equals(dropbox_grid.Rows(i).Cells(3).Value, "n") = True Then
                dropbox_grid.Rows(i).DefaultCellStyle.BackColor = Color.Peru
            Else
                dropbox_grid.Rows(i).DefaultCellStyle.BackColor = Color.WhiteSmoke
            End If
        Next
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click


        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Sender, role_s as role, title, read_m, date_s from management.Dropbox where receiver = '" & current_user & "' order by date_s desc", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            dropbox_grid.DataSource = table

            dropbox_grid.Columns(0).Width = 200
            dropbox_grid.Columns(1).Width = 200
            dropbox_grid.Columns(2).Width = 780
            dropbox_grid.Columns(3).Width = 80
            dropbox_grid.Columns(4).Width = 300

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        toggle = True
        Label6.Text = "Refresh on"
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        toggle = False

        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT receiver, role_r as role, title, read_m, date_s from management.Dropbox where sender = '" & current_user & "' order by date_s desc", Login.Connection)

            adapter.Fill(table)     'DataViewGrid1 fill
            dropbox_grid.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            dropbox_grid.Columns(0).Width = 200
            dropbox_grid.Columns(1).Width = 200
            dropbox_grid.Columns(2).Width = 780
            dropbox_grid.Columns(3).Width = 80
            dropbox_grid.Columns(4).Width = 300

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub dropbox_grid_DoubleClick(sender As Object, e As EventArgs) Handles dropbox_grid.DoubleClick

        Dim index As Integer : index = 0

        index = dropbox_grid.CurrentRow.Index

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@sender", dropbox_grid.Rows(index).Cells(0).Value)
            cmd4.Parameters.AddWithValue("@role_s", dropbox_grid.Rows(index).Cells(1).Value)
            cmd4.Parameters.AddWithValue("@title", dropbox_grid.Rows(index).Cells(2).Value)
            cmd4.Parameters.AddWithValue("@read_m", dropbox_grid.Rows(index).Cells(3).Value)
            cmd4.Parameters.AddWithValue("@date_s", dropbox_grid.Rows(index).Cells(4).Value)
            cmd4.Parameters.AddWithValue("@receiver", current_user)
            cmd4.CommandText = "SELECT Mail from management.Dropbox where Sender = @sender and role_s = @role_s and title = @title and read_m = @read_m and date_s = @date_s and Receiver = @receiver"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    TextBox1.Text = reader4(0).ToString
                End While
            End If

            reader4.Close()

            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@sender", dropbox_grid.Rows(index).Cells(0).Value)
            Create_cmd.Parameters.AddWithValue("@role_s", dropbox_grid.Rows(index).Cells(1).Value)
            Create_cmd.Parameters.AddWithValue("@title", dropbox_grid.Rows(index).Cells(2).Value)
            Create_cmd.Parameters.AddWithValue("@read_m", dropbox_grid.Rows(index).Cells(3).Value)
            Create_cmd.Parameters.AddWithValue("@date_s", dropbox_grid.Rows(index).Cells(4).Value)
            Create_cmd.Parameters.AddWithValue("@receiver", current_user)

            Create_cmd.CommandText = "UPDATE management.Dropbox SET read_m = 'Y' where Sender = @sender and role_s = @role_s and title = @title and date_s = @date_s and Receiver = @receiver"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        '--------- priority ---------

        toggle = False

        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Sender, role_s as role, title, read_m, date_s from management.Dropbox where receiver = '" & current_user & "' and priority = 'Y' order by date_s desc", Login.Connection)

            adapter.Fill(table)     'DataViewGrid1 fill
            dropbox_grid.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            dropbox_grid.Columns(0).Width = 200
            dropbox_grid.Columns(1).Width = 200
            dropbox_grid.Columns(2).Width = 780
            dropbox_grid.Columns(3).Width = 80
            dropbox_grid.Columns(4).Width = 300

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        '-- refresn on/off
        If String.Equals(Label6.Text, "Refresh on") = True Then

            Label6.Text = "Refresh off"
            toggle = False
        Else
            Label6.Text = "Refresh on"
            toggle = True
        End If
    End Sub
End Class