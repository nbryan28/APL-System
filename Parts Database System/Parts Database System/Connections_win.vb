Imports MySql.Data.MySqlClient

Public Class Connections_win
    Private Sub Connections_win_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Fill user
        Dim table As New DataTable
        Dim adapter As New MySqlDataAdapter("SELECT * from management.status_user where status = 'Active'", Login.Connection)

        adapter.Fill(table)     'DataViewGrid1 fill
        active_user_grid.DataSource = table

        For i = 1 To active_user_grid.ColumnCount - 1
            With active_user_grid.Columns(i)
                .Width = 300
            End With
        Next i

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub
End Class