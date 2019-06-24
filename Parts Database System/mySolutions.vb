Imports MySql.Data.MySqlClient

Public Class mySolutions

    Public go_b As Boolean  'avoid datatable to be clear out when populating sol_grid

    Private Sub mySolutions_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'current datatable solutions to grid
        For i = 0 To myQuote.datatable.Rows.Count - 1

            sol_grid.Rows.Add(New String() {myQuote.datatable.Rows(i).Item(0).ToString, myQuote.datatable.Rows(i).Item(1).ToString, "", myQuote.datatable.Rows(i).Item(3).ToString})
            Dim dgvcc As DataGridViewComboBoxCell
            dgvcc = sol_grid.Rows(i).Cells(2)
            dgvcc.Items.Add(myQuote.datatable.Rows(i).Item(2).ToString)
            dgvcc.Items.Add("Solution 2")
            sol_grid.Rows(i).Cells(2).Value = myQuote.datatable.Rows(i).Item(2).ToString

        Next

        go_b = True

    End Sub

    Private Sub sol_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles sol_grid.CellValueChanged
        'update datatable
        If go_b = True Then   'make sure grid is fill for the first time
            myQuote.datatable.Rows.Clear()

            For i = 0 To sol_grid.Rows.Count - 1
                myQuote.datatable.Rows.Add(sol_grid.Rows(i).Cells(0).Value().ToString, sol_grid.Rows(i).Cells(1).Value().ToString, sol_grid.Rows(i).Cells(2).Value().ToString, sol_grid.Rows(i).Cells(3).Value().ToString)
            Next
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        sol_grid.Rows.Clear()
        Dim i As Integer : i = 0

        Try
            '-------- Panel Specific type -------------
            Dim cmd_panel As New MySqlCommand
            cmd_panel.CommandText = "SELECT feature_code, description, solution, solution_description from quote_table.feature_codes where solution_default = 'Y' and type = 'Field'"
            cmd_panel.Connection = Login.Connection
            Dim reader_panel As MySqlDataReader
            reader_panel = cmd_panel.ExecuteReader

            If reader_panel.HasRows Then
                While reader_panel.Read
                    sol_grid.Rows.Add(New String() {reader_panel(0).ToString, reader_panel(1).ToString, "", reader_panel(3).ToString})
                    Dim dgvcc As DataGridViewComboBoxCell
                    dgvcc = sol_grid.Rows(i).Cells(2)
                    dgvcc.Items.Add(reader_panel(2).ToString)
                    dgvcc.Items.Add("PEPPER")
                    sol_grid.Rows(i).Cells(2).Value = reader_panel(2).ToString

                    i = i + 1
                End While
            End If

            reader_panel.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

    End Sub
End Class