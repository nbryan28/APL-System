Imports MySql.Data.MySqlClient

Public Class FMR_form
    Private Sub FMR_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'fill job combobox
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    job_com.Items.Add(readerj(0))
                End While
            End If

            readerj.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Part_selector.ShowDialog()
    End Sub
End Class