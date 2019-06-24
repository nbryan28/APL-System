Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail


Public Class Job_description
    Private Sub Job_description_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            Dim job As String = Procurement_Overview.job_label.Text

            If Build_mfg.Visible = True Then
                job = Build_mfg.job_label.Text
            End If


            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Job_number", job)
            cmd.CommandText = "SELECT * from management.projects where Job_number = @Job_number"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    job_name.Text = reader(0).ToString
                    desc_job.Text = reader(1).ToString
                    ship_a.Text = reader(2).ToString
                    a_ship.Text = reader(3).ToString
                    onsite.Text = reader(4).ToString
                    CPM.Text = reader(5).ToString
                End While

            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub
End Class