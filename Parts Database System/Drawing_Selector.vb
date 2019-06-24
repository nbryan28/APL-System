Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Drawing_Selector
    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles pdf_grid.DoubleClick

        'Dim index_row As Integer = pdf_grid.CurrentRow.Index
        'Dim file_n As String = pdf_grid.Rows(index_row).Cells(1).Value


        'If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\" & file_n) = True Then
        '    System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\" & file_n)
        '    Me.Visible = False
        'Else

        '    MessageBox.Show("No PDF Drawing Found")

        'End If
    End Sub

    Private Sub Drawing_Selector_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'pdf_grid.Rows.Clear()

        'Try
        '    Dim cmd4 As New MySqlCommand
        '    cmd4.Parameters.AddWithValue("@mr_name", Name)
        '    cmd4.CommandText = "SELECT draw_description, file_name, date_upload, uploaded_by, job from management.draw_pdf"
        '    cmd4.Connection = Login.Connection
        '    Dim reader4 As MySqlDataReader
        '    reader4 = cmd4.ExecuteReader

        '    If reader4.HasRows Then
        '        While reader4.Read
        '            pdf_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString})
        '        End While
        '    End If

        '    reader4.Close()
        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try
    End Sub
End Class