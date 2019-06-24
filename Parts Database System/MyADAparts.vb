Imports MySql.Data.MySqlClient

Public Class Visualizer

    Public datatable As DataTable

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If IsNumeric(TextBox1.Text) = True Then

            For i = 0 To part_assembly.Rows.Count - 1
                part_assembly.Rows(i).Cells(1).Value = datatable.Rows(i).Item(0) * TextBox1.Text
                part_assembly.Rows(i).Cells(7).Value = datatable.Rows(i).Item(1) * TextBox1.Text
            Next

        End If

    End Sub

    Private Sub part_assembly_VisibleChanged(sender As Object, e As EventArgs) Handles part_assembly.VisibleChanged

        datatable = New DataTable
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("cost", GetType(String))

        For i = 0 To part_assembly.Rows.Count - 1
            datatable.Rows.Add(part_assembly.Rows(i).Cells(1).Value, part_assembly.Rows(i).Cells(7).Value)

            '--- check if its in inventory
            Dim cmd5 As New MySqlCommand
            cmd5.Parameters.Clear()
            cmd5.Parameters.AddWithValue("@part_name", part_assembly.Rows(i).Cells(2).Value)
            cmd5.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
            cmd5.Connection = Login.Connection
            Dim reader5 As MySqlDataReader
            reader5 = cmd5.ExecuteReader

            If reader5.HasRows Then
                While reader5.Read
                    part_assembly.Rows(i).Cells(8).Value = "Yes"
                End While
            Else
                part_assembly.Rows(i).Cells(8).Value = "No"
            End If

            reader5.Close()


        Next
    End Sub



    Private Sub part_assembly_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles part_assembly.CellValueChanged

        Dim temp As Double : temp = 0
        For Each row As DataGridViewRow In part_assembly.Rows
            If row.IsNewRow Then Continue For
            If IsNumeric(row.Cells(7).Value) = True Then
                temp = temp + row.Cells(7).Value
            End If
        Next

        mat_l.Text = "Material Cost $ " & temp
    End Sub

    Private Sub Visualizer_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class