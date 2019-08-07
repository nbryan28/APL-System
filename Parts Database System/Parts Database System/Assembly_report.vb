Imports MySql.Data.MySqlClient

Public Class Assembly_report
    'Private Sub Label2_Click(sender As Object, e As EventArgs)
    '    If Log_out = True Then
    '        Me.Visible = False
    '        MyDash.Visible = True
    '    Else
    '        Me.Visible = False
    '    End If
    'End Sub

    Private Sub Assembly_report_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim got_assemblies = New List(Of String)()

        Try

            '------- populate assemblies list -------

            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    got_assemblies.Add(reader2(0))
                End While
            End If

            reader2.Close()
            '------------------------------------------

            '--- populate asm_grid with assemblies needed from material order

            For i = 0 To got_assemblies.Count - 1

                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@part_name", got_assemblies.Item(i).ToString)
                cmd3.CommandText = "SELECT Part_No, Qty_needed from inventory.Material_orders where Part_No = @part_name"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        asm_grid.Rows.Add(New String() {})
                        asm_grid.Rows(asm_grid.Rows.Count - 1).Cells(0).Value = reader3(0).ToString
                        asm_grid.Rows(asm_grid.Rows.Count - 1).Cells(1).Value = reader3(1).ToString
                    End While
                End If

                reader3.Close()

            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub asm_grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles asm_grid.CellContentClick
        Dim senderGrid = DirectCast(sender, DataGridView)

        If TypeOf senderGrid.Columns(e.ColumnIndex) Is DataGridViewButtonColumn AndAlso
           e.RowIndex >= 0 Then
            MessageBox.Show(asm_grid.Rows(e.RowIndex).Cells(0).Value)
        End If
    End Sub
End Class