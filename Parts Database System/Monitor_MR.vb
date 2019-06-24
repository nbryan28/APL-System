Imports MySql.Data.MySqlClient

Public Class Monitor_MR

    Public name_t As String

    Private Sub Monitor_MR_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Text = "All MFG Types"
        total_grid.Rows.Clear()

        name_t = Project_dash.mr_box.Text

        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", Project_dash.mr_box.Text)
            cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, mfg_type,  ADA_Number,  Vendor, Price, Qty, subtotal, qty_fullfilled, qty_needed, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    total_grid.Rows.Add(New String() {})
                    total_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    total_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    total_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    total_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                    total_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                    total_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                    total_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                    total_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                    total_grid.Rows(i).Cells(8).Value = If(IsNumeric(reader4(6)) = True, reader4(6).ToString, 0) * If(IsNumeric(reader4(7)) = True, reader4(7).ToString, 0)
                    total_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                    total_grid.Rows(i).Cells(10).Value = reader4(10).ToString

                    i = i + 1
                End While

            End If

            reader4.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub total_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles total_grid.CellValueChanged
        For Each row As DataGridViewRow In total_grid.Rows
            If row.IsNewRow Then Continue For

            If (IsNumeric(row.Cells(7).Value) = True And IsNumeric(row.Cells(9).Value)) Then

                row.Cells(10).Value = CType(row.Cells(7).Value, Double) - CType(row.Cells(9).Value, Double)

                If CType(row.Cells(10).Value, Double) > 0 Then
                    row.Cells(10).Style.BackColor = Color.Firebrick
                End If

            End If
        Next

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        Try
            total_grid.Rows.Clear()
            Dim cmd4 As New MySqlCommand

            cmd4.Parameters.AddWithValue("@mr_name", name_t)
            cmd4.Parameters.AddWithValue("@mfg_type", ComboBox1.Text)
            Dim query_m As String : query_m = "SELECT Part_No, Description, Manufacturer, mfg_type,  ADA_Number,  Vendor, Price, Qty, subtotal, qty_fullfilled, qty_needed, need_by_date  from Material_Request.mr_data where mr_name = @mr_name"

            If String.Equals("All MFG Types", ComboBox1.Text) = False Then
                query_m = query_m & " and mfg_type = @mfg_type"
            End If

            cmd4.CommandText = query_m
            cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    Dim i As Integer : i = 0
                While reader4.Read
                    total_grid.Rows.Add(New String() {})
                    total_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    total_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    total_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    total_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                    total_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                    total_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                    total_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                    total_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                    total_grid.Rows(i).Cells(8).Value = If(IsNumeric(reader4(6)) = True, reader4(6).ToString, 0) * If(IsNumeric(reader4(7)) = True, reader4(7).ToString, 0)
                    total_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                    total_grid.Rows(i).Cells(10).Value = reader4(10).ToString

                    i = i + 1
                End While

            End If

                reader4.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

    End Sub
End Class