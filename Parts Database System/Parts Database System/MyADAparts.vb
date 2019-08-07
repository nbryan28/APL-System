Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Reflection
Imports System.Runtime.InteropServices

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



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'export Inventory to excel file
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:B").ColumnWidth = 40
                xlWorkSheet.Range("C:C").ColumnWidth = 30
                xlWorkSheet.Range("D:L").ColumnWidth = 20
                xlWorkSheet.Range("A:L").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To part_assembly.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = part_assembly.Columns(i).HeaderText
                    For j As Integer = 0 To part_assembly.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = part_assembly.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Assembly_composition.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Inventory Summary exported successfully!")
            End If
        End If
    End Sub
End Class