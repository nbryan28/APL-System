Imports Microsoft.Office.Interop

Public Class BOM
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        Try
            If e.ColumnIndex = 4 Or e.ColumnIndex = 5 Then

                If IsNumeric(DataGridView1.CurrentRow.Cells("Column6").Value) And IsNumeric(DataGridView1.CurrentRow.Cells("Column5").Value) Then

                    DataGridView1.CurrentRow.Cells("Column7").Value = CType(DataGridView1.CurrentRow.Cells("Column6").Value, Integer) * CType(DataGridView1.CurrentRow.Cells("Column5").Value, Decimal)
                Else
                    DataGridView1.CurrentRow.Cells("Column7").Value = 0
                End If

            End If
        Catch ex As Exception
        End Try

        Call Total_sum()

    End Sub


    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()


        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:D").ColumnWidth = 48
                xlWorkSheet.Range("A:G").HorizontalAlignment = Excel.Constants.xlCenter

                'copy data to excel
                For i As Integer = 0 To DataGridView1.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = DataGridView1.Columns(i).HeaderText
                    For j As Integer = 0 To DataGridView1.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = DataGridView1.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Bill_Of_Material.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                'xlApp.Quit()

                MessageBox.Show("Bill Of Material Created Succesfully!")
            End If
        End If

    End Sub

    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        'Clear datagridview
        DataGridView1.Rows.Clear()
        Call Total_sum()
    End Sub

    Private Sub DeleteRowsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRowsToolStripMenuItem.Click
        '==================== DELETE ROWS BUTTON ===========================

        If DataGridView1.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(r)
            Next

            Call Total_sum()
        Else
            MessageBox.Show("Please, select at least one row")
        End If
    End Sub

    Sub Total_sum()

        'Calculate total cost BOM

        Dim sum_t As Double : sum_t = 0

        For i = 0 To DataGridView1.Rows.Count - 1
            sum_t = sum_t + DataGridView1.Rows(i).Cells(6).Value()
        Next

        Label1.Text = "Total: " & sum_t

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ' ----------------  Open BOM Browser -------------------

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Visible = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub BOM_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class