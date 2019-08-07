Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient

Public Class MPL_mfg
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ExportMPLToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMPLToExcelToolStripMenuItem.Click
        'export  MPL to excel file

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
                xlWorkSheet.Range("D:E").ColumnWidth = 20
                xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To PR_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = PR_grid.Columns(i).HeaderText
                    For j As Integer = 0 To PR_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = PR_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Master_Packing_List.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Master Packing List exported successfully!")
            End If
        End If

    End Sub

    Private Sub OpenMPLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenMPLToolStripMenuItem.Click

        My_Material_r.mode = "mpl_mfg"
        Open_file.ShowDialog()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In PR_grid.Rows
            If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
                If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text) = 0 Then
                    rowindex = row.Index
                    PR_grid.CurrentCell = PR_grid.Rows(rowindex).Cells(0)
                    found_po = True
                    Exit For
                End If

            End If
        Next

        If found_po = False Then
            MsgBox("Part not found!")
        End If
    End Sub

    Private Sub MPL_mfg_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub



End Class