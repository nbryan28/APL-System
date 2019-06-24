Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class MPL_eng
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub MPL_eng_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub OpenMPLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenMPLToolStripMenuItem.Click
        My_Material_r.mode = "mpl_eng"
        Open_file.ShowDialog()
    End Sub

    Private Sub ExportMPLToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMPLToExcelToolStripMenuItem.Click
        '--------------- GENERATE A MASTER PACKING LIST -----------------
        If PR_grid.Columns.Count > 1 Then

            Dim count_add_rows As Integer : count_add_rows = 8
            Dim appPath As String = Application.StartupPath()

            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
            Dim xlWorkSheet As Excel.Worksheet
            xlApp.DisplayAlerts = False

            If xlApp Is Nothing Then
                MessageBox.Show("Excel is not properly installed!!")
            Else
                Try
                    ProgressBar1.Visible = True
                    Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\1XXXXX ADA Packing List r2.0.xlsm")
                    xlWorkSheet = wb.Sheets("Master Packing List")


                    For i = 0 To PR_grid.Rows.Count - 1

                        If (ProgressBar1.Value < 400) Then
                            ProgressBar1.Value = ProgressBar1.Value + 10
                        End If


                        xlWorkSheet.Cells(count_add_rows, 1) = PR_grid.Rows(i).Cells(3).Value
                        xlWorkSheet.Cells(count_add_rows, 3) = PR_grid.Rows(i).Cells(2).Value
                        xlWorkSheet.Cells(count_add_rows, 4) = PR_grid.Rows(i).Cells(0).Value
                        xlWorkSheet.Cells(count_add_rows, 5) = PR_grid.Rows(i).Cells(1).Value

                        count_add_rows = count_add_rows + 1
                    Next

                    xlWorkSheet.Range("D:D").HorizontalAlignment = Excel.Constants.xlCenter
                    xlWorkSheet.Range("E:E").HorizontalAlignment = Excel.Constants.xlCenter

                    SaveFileDialog1.Filter = "Excel Files|*.xlsm"

                    If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                        wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                    End If

                    wb.Close(False)

                    Marshal.ReleaseComObject(xlApp)
                    ProgressBar1.Visible = False
                    ProgressBar1.Value = ProgressBar1.Minimum
                    MessageBox.Show("Master Packing List generated successfully!")

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)

                End Try

            End If
        End If
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
End Class