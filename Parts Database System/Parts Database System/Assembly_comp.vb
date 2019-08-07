Imports MySql.Data.MySqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection


Public Class Assembly_comp

    '------------------- datatable assemblies
    Public table_assem As DataTable

    '-------------- datatable adv
    Public table_adv As DataTable


    Private Sub Assembly_comp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '---- assem
        table_assem = New DataTable
        table_assem.Columns.Add("Legacy_ADA_Number", GetType(String))
        table_assem.Columns.Add("Labor_Cost", GetType(Decimal))
        table_assem.Columns.Add("Bulk_Cost", GetType(Decimal))


        '----- adv
        table_adv = New DataTable
        table_adv.Columns.Add("Part_Name", GetType(String))
        table_adv.Columns.Add("ADV_Number", GetType(String))
        table_adv.Columns.Add("Qty", GetType(Double))
        table_adv.Columns.Add("Legacy_ADA", GetType(String))

        Try
            '----------- assem
            Dim cmd8 As New MySqlCommand
            cmd8.CommandText = "SELECT Legacy_ADA_Number, Labor_Cost, Bulk_Cost  from devices"
            cmd8.Connection = Login.Connection
            Dim reader8 As MySqlDataReader
            reader8 = cmd8.ExecuteReader

            If reader8.HasRows Then
                While reader8.Read
                    table_assem.Rows.Add(reader8(0), reader8(1), reader8(2))
                End While
            End If

            reader8.Close()

            '---------- adv
            Dim cmd9 As New MySqlCommand
            cmd9.CommandText = "SELECT Part_Name, ADV_Number, Qty, Legacy_ADA from adv"
            cmd9.Connection = Login.Connection
            Dim reader9 As MySqlDataReader
            reader9 = cmd9.ExecuteReader

            If reader9.HasRows Then
                While reader9.Read
                    table_adv.Rows.Add(reader9(0), reader9(1), reader9(2), reader9(3))
                End While
            End If

            reader9.Close()

            '---------------------------------

            '------------------  enter parts ------------------------
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select distinct Part_Name from adv;"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    Panel_grid.Rows.Add(New String() {})
                    Panel_grid.Rows(i).Cells(0).Value = reader(0).ToString
                    i = i + 1
                End While
            End If

            reader.Close()

            '---------- enter description ------------
            For i = 0 To Panel_grid.Rows.Count - 1

                Dim check_cmd2 As New MySqlCommand
                check_cmd2.Parameters.Clear()
                check_cmd2.Parameters.AddWithValue("@part", Panel_grid.Rows(i).Cells(0).Value)
                check_cmd2.CommandText = "select Part_Description from parts_table where Part_Name = @part"

                check_cmd2.Connection = Login.Connection
                check_cmd2.ExecuteNonQuery()

                Dim reader2 As MySqlDataReader
                reader2 = check_cmd2.ExecuteReader

                If reader2.HasRows Then
                    While reader2.Read
                        Panel_grid.Rows(i).Cells(1).Value = reader2(0).ToString
                    End While
                End If

                reader2.Close()
            Next

            '------------------------------------------


            '--- enter assembly columns -----

            For i = 0 To table_assem.Rows.Count - 1
                Panel_grid.Columns.Add(table_assem.Rows(i).Item(0).ToString, table_assem.Rows(i).Item(0).ToString)
            Next

            For i = 2 To Panel_grid.Columns.Count - 1
                Panel_grid.Columns(i).Width = 290
                Panel_grid.Columns(i).HeaderCell.Style.BackColor = Color.DarkCyan
                Panel_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            '-------------------------------------

            '------ get qty of that specific part --------------
            For i = 0 To Panel_grid.Rows.Count - 1
                For j = 2 To Panel_grid.Columns.Count - 1

                    'Dim check_cmd2 As New MySqlCommand
                    'check_cmd2.Parameters.Clear()
                    'check_cmd2.Parameters.AddWithValue("@part", Panel_grid.Rows(i).Cells(0).Value)
                    'check_cmd2.Parameters.AddWithValue("@assem", Panel_grid.Columns.Item(j).HeaderText)
                    'check_cmd2.CommandText = "select count(Qty) from adv where Part_Name = @part and Legacy_ADA = @assem"

                    'check_cmd2.Connection = Login.Connection
                    'check_cmd2.ExecuteNonQuery()

                    'Dim reader2 As MySqlDataReader
                    'reader2 = check_cmd2.ExecuteReader

                    'If reader2.HasRows Then
                    '    While reader2.Read
                    '        Panel_grid.Rows(i).Cells(j).Value = reader2(0).ToString
                    '    End While
                    'End If

                    'reader2.Close()
                    Dim count_i As Integer : count_i = i
                    Dim count_j As Integer : count_j = j
                    Dim count_as = (From feature_c In table_adv Where feature_c.Field(Of String)("Part_Name") = Panel_grid.Rows(count_i).Cells(0).Value And feature_c.Field(Of String)("Legacy_ADA") = Panel_grid.Columns.Item(count_j).HeaderText Select feature_c).Count
                    Panel_grid.Rows(i).Cells(j).Value = count_as
                Next
            Next




            Call EnableDoubleBuffered(Panel_grid)

            Me.DoubleBuffered = True

            Panel_grid.Rows.Insert(0)
            Panel_grid.Rows(0).Cells(0).Value = ""
            Panel_grid.Rows(0).DefaultCellStyle.BackColor = Color.CadetBlue
            Panel_grid.Rows(0).ReadOnly = False
            '---------------------------------


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
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
                xlWorkSheet.Range("A:BB").ColumnWidth = 40
                xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To Panel_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = Panel_grid.Columns(i).HeaderText
                    For j As Integer = 0 To Panel_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = Panel_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Dim valid_mr_name As String = Me.Text

                For Each c In Path.GetInvalidFileNameChars()
                    If valid_mr_name.Contains(c) Then
                        valid_mr_name = valid_mr_name.Replace(c, "")
                    End If
                Next

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\" & valid_mr_name & ".xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Table exported successfully!")
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        If Panel_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(Panel_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub
End Class