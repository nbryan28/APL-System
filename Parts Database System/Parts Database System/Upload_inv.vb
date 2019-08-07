Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail

Public Class Upload_inv
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*ods;"

            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

                load_grid.Rows.Clear()

                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)

                Dim i As Integer : i = 2

                While (wb.ActiveSheet.Cells(i, 1).Value IsNot Nothing)

                    If String.Equals(ComboBox1.Text, "Add Parts Qty") = False Then
                        Call update_inv(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 2).value)
                    Else
                        Call update_add(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 2).value)
                    End If
                    i = i + 1

                End While
                '---------------------------------------------
                wb.Close(False)

                Call Inventory_manage.refresh_table()
                Call Inventory_manage.General_inv_cal()


            End If
        End If

    End Sub

    Sub update_inv(part As String, qty As String)

        Try
            Dim exist_c As Boolean : exist_c = False

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part)
            cmd4.CommandText = "SELECT * from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    exist_c = True
                End While
            End If

            reader4.Close()

            If exist_c = False Then
                load_grid.Rows.Add(New String() {part, 0, If(IsNumeric(qty) = True, qty, 0), "Not Found"})
            Else

                '------ update qty ----------
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@part_name", part)
                Create_cmd.Parameters.AddWithValue("@current_qty", qty)
                Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()


                '----------------------------

                load_grid.Rows.Add(New String() {part, 0, If(IsNumeric(qty) = True, qty, 0), "Found"})
            End If

            Call refresh_inv()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Sub update_add(part As String, qty As String)

        Try
            Dim exist_c As Boolean : exist_c = False
            Dim c_q As Double : c_q = 0

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part_name", part)
            cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    exist_c = True
                    c_q = reader4(0)
                End While
            End If

            reader4.Close()

            If exist_c = False Then
                load_grid.Rows.Add(New String() {part, 0, If(IsNumeric(qty) = True, qty, 0), "Not Found"})
            Else

                '------ add qty ----------
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@part_name", part)
                Create_cmd.Parameters.AddWithValue("@current_qty", c_q + qty)

                Create_cmd.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty  where part_name = @part_name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()


                '----------------------------


                load_grid.Rows.Add(New String() {part, 0, If(IsNumeric(qty) = True, qty, 0), "Found"})
            End If

            Call refresh_inv()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Upload_inv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Add Parts Qty"
        load_grid.Rows.Clear()
    End Sub

    Sub refresh_inv()

        Try
            For i = 0 To load_grid.Rows.Count - 1

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.Clear()
                cmd4.Parameters.AddWithValue("@part_name", load_grid.Rows(i).Cells(0).Value)
                cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        load_grid.Rows(i).Cells(1).Value = reader4(0).ToString
                    End While
                End If

                reader4.Close()

                If String.Equals(load_grid.Rows(i).Cells(3).Value, "Not Found") = True Then
                    load_grid.Rows(i).DefaultCellStyle.BackColor = Color.DarkOrange
                End If

            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
                For i As Integer = 0 To load_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = load_grid.Columns(i).HeaderText
                    For j As Integer = 0 To load_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = load_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Inventory_summary.xlsx")
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