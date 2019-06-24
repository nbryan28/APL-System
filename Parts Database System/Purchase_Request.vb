Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class Purchase_Request

    Public row_in As Integer ' this variable keep the row index of the PR grid selected when (Find Alternative part option is used)
    Public myQTY As Double 'store PR qty temporarily
    Public datatable As DataTable  'store part and qty in order to get description and vendor from parts table
    Public whatipress As Integer


    Private Sub ToolStripMenuItem30_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem30.Click
        If PR_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In PR_grid.SelectedRows
                PR_grid.Rows.Remove(r)
            Next
            Call Total_rows()
        Else
            MessageBox.Show("Select the row you want to delete")
        End If
    End Sub

    Private Sub FindAlternativePartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindAlternativePartToolStripMenuItem.Click
        'Replace the select part with an alternative if found

        If PR_grid.CurrentCell.Value.ToString <> String.Empty Then

            row_in = Find_index(PR_grid.CurrentCell.Value.ToString)
            ' myQTY = PR_grid.Rows(row_in).Cells(6).Value

            If row_in > -1 Then
                myQTY = PR_grid.Rows(row_in).Cells(6).Value
                Call Show_alternatives(PR_grid.Rows(row_in).Cells(2).Value, PR_grid.CurrentCell.Value)

            End If
        End If
    End Sub
    Sub Total_rows()

        'Calculate total row number

        Dim total_parts As Double : total_parts = 0

        For i = 0 To PR_grid.Rows.Count - 1
            total_parts = total_parts + PR_grid.Rows(i).Cells(6).Value()
        Next

        Label1.Text = "# Of Parts: " & total_parts


        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(5).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = row.Cells(6).Value * row.Cells(5).Value
            End If
        Next

    End Sub

    Private Sub PR_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles PR_grid.CellValueChanged
        Call Total_rows()
    End Sub

    Function Find_index(name As String) As Integer

        'Find and return row index of ADA part in datagrid

        Find_index = -1

        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If String.Compare(row.Cells.Item("Column10").Value.ToString, name) = 0 Then
                Find_index = row.Index
                Exit For
            End If
        Next

    End Function

    Sub Show_alternatives(ADA_part As String, name_part As String)

        ADA_Alternatives.Visible = True
        ADA_Alternatives.alt_grid.Rows.Clear()

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@ADA_name", ADA_part)
            cmd.Parameters.AddWithValue("@name_part", name_part)
            cmd.CommandText = "select Legacy_ADA_Number, Part_Name, Part_Description, Primary_Vendor from parts_table where LEGACY_ADA_Number = @ADA_name and Part_Name <> @name_part"

            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    ADA_Alternatives.alt_grid.Rows.Add(New DataGridViewRow)
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(0).Value = reader(0)  'ADA Number                        
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(1).Value = reader(1)  'part name
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(2).Value = reader(2)  'Part description
                    ADA_Alternatives.alt_grid.Rows(ADA_Alternatives.alt_grid.Rows.Count - 1).Cells(3).Value = reader(3)  'vendor
                End While
            End If

            reader.Close()

            For Each row As DataGridViewRow In ADA_Alternatives.alt_grid.Rows
                If row.IsNewRow Then Continue For
                row.Cells(4).Value = Form1.Get_Latest_Cost(Login.Connection, row.Cells(1).Value, row.Cells(3).Value)
                row.Height = 52
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub ExportToCSVToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToCSVToolStripMenuItem.Click
        'export Purchase Request to excel file
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
                xlWorkSheet.Range("D:I").ColumnWidth = 20
                xlWorkSheet.Range("A:I").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To PR_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = PR_grid.Columns(i).HeaderText
                    For j As Integer = 0 To PR_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = PR_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Purchase_Request.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Request exported successfully!")
            End If
        End If

    End Sub

    Private Sub Purchase_Request_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' ProgressBar1.Visible = False
        whatipress = 0

        datatable = New DataTable
        datatable.Columns.Add("qty", GetType(String))
        datatable.Columns.Add("vendor", GetType(String))
        datatable.Columns.Add("parts", GetType(String))
        datatable.Columns.Add("desc", GetType(String))


        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1

            If isintable(myQuote.my_alloc_table.Rows(i).Item(0).ToString) = True Then
                Call update_Qty(myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString)
            Else
                PR_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(9).ToString, myQuote.my_alloc_table.Rows(i).Item(2).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
            End If

        Next

        Call Total_rows()

    End Sub


    Function isintable(part_name As String) As Boolean
        isintable = False

        For i = 0 To PR_grid.Rows.Count - 1
            If String.Equals(part_name, PR_grid.Rows(i).Cells(0).Value) = True Then
                isintable = True
                Exit For
            End If
        Next

    End Function


    Sub update_Qty(part_name As String, qty As Double)

        For i = 0 To PR_grid.Rows.Count - 1
            If String.Equals(part_name, PR_grid.Rows(i).Cells(0).Value) = True Then
                PR_grid.Rows(i).Cells(6).Value = If(IsNumeric(PR_grid.Rows(i).Cells(6).Value), PR_grid.Rows(i).Cells(6).Value, 0) + qty
                Exit For
            End If
        Next

    End Sub


    'Private Sub CreateMasterPackingListToolStripMenuItem_Click(sender As Object, e As EventArgs)

    '    '--------------- GENERATE A MASTER PACKING LIST -----------------
    '    If myQuote.Panel_grid.Columns.Count > 1 Then

    '        Dim count_add_rows As Integer : count_add_rows = 8
    '        Dim appPath As String = Application.StartupPath()

    '        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
    '        Dim xlWorkSheet As Excel.Worksheet
    '        xlApp.DisplayAlerts = False

    '        If xlApp Is Nothing Then
    '            MessageBox.Show("Excel is not properly installed!!")
    '        Else
    '            Try
    '                ProgressBar1.Visible = True
    '                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\1XXXXX ADA Packing List r2.0.xlsm")
    '                xlWorkSheet = wb.Sheets("Master Packing List")

    '                Dim totals_c As Double : totals_c = 0

    '                For i = 0 To myQuote.Field_grid.Rows.Count - 1

    '                    ' datatable.Rows.Clear()

    '                    If (ProgressBar1.Value < 400) Then
    '                        ProgressBar1.Value = ProgressBar1.Value + 10
    '                    End If

    '                    totals_c = 0

    '                    '----- get total qty of all sets
    '                    For j = 1 To myQuote.Field_grid.Columns.Count - 1
    '                        If IsNumeric(myQuote.Field_grid.Rows(i).Cells(j).Value) = True Then
    '                            totals_c = totals_c + myQuote.Field_grid.Rows(i).Cells(j).Value
    '                        End If
    '                    Next

    '                    Dim my_feature_code As String : my_feature_code = "notfound"

    '                    '------- Get Feature code from feature description -------- (Note: I may add this in the feature_parts to make everything faster)
    '                    Dim cmd As New MySqlCommand
    '                    cmd.Parameters.AddWithValue("@feature_desc", myQuote.Field_grid.Rows(i).Cells(0).Value)
    '                    cmd.CommandText = "SELECT Feature_code from quote_table.feature_codes where description = @feature_desc"
    '                    cmd.Connection = Login.Connection
    '                    Dim reader As MySqlDataReader
    '                    reader = cmd.ExecuteReader

    '                    If reader.HasRows Then
    '                        While reader.Read
    '                            my_feature_code = reader(0).ToString
    '                        End While
    '                    End If

    '                    reader.Close()



    '                    'get BOM for each feature code
    '                    Dim cmd2 As New MySqlCommand
    '                    cmd2.Parameters.AddWithValue("@feature_code", my_feature_code)
    '                    cmd2.Parameters.AddWithValue("@solution", myQuote.sol_label.Text)
    '                    cmd2.Parameters.AddWithValue("@type", "Field")
    '                    cmd2.CommandText = "SELECT part_name, qty from quote_table.feature_parts where Feature_code = @feature_code and solution = @solution and type = @type"
    '                    cmd2.Connection = Login.Connection
    '                    Dim reader2 As MySqlDataReader
    '                    reader2 = cmd2.ExecuteReader

    '                    If reader2.HasRows Then
    '                        While reader2.Read

    '                            If totals_c > 0 Then

    '                                If is_in_table(reader2(0).ToString) = True Then  'if the part exist just overwrite the qty
    '                                    Call insert_into_table(reader2(0).ToString, If(IsNumeric(reader2(1)) = True, reader2(1), 0) * totals_c)
    '                                Else
    '                                    datatable.Rows.Add(If(IsNumeric(reader2(1)) = True, reader2(1), 0) * totals_c, "", reader2(0).ToString, "")
    '                                End If

    '                            End If

    '                        End While
    '                    End If

    '                    reader2.Close()

    '                Next

    '                'loop through the datatable and get vendors and description
    '                For z = 0 To datatable.Rows.Count - 1
    '                    Dim cmd5 As New MySqlCommand
    '                    cmd5.Parameters.AddWithValue("@part", datatable.Rows(z).Item(2).ToString)
    '                    cmd5.CommandText = "SELECT Primary_Vendor, Part_Description from parts_table where part_name = @part"
    '                    cmd5.Connection = Login.Connection
    '                    Dim reader5 As MySqlDataReader
    '                    reader5 = cmd5.ExecuteReader

    '                    If reader5.HasRows Then
    '                        While reader5.Read
    '                            datatable.Rows(z).Item(1) = reader5(0).ToString
    '                            datatable.Rows(z).Item(3) = reader5(1).ToString
    '                        End While
    '                    End If

    '                    reader5.Close()

    '                    xlWorkSheet.Cells(count_add_rows, 1) = If(datatable.Rows(z).Item(0) = True, datatable.Rows(z).Item(0), 0) ' qty
    '                    xlWorkSheet.Cells(count_add_rows, 3) = datatable.Rows(z).Item(1) ' supplier
    '                    xlWorkSheet.Cells(count_add_rows, 4) = datatable.Rows(z).Item(2) 'part name
    '                    xlWorkSheet.Cells(count_add_rows, 5) = datatable.Rows(z).Item(3) ' description
    '                    count_add_rows = count_add_rows + 1
    '                Next

    '                xlWorkSheet.Range("D:D").HorizontalAlignment = Excel.Constants.xlCenter
    '                xlWorkSheet.Range("E:E").HorizontalAlignment = Excel.Constants.xlCenter

    '                SaveFileDialog1.Filter = "Excel Files|*.xlsm"

    '                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
    '                    wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)
    '                End If

    '                wb.Close(False)

    '                Marshal.ReleaseComObject(xlApp)
    '                ProgressBar1.Visible = False
    '                ProgressBar1.Value = ProgressBar1.Minimum
    '                MessageBox.Show("Master Packing List generated successfully!")

    '            Catch ex As Exception
    '                MessageBox.Show(ex.ToString)
    '                '    MessageBox.Show("Master Packing List Template not found or corrupted")
    '            End Try

    '        End If
    '    End If
    'End Sub

    Function is_in_table(part As String) As Boolean

        '----- check if its in datatable -----

        is_in_table = False

        For i = 0 To datatable.Rows.Count - 1
            If String.Equals(part, datatable.Rows(i).Item(2)) = True Then
                is_in_table = True
                Exit For
            End If
        Next

    End Function

    Sub insert_into_table(part As String, qty As Integer)

        For i = 0 To datatable.Rows.Count - 1
            If String.Equals(part, datatable.Rows(i).Item(2)) = True Then
                datatable.Rows(i).Item(0) = CType(datatable.Rows(i).Item(0), Double) + qty
                Exit For
            End If
        Next

    End Sub

    'Sub Create_MPL()
    '    '-------- Create MPL Master Packing List ------------
    '    If myQuote.Panel_grid.Columns.Count > 1 Then

    '        Dim count_add_rows As Integer : count_add_rows = 0

    '        Try
    '            Dim totals_c As Double : totals_c = 0

    '            For i = 0 To myQuote.Field_grid.Rows.Count - 1

    '                totals_c = 0

    '                '----- get total qty of all sets
    '                For j = 1 To myQuote.Field_grid.Columns.Count - 1
    '                    If IsNumeric(myQuote.Field_grid.Rows(i).Cells(j).Value) = True Then
    '                        totals_c = totals_c + myQuote.Field_grid.Rows(i).Cells(j).Value
    '                    End If
    '                Next

    '                Dim my_feature_code As String : my_feature_code = "notfound"

    '                '------- Get Feature code from feature description -------- (Note: I may add this in the feature_parts to make everything faster)
    '                Dim cmd As New MySqlCommand
    '                cmd.Parameters.AddWithValue("@feature_desc", myQuote.Field_grid.Rows(i).Cells(0).Value)
    '                cmd.CommandText = "SELECT Feature_code from quote_table.feature_codes where description = @feature_desc"
    '                cmd.Connection = Login.Connection
    '                Dim reader As MySqlDataReader
    '                reader = cmd.ExecuteReader

    '                If reader.HasRows Then
    '                    While reader.Read
    '                        my_feature_code = reader(0).ToString
    '                    End While
    '                End If

    '                reader.Close()



    '                'get BOM for each feature code
    '                Dim cmd2 As New MySqlCommand
    '                cmd2.Parameters.AddWithValue("@feature_code", my_feature_code)
    '                cmd2.Parameters.AddWithValue("@solution", myQuote.sol_label.Text)
    '                cmd2.Parameters.AddWithValue("@type", "Field")
    '                cmd2.CommandText = "SELECT part_name, qty from quote_table.feature_parts where Feature_code = @feature_code and solution = @solution and type = @type"
    '                cmd2.Connection = Login.Connection
    '                Dim reader2 As MySqlDataReader
    '                reader2 = cmd2.ExecuteReader

    '                If reader2.HasRows Then
    '                    While reader2.Read

    '                        If totals_c > 0 Then

    '                            If is_in_table(reader2(0).ToString) = True Then  'if the part exist just overwrite the qty
    '                                Call insert_into_table(reader2(0).ToString, If(IsNumeric(reader2(1)) = True, reader2(1), 0) * totals_c)
    '                            Else
    '                                datatable.Rows.Add(If(IsNumeric(reader2(1)) = True, reader2(1), 0) * totals_c, "", reader2(0).ToString, "")
    '                            End If

    '                        End If

    '                    End While
    '                End If

    '                reader2.Close()

    '            Next

    '            'loop through the datatable and get vendors and description
    '            For z = 0 To datatable.Rows.Count - 1
    '                Dim cmd5 As New MySqlCommand
    '                cmd5.Parameters.AddWithValue("@part", datatable.Rows(z).Item(2).ToString)
    '                cmd5.CommandText = "SELECT Primary_Vendor, Part_Description from parts_table where part_name = @part"
    '                cmd5.Connection = Login.Connection
    '                Dim reader5 As MySqlDataReader
    '                reader5 = cmd5.ExecuteReader

    '                If reader5.HasRows Then
    '                    While reader5.Read
    '                        datatable.Rows(z).Item(1) = reader5(0).ToString
    '                        datatable.Rows(z).Item(3) = reader5(1).ToString
    '                    End While
    '                End If

    '                reader5.Close()

    '                '  MPL_grid.Rows(count_add_rows).Cells(3).Value = If(datatable.Rows(z).Item(0) = True, datatable.Rows(z).Item(0), 0) ' qty
    '                '  MPL_grid.Rows(count_add_rows).Cells(2).Value = datatable.Rows(z).Item(1) ' supplier
    '                ' MPL_grid.Rows(count_add_rows).Cells(0).Value = datatable.Rows(z).Item(2) 'part name
    '                '  MPL_grid.Rows(count_add_rows).Cells(1).Value = datatable.Rows(z).Item(3) ' description
    '                ' PR_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(9).ToString, myQuote.my_alloc_table.Rows(i).Item(2).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
    '                MPL_grid.Rows.Add(New String() {datatable.Rows(z).Item(2), datatable.Rows(z).Item(3), datatable.Rows(z).Item(1), If(datatable.Rows(z).Item(0) = True, datatable.Rows(z).Item(0), 0)})
    '                count_add_rows = count_add_rows + 1
    '            Next



    '        Catch ex As Exception
    '            MessageBox.Show(ex.ToString)
    '        End Try

    '    End If


    'End Sub

    Private Sub SavePurcharseRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SavePurcharseRequestToolStripMenuItem.Click
        whatipress = 1
        Save_MR.ShowDialog()

    End Sub

    'Private Sub SaveMasterPackingListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveMasterPackingListToolStripMenuItem.Click

    '    whatipress = 2
    '    Save_MR.ShowDialog()

    'End Sub
End Class