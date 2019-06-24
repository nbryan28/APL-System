Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class Allocation_parts
    Private Sub Allocation_parts_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0

        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1
            alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
        Next

        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next

        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")

        Call Check_zero()

    End Sub


    Sub Check_zero()
        For i = 0 To alloc_grid.Rows.Count - 1
            If String.Equals("$0", alloc_grid.Rows(i).Cells(4).Value.ToString) = True Then
                alloc_grid.Rows(i).DefaultCellStyle.BackColor = Color.Red
            End If
        Next
    End Sub

    Private Sub alloc_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles alloc_grid.CellValueChanged

        'updates cells
        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0

        For Each row As DataGridViewRow In alloc_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(3).Value) = True And IsNumeric(row.Cells(4).Value)) Then
                row.Cells(5).Value = row.Cells(3).Value * row.Cells(4).Value
            End If
        Next

        'cal totals
        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next

        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
    End Sub

    Private Sub alloc_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles alloc_grid.RowsRemoved


        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0


        'cal totals
        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next

        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
    End Sub

    Private Sub alloc_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles alloc_grid.RowsAdded
        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0


        'cal totals
        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next

        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim found_po As Boolean : found_po = False
        'Dim rowindex As Integer

        ''search
        'For Each row As DataGridViewRow In alloc_grid.Rows
        '    If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
        '        If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text, True) = 0 Then
        '            rowindex = row.Index
        '            alloc_grid.CurrentCell = alloc_grid.Rows(rowindex).Cells(0)
        '            found_po = True
        '            Exit For
        '        End If
        '    End If
        'Next

        'If found_po = False Then
        '    MsgBox("Part not found!")
        'End If
        alloc_grid.Rows.Clear()

        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1
            If myQuote.my_alloc_table.Rows(i).Item(0).ToString.IndexOf(TextBox1.Text, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
            End If
        Next

    End Sub

    Private Sub type_b_SelectedValueChanged(sender As Object, e As EventArgs) Handles type_b.SelectedValueChanged
        'filter type

        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0
        alloc_grid.Rows.Clear()



        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1
            If String.Equals(type_b.Text, myQuote.my_alloc_table.Rows(i).Item(8).ToString) = True Then
                alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
            End If
        Next



        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next

        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'show all
        Dim total_parts As Double : total_parts = 0
        Dim total_cost As Double : total_cost = 0
        alloc_grid.Rows.Clear()

        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1
            alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})

        Next

        For i = 0 To alloc_grid.Rows.Count - 1
            If IsNumeric(alloc_grid.Rows(i).Cells(3).Value()) = True Then
                total_parts = total_parts + alloc_grid.Rows(i).Cells(3).Value()
            End If

            If IsNumeric(alloc_grid.Rows(i).Cells(5).Value()) = True Then
                total_cost = total_cost + alloc_grid.Rows(i).Cells(5).Value()
            End If
        Next



        pt_label.Text = "Parts Needed: " & total_parts
        mt_label.Text = "Materials Cost: $" & Decimal.Round(total_cost, 2, MidpointRounding.AwayFromZero).ToString("N")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
                For i As Integer = 0 To alloc_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = alloc_grid.Columns(i).HeaderText
                    For j As Integer = 0 To alloc_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = alloc_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\allocations_of_parts.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Data exported successfully!")
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'search description
        alloc_grid.Rows.Clear()

        For i = 0 To myQuote.my_alloc_table.Rows.Count - 1
            If myQuote.my_alloc_table.Rows(i).Item(1).ToString.IndexOf(desc.Text, StringComparison.CurrentCultureIgnoreCase) > -1 Then
                alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
            End If
        Next
    End Sub



    Function isintable(part_name As String) As Boolean
        isintable = False

        For i = 0 To alloc_grid.Rows.Count - 1
            If String.Equals(part_name, alloc_grid.Rows(i).Cells(0).Value) = True Then
                isintable = True
                Exit For
            End If
        Next

    End Function


    Sub update_Qty(part_name As String, qty As Double)

        For i = 0 To alloc_grid.Rows.Count - 1
            If String.Equals(part_name, alloc_grid.Rows(i).Cells(0).Value) = True Then
                alloc_grid.Rows(i).Cells(3).Value = If(IsNumeric(alloc_grid.Rows(i).Cells(3).Value), alloc_grid.Rows(i).Cells(3).Value, 0) + qty
                Exit For
            End If
        Next

    End Sub

    'Sub merge_table()

    '    For i = 0 To myQuote.my_alloc_table.Rows.Count - 1

    '        If isintable(myQuote.my_alloc_table.Rows(i).Item(0).ToString) = True Then
    '            Call update_Qty(myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString)
    '        Else
    '            alloc_grid.Rows.Add(New String() {myQuote.my_alloc_table.Rows(i).Item(0).ToString, myQuote.my_alloc_table.Rows(i).Item(1).ToString, myQuote.my_alloc_table.Rows(i).Item(3).ToString, myQuote.my_alloc_table.Rows(i).Item(5).ToString, myQuote.my_alloc_table.Rows(i).Item(4).ToString, myQuote.my_alloc_table.Rows(i).Item(6).ToString, myQuote.my_alloc_table.Rows(i).Item(7).ToString})
    '        End If

    '    Next
    'End Sub
End Class