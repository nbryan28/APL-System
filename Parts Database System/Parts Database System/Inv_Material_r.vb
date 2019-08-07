Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class Inv_Material_r

    Public open_job As String
    Public released_by As String
    Public go_ahead As Boolean


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub OpenProjecrToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenProjecrToolStripMenuItem.Click
        My_Material_r.mode = "inv_rel"
        Open_file.ShowDialog()
    End Sub

    Private Sub fullfill_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles fullfill_grid.CellValueChanged

        If go_ahead = True Then

            For Each row As DataGridViewRow In fullfill_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(7).Value) = True And IsNumeric(row.Cells(9).Value)) Then

                    If CType(row.Cells(8).Value, Double) < CType(row.Cells(9).Value, Double) Then
                        ' MessageBox.Show("Not enough qty in current inventory")
                        row.Cells(9).Value = 0

                    End If


                    'If CType(row.Cells(9).Value, Double) < 0 Then
                    '    MessageBox.Show("Value cannot be negative")
                    '    row.Cells(9).Value = 0
                    'End If

                    If (IsNumeric(row.Cells(7).Value) = True And IsNumeric(row.Cells(10).Value)) Then
                        row.Cells(11).Value = CType(row.Cells(7).Value, Double) - CType(row.Cells(10).Value, Double)
                    Else
                        row.Cells(11).Value = If(IsNumeric(row.Cells(7).Value) = False, 0, row.Cells(7).Value)
                    End If


                    If CType(row.Cells(11).Value, Double) <> 0 Then
                        row.Cells(11).Style.BackColor = Color.Firebrick
                    Else
                        row.Cells(11).Style.BackColor = Color.Gray
                    End If

                End If
            Next
        End If
    End Sub

    Private Sub NotifyProgressToProcurementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NotifyProgressToProcurementToolStripMenuItem.Click

        If String.Equals(Me.Text, "My Material Requests") = False Then

            Try

                '---------- calculate percentage ---------
                Dim complete_mr As Double : complete_mr = 0
                Dim total_qty As Double : total_qty = 0
                Dim fullf As Double : fullf = 0


                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
                cmd3.Connection = Login.Connection
                Dim reader3 As MySqlDataReader
                reader3 = cmd3.ExecuteReader

                If reader3.HasRows Then
                    While reader3.Read
                        If IsDBNull(reader3(0)) Then
                            fullf = 0
                        Else
                            fullf = CType(reader3(0), Double)
                        End If
                    End While
                End If

                reader3.Close()

                Dim cmdx As New MySqlCommand
                cmdx.Parameters.AddWithValue("@mr_name", Me.Text)
                cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
                cmdx.Connection = Login.Connection
                Dim readerx As MySqlDataReader
                readerx = cmdx.ExecuteReader

                If readerx.HasRows Then
                    While readerx.Read
                        If IsDBNull(readerx(0)) Then
                            total_qty = 0
                        Else
                            total_qty = CType(readerx(0), Double)
                        End If
                    End While
                End If

                readerx.Close()

                If total_qty > 0 Then
                    complete_mr = Math.Round((fullf / total_qty) * 100)
                End If
                '---------------------------------------

                '---------- send notification ------------

                If enable_mess = True Then

                    Dim email_m As String
                    email_m = "Material Request Allocation for Project " & open_job & "  has been updated" & vbCrLf & vbCrLf _
                                & "Sent by: " & current_user & vbCrLf _
                                & "Sent Date: " & Now & vbCrLf _
                                & "Project: " & open_job & vbCrLf _
                                & "MR Filename: " & Me.Text & vbCrLf _
                                & "Material Request Status: " & complete_mr & "% fulfilled" & vbCrLf


                    Dim message_title As String : message_title = "Material Request Allocation for Project " & open_job & " updated"

                    Call Sent_mail.Sent_multiple_emails("Procurement", message_title, email_m)
                    Call Sent_mail.Sent_multiple_emails("Procurement Management", message_title, email_m)
                    Call Sent_mail.Sent_multiple_emails("Inventory", message_title, email_m)
                    'Call Sent_mail.Sent_multiple_emails("Manufacturing", message_title, email_m)
                    ' Call Sent_mail.Sent_multiple_emails("General Management", message_title, email_m)
                    ' Call Sent_mail.Sent_multiple_emails("Engineer Management", message_title, email_m)
                End If

                MessageBox.Show("Notification sent!")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Inv_Material_r_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        EnableDoubleBuffered(fullfill_grid)
        go_ahead = True

    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '--------------- Start updating the Mr table, the current inventory qty and generate a material report -------------------


        If String.Equals(Me.Text, "My Material Requests") = False Then

            wait_la.Visible = True
            Application.DoEvents()

            '------- check if its an assembly MR --------
            Dim asm_bom As Boolean : asm_bom = False

            For i = 0 To fullfill_grid.Rows.Count - 1
                If String.IsNullOrEmpty(fullfill_grid.Rows(i).Cells(14).Value) = False Then
                    If String.Equals(fullfill_grid.Rows(i).Cells(14).Value, "") = False Then
                        asm_bom = True
                        Exit For
                    End If
                End If
            Next

            If asm_bom = False Then

                '--------------- find latest version --------------
                Dim latest_v As String : latest_v = Proc_Material_R.get_last_revision(Me.Text)
                '---------------------------------------------------
                Call Update_mr_fullfill(latest_v)   'update latest revision mr
                Call Update_mr_fullfill(Me.Text)    'update current mr which can be the same as above
                Call Update_current_inv()  'update inventory current qtys
                Call Inventory_manage.General_inv_cal()  'fill material order form

                go_ahead = False

                '--- reload entire table except current inv----
                fullfill_grid.Rows.Clear()


                go_ahead = True
                job_label.Text = "Open Project: "

                wait_la.Visible = False
                MessageBox.Show("Material Request updated successfully!")

                Call open_recent(Me.Text, open_job)

            Else

                Call Update_mr_fullfill_asm(Me.Text)    'update current mr which can be the same as above
                ' Call Update_current_inv_asm()  'update inventory current qtys
                Call Inventory_manage.General_inv_cal()  'fill material order form

                go_ahead = False

                '--- reload entire table except current inv----
                fullfill_grid.Rows.Clear()

                go_ahead = True
                job_label.Text = "Open Project: "

                wait_la.Visible = False
                MessageBox.Show("Assembly BOM updated successfully!")
            End If

        End If


    End Sub

    Sub Update_mr_fullfill(mr_name As String)

        Try

            For i = 0 To fullfill_grid.Rows.Count - 1

                If fullfill_grid.Rows(i).Cells(9).Value <> 0 Then

                    Dim qty_f As Double = 0


                    If IsNumeric(fullfill_grid.Rows(i).Cells(10).Value) = True Then
                        qty_f = CType(fullfill_grid.Rows(i).Cells(10).Value, Double)
                    End If

                    If IsNumeric(fullfill_grid.Rows(i).Cells(9).Value) = True Then
                        qty_f = qty_f + CType(fullfill_grid.Rows(i).Cells(9).Value, Double)
                    End If



                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                    cmd5.Parameters.AddWithValue("@mr_name", mr_name)

                    cmd5.CommandText = "UPDATE Material_Request.mr_data SET qty_fullfilled = @qty_fullfilled where mr_name = @mr_name and Part_No = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Update_current_inv()

        Try
            For i = 0 To fullfill_grid.Rows.Count - 1

                If fullfill_grid.Rows(i).Cells(9).Value <> 0 Then

                    Dim current_q As Double : current_q = 0
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        While reader4.Read
                            current_q = reader4(0).ToString
                        End While
                    End If

                    reader4.Close()

                    current_q = current_q - If(IsNumeric(fullfill_grid.Rows(i).Cells(9).Value), fullfill_grid.Rows(i).Cells(9).Value, 0)
                    current_q = If(current_q < 0, 0, current_q)


                    '---- update the value
                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@current_qty", current_q)

                    cmd5.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    'Private Sub AllocateProjectSpecificPartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllocateProjectSpecificPartToolStripMenuItem.Click
    '    If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

    '        Dim component As String : component = ""
    '        Dim index As Integer : index = 0
    '        component = fullfill_grid.CurrentCell.Value.ToString
    '        index = fullfill_grid.CurrentCell.RowIndex

    '        If fullfill_grid.Rows(index).Cells(3).Value Is Nothing = False Then

    '            If String.Equals(fullfill_grid.Rows(index).Cells(12).Value, "No") = True Then

    '                Dim qty_change As Integer
    '                Dim qty_test As Integer

    '                Try

    '                    qty_change = InputBox("Please enter the number of " & component & " to allocate in the MR")

    '                    If Integer.TryParse(qty_change, qty_test) Then
    '                        fullfill_grid.Rows(index).Cells(10).Value = qty_change
    '                    Else
    '                        MsgBox("Please input an integer.")

    '                    End If

    '                Catch
    '                    MsgBox("Please input a whole number.")
    '                End Try

    '            Else
    '                MessageBox.Show("Not a Special Order!")
    '            End If

    '        Else
    '            MessageBox.Show("Not a Special Order!")
    '        End If
    '    End If
    'End Sub

    'Private Sub OvewriteQtyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OvewriteQtyToolStripMenuItem.Click
    '    '-- overwrite qty for now it will handle units
    '    If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

    '        Dim component As String : component = ""
    '        Dim index As Integer : index = 0
    '        component = fullfill_grid.CurrentCell.Value.ToString
    '        index = fullfill_grid.CurrentCell.RowIndex

    '        If fullfill_grid.Rows(index).Cells(3).Value Is Nothing = False Then

    '            If String.Equals(fullfill_grid.Rows(index).Cells(12).Value, "No") = True Then

    '                Dim qty_change As Integer
    '                Dim qty_test As Integer

    '                Try

    '                    qty_change = InputBox("Please enter the number of " & component & " to allocate in the MR")

    '                    If Integer.TryParse(qty_change, qty_test) Then
    '                        fullfill_grid.Rows(index).Cells(10).Value = qty_change
    '                    Else
    '                        MsgBox("Please input an integer.")

    '                    End If

    '                Catch
    '                    MsgBox("Please input a whole number.")
    '                End Try

    '            Else
    '                MessageBox.Show("Not a Special Order!")
    '            End If

    '        Else
    '            MessageBox.Show("Not a Special Order!")
    '        End If
    '    End If
    'End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
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
                xlWorkSheet.Range("D:M").ColumnWidth = 20
                xlWorkSheet.Range("A:M").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To fullfill_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = fullfill_grid.Columns(i).HeaderText
                    For j As Integer = 0 To fullfill_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = fullfill_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Report_inventory.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Inventory Material Request exported successfully!")
            End If
        End If
    End Sub

    Sub Update_mr_fullfill_asm(mr_name As String)

        Try
            For i = 0 To fullfill_grid.Rows.Count - 1

                Dim update_v As Boolean : update_v = False

                '------- get inventory ---------
                Dim current_q As Double : current_q = 0
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        current_q = reader4(0).ToString
                    End While
                End If

                reader4.Close()
                '----------------------------------

                Dim qty_f As Double = 0

                If IsNumeric(fullfill_grid.Rows(i).Cells(10).Value) = True Then
                    qty_f = CType(fullfill_grid.Rows(i).Cells(10).Value, Double)
                End If

                If IsNumeric(fullfill_grid.Rows(i).Cells(9).Value) = True Then

                    If fullfill_grid.Rows(i).Cells(9).Value <= current_q Then
                        qty_f = qty_f + CType(fullfill_grid.Rows(i).Cells(9).Value, Double)
                        update_v = True
                    End If

                End If

                If update_v = True Then

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                    cmd5.Parameters.AddWithValue("@mr_name", mr_name)
                    cmd5.Parameters.AddWithValue("@asm", fullfill_grid.Rows(i).Cells(14).Value)

                    cmd5.CommandText = "UPDATE Material_Request.mr_data SET qty_fullfilled = @qty_fullfilled where mr_name = @mr_name and Part_No = @part_name and Assembly_name = @asm"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                    Dim qty_ad As Double : qty_ad = 0
                    qty_ad = If(IsNumeric(fullfill_grid.Rows(i).Cells(9).Value), fullfill_grid.Rows(i).Cells(9).Value, 0)

                    If current_q >= qty_ad Then

                        current_q = current_q - qty_ad
                        current_q = If(current_q < 0, 0, current_q)

                        '---- update the value
                        Dim cmd51 As New MySqlCommand
                        cmd51.Parameters.Clear()
                        cmd51.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                        cmd51.Parameters.AddWithValue("@current_qty", current_q)

                        cmd51.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                        cmd51.Connection = Login.Connection
                        cmd51.ExecuteNonQuery()

                    End If

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Update_current_inv_asm()

        Try
            For i = 0 To fullfill_grid.Rows.Count - 1

                Dim current_q As Double : current_q = 0
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        current_q = reader4(0).ToString
                    End While
                End If

                reader4.Close()

                Dim qty_ad As Double : qty_ad = 0
                qty_ad = If(IsNumeric(fullfill_grid.Rows(i).Cells(9).Value), fullfill_grid.Rows(i).Cells(9).Value, 0)

                If current_q >= qty_ad Then

                    current_q = current_q - qty_ad 'If(IsNumeric(fullfill_grid.Rows(i).Cells(9).Value), fullfill_grid.Rows(i).Cells(9).Value, 0)
                    current_q = If(current_q < 0, 0, current_q)

                    '---- update the value
                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@current_qty", current_q)

                    cmd5.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Dim found_po As Boolean : found_po = False
    '    Dim rowindex As Integer

    '    For Each row As DataGridViewRow In fullfill_grid.Rows
    '        If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
    '            If TextBox1.Text.IndexOf(row.Cells.Item("Column10").Value, StringComparison.CurrentCultureIgnoreCase) > -1 Then
    '                rowindex = row.Index
    '                fullfill_grid.CurrentCell = fullfill_grid.Rows(rowindex).Cells(0)
    '                found_po = True
    '                Exit For
    '            End If

    '        End If
    '    Next

    '    If found_po = False Then
    '        MsgBox("Part not found in Inventory Request !")
    '    End If
    'End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If fullfill_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(fullfill_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub DisableSecurityToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) Handles DisableSecurityToolStripMenuItem.CheckedChanged
        If DisableSecurityToolStripMenuItem.Checked = True Then
            go_ahead = False
        Else
            go_ahead = True
        End If
    End Sub

    Private Sub DisableSecurityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisableSecurityToolStripMenuItem.Click
        If DisableSecurityToolStripMenuItem.Checked = True Then
            DisableSecurityToolStripMenuItem.Checked = False
        Else
            DisableSecurityToolStripMenuItem.Checked = True
        End If
    End Sub

    Sub open_recent(name As String, job As String)
        'this sub open the recently close mr

        '---- datatable store MR without assemblies --------
        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("ADA_Number", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("mfg_type", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("qty_needed", GetType(String))
        dimen_table.Columns.Add("Assembly_name", GetType(String))
        dimen_table.Columns.Add("Part_status", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", name)
            cmd4.CommandText = "SELECT Part_No, Description, Manufacturer,  ADA_Number, Vendor, Price, mfg_type, Qty, qty_fullfilled, qty_needed , Assembly_name, Part_status, Notes from Material_Request.mr_data where mr_name = @mr_name and (full_panel is null or full_panel <> 'Yes')"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString, reader4(12).ToString)
                End While

            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                fullfill_grid.Rows.Add(New String() {})
                fullfill_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                fullfill_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                fullfill_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                fullfill_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                fullfill_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                fullfill_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                fullfill_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString
                fullfill_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString
                fullfill_grid.Rows(i).Cells(8).Value = 0
                fullfill_grid.Rows(i).Cells(9).Value = 0
                fullfill_grid.Rows(i).Cells(10).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)
                fullfill_grid.Rows(i).Cells(11).Value = If(IsNumeric(dimen_table.Rows(i).Item(9).ToString) = True, dimen_table.Rows(i).Item(9).ToString, 0)
                fullfill_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(12).ToString
                fullfill_grid.Rows(i).Cells(14).Value = dimen_table.Rows(i).Item(10).ToString
            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To fullfill_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , min_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        fullfill_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                        If reader5(1) > 0 Then
                            fullfill_grid.Rows(i).Cells(12).Value = "Yes"
                        Else
                            fullfill_grid.Rows(i).Cells(12).Value = "no"
                        End If
                    End While
                Else

                    fullfill_grid.Rows(i).Cells(12).Value = "no"
                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

            open_job = job
            job_label.Text = job


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        '-- update inv request according to search
        fullfill_grid.Rows.Clear()

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("ADA_Number", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("mfg_type", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("qty_needed", GetType(String))
        dimen_table.Columns.Add("Assembly_name", GetType(String))
        dimen_table.Columns.Add("Part_status", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
            cmd4.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")
            cmd4.CommandText = "SELECT Part_No, Description, Manufacturer,  ADA_Number, Vendor, Price, mfg_type, Qty, qty_fullfilled, qty_needed , Assembly_name, Part_status, Notes from Material_Request.mr_data where mr_name = @mr_name and Part_No LIKE @search"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString, reader4(12).ToString)
                End While

            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                fullfill_grid.Rows.Add(New String() {})
                fullfill_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                fullfill_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                fullfill_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                fullfill_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                fullfill_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                fullfill_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                fullfill_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString
                fullfill_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString
                fullfill_grid.Rows(i).Cells(8).Value = 0
                fullfill_grid.Rows(i).Cells(9).Value = 0
                fullfill_grid.Rows(i).Cells(10).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)
                fullfill_grid.Rows(i).Cells(11).Value = If(IsNumeric(dimen_table.Rows(i).Item(9).ToString) = True, dimen_table.Rows(i).Item(9).ToString, 0)
                fullfill_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(12).ToString
                fullfill_grid.Rows(i).Cells(14).Value = dimen_table.Rows(i).Item(10).ToString
            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To fullfill_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , min_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        fullfill_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                        If reader5(1) > 0 Then
                            fullfill_grid.Rows(i).Cells(12).Value = "Yes"
                        Else
                            fullfill_grid.Rows(i).Cells(12).Value = "no"
                        End If
                    End While

                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class