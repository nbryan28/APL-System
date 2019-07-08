Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Net.Mail

Public Class Build_mfg

    Public Smtp_Server As New SmtpClient


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    'Private Sub ExportMPLToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMPLToExcelToolStripMenuItem.Click
    '    'export  MPL to excel file

    '    Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
    '    xlApp.DisplayAlerts = False

    '    If xlApp Is Nothing Then
    '        MessageBox.Show("Excel is not properly installed!!")
    '    Else
    '        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

    '            Dim xlWorkBook As Excel.Workbook
    '            Dim xlWorkSheet As Excel.Worksheet
    '            Dim misValue As Object = System.Reflection.Missing.Value
    '            xlWorkBook = xlApp.Workbooks.Add(misValue)
    '            xlWorkSheet = xlWorkBook.Sheets("sheet1")
    '            xlWorkSheet.Range("A:B").ColumnWidth = 40
    '            xlWorkSheet.Range("C:C").ColumnWidth = 30
    '            xlWorkSheet.Range("D:E").ColumnWidth = 20
    '            xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter


    '            'copy data to excel
    '            For i As Integer = 0 To PR_grid.ColumnCount - 1

    '                xlWorkSheet.Cells(1, i + 1) = PR_grid.Columns(i).HeaderText
    '                For j As Integer = 0 To PR_grid.RowCount - 1
    '                    xlWorkSheet.Cells(j + 2, i + 1) = PR_grid.Rows(j).Cells(i).Value
    '                Next j

    '            Next i

    '            Try
    '                xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Build_Request.xlsx")
    '            Catch ex As Exception

    '            End Try
    '            xlWorkBook.Close(True)
    '            xlApp.Quit()


    '            Marshal.ReleaseComObject(xlApp)

    '            MessageBox.Show("Build Request exported successfully!")
    '        End If
    '    End If
    'End Sub

    'Private Sub OpenMPLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenMPLToolStripMenuItem.Click

    '    My_Material_r.mode = "br_mfg"
    '    Open_file.ShowDialog()

    'End Sub

    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Dim found_po As Boolean : found_po = False
    '    Dim rowindex As Integer

    '    For Each row As DataGridViewRow In PR_grid.Rows
    '        If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
    '            If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text) = 0 Then
    '                rowindex = row.Index
    '                PR_grid.CurrentCell = PR_grid.Rows(rowindex).Cells(0)
    '                found_po = True
    '                Exit For
    '            End If

    '        End If
    '    Next

    '    If found_po = False Then
    '        MsgBox("Part not found!")
    '    End If
    'End Sub

    Private Sub PR_grid_DoubleClick(sender As Object, e As EventArgs) Handles PR_grid.DoubleClick
        ''---- double click to open associated pdf drawing

        'If String.Equals(Me.Text, "Build Request") = False Then

        '    Dim index_row As Integer = PR_grid.CurrentRow.Index
        '    Dim panel_name As String = PR_grid.Rows(index_row).Cells(0).Value
        '    Dim count_m As Integer : count_m = 0
        '    Dim file_n As String : file_n = ""

        '    '--- get draw_Description from full panel table and the pdf file names from management.draw_pdf
        '    Dim draw_desc As String : draw_desc = ""


        '    Try
        '        Dim cmd4 As New MySqlCommand
        '        cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
        '        cmd4.Parameters.AddWithValue("@panel_name", panel_name)
        '        cmd4.CommandText = "SELECT draw_description from Material_Request.full_panels where mr_name = @mr_name and panel_name = @panel_name"
        '        cmd4.Connection = Login.Connection
        '        Dim reader4 As MySqlDataReader
        '        reader4 = cmd4.ExecuteReader

        '        If reader4.HasRows Then
        '            While reader4.Read
        '                draw_desc = reader4(0).ToString
        '            End While
        '        End If

        '        reader4.Close()

        '        '---------- see if there is more than one drawing associated to ----------------
        '        Dim cmd41 As New MySqlCommand
        '        cmd41.Parameters.AddWithValue("@draw", draw_desc)
        '        cmd41.CommandText = "SELECT count(*) from management.draw_pdf where draw_description = @draw"
        '        cmd41.Connection = Login.Connection
        '        Dim reader41 As MySqlDataReader
        '        reader41 = cmd41.ExecuteReader

        '        If reader41.HasRows Then
        '            While reader41.Read
        '                count_m = reader41(0).ToString
        '            End While
        '        End If

        '        reader41.Close()
        '        '------------------------------------------------------------------------------

        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)
        '    End Try

        '    If count_m > 1 Then
        '        Drawing_Selector.ShowDialog()
        '    Else

        '        Try
        '            '----- get filename -----
        '            Dim cmd41 As New MySqlCommand
        '            cmd41.Parameters.AddWithValue("@draw_description", draw_desc)
        '            cmd41.CommandText = "SELECT file_name from management.draw_pdf where draw_description = @draw_description"
        '            cmd41.Connection = Login.Connection
        '            Dim reader41 As MySqlDataReader
        '            reader41 = cmd41.ExecuteReader

        '            If reader41.HasRows Then
        '                While reader41.Read
        '                    file_n = reader41(0).ToString
        '                End While
        '            End If

        '            reader41.Close()

        '            '---------------------------

        '            If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\" & file_n) = True Then
        '                System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\" & file_n)

        '            Else
        '                MessageBox.Show("No PDF Drawing Found")
        '            End If

        '        Catch ex As Exception
        '            MessageBox.Show(ex.ToString)
        '        End Try

        '    End If

        'End If

        If PR_grid.Rows.Count > 0 Then

            Dim index_k = PR_grid.CurrentCell.RowIndex
            Dim panel As String = ""


            If PR_grid.Rows.Count > 0 Then
                panel = If(String.IsNullOrEmpty(PR_grid.Rows(index_k).Cells(0).Value) = True, "", PR_grid.Rows(index_k).Cells(0).Value)
            End If

            If String.Equals(panel, "") = False Then


                '---- fil panel data -----
                Dim mr_name As String : mr_name = ""
                Dim temp_panel_Qty As Double : temp_panel_Qty = 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@job", job_label.Text)
                cmd5.Parameters.AddWithValue("@Panel_name", panel)
                cmd5.CommandText = "SELECT mr_name from Material_Request.mr where Panel_name = @Panel_name and job = @job"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        mr_name = reader5(0).ToString
                    End While
                End If
                reader5.Close()

                '-get latest rev
                mr_name = Procurement_Overview.get_last_revision(mr_name)

                '--- get qty panel ---
                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", mr_name)
                cmd419.CommandText = "SELECT Panel_qty from Material_Request.mr where mr_name = @mr_name"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read

                        n_panels.Text = "(" & If(reader419(0) Is DBNull.Value, 1, reader419(0)) & ")"
                        temp_panel_Qty = If(reader419(0) Is DBNull.Value, 1, reader419(0))

                    End While
                End If

                reader419.Close()
                '----------------------------------


                If String.Equals(mr_name, "not found") = False Then
                    Call Open_panel(mr_name, job_label.Text, temp_panel_Qty)
                End If

                '-------------------------

                Call color_needed(total_grid)

                panel_n.Text = panel

                TabControl1.TabPages.Insert(1, TabPage3)
                TabControl1.TabPages.Remove(TabPage1)
                TabControl1.SelectedTab = TabPage3

            End If

        End If

    End Sub

    Sub Open_panel(name As String, job As String, qty_p As Double)

        'this sub open  the data in the inventory request datagridview
        total_grid.Rows.Clear()

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("qty_needed", GetType(String))


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", name)
            cmd4.CommandText = "SELECT Part_No, Description, Qty, Manufacturer, Vendor, Price, qty_fullfilled, qty_needed from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString * qty_p, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString)
                End While
            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                total_grid.Rows.Add(New String() {})
                total_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                total_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                total_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                total_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                total_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                total_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                total_grid.Rows(i).Cells(6).Value = If(IsNumeric(dimen_table.Rows(i).Item(6).ToString) = True, dimen_table.Rows(i).Item(6).ToString, 0)
                total_grid.Rows(i).Cells(7).Value = If(IsNumeric(dimen_table.Rows(i).Item(7).ToString) = True, dimen_table.Rows(i).Item(7).ToString, 0)
            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To total_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        total_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                        total_grid.Rows(i).Cells(9).Value = reader5(1).ToString
                        total_grid.Rows(i).Cells(10).Value = reader5(2).ToString
                        total_grid.Rows(i).Cells(11).Value = reader5(3).ToString
                    End While
                End If

                reader5.Close()

                '---- get project specific qty-----

                Dim cmd51 As New MySqlCommand
                cmd51.Parameters.Clear()
                cmd51.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(0).Value)
                cmd51.Parameters.AddWithValue("@job", job)
                cmd51.CommandText = "SELECT qty_needed, PO, es_date_of_arrival from Tracking_Reports.my_tracking_reports where Part_No = @Part_No and job = @job"
                cmd51.Connection = Login.Connection
                Dim reader51 As MySqlDataReader
                reader51 = cmd51.ExecuteReader

                If reader51.HasRows Then
                    While reader51.Read

                        If String.IsNullOrEmpty(reader51(1).ToString) = False Or String.IsNullOrEmpty(reader51(2).ToString) = False Then

                            total_grid.Rows(i).Cells(9).Value = reader51(0).ToString

                        End If

                    End While
                End If

                reader51.Close()



            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub color_needed(mygrid As DataGridView)

        For Each row As DataGridViewRow In mygrid.Rows
            If row.IsNewRow Then Continue For

            If (IsNumeric(row.Cells(2).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = CType(row.Cells(2).Value, Double) - CType(row.Cells(6).Value, Double)
            Else
                row.Cells(7).Value = 0
            End If

            If CType(row.Cells(7).Value, Double) <> 0 Then
                row.Cells(7).Style.BackColor = Color.Firebrick
            Else
                row.Cells(7).Style.BackColor = Color.Gray
            End If
        Next
    End Sub

    Private Sub Build_mfg_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TabControl1.TabPages.Remove(TabPage3)

        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"


        ComboBox1.Items.Clear()

        Dim check_cmd2 As New MySqlCommand
        check_cmd2.CommandText = "select distinct job from Build_request.build_r order by job"
        check_cmd2.Connection = Login.Connection
        check_cmd2.ExecuteNonQuery()

        Dim reader2 As MySqlDataReader
        reader2 = check_cmd2.ExecuteReader

        If reader2.HasRows Then
            While reader2.Read
                ComboBox1.Items.Add(reader2(0))
            End While
        End If

        reader2.Close()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        Dim shipping_ad As String : shipping_ad = "No Shipping Address"
        job_label.Text = ComboBox1.SelectedItem.ToString
        PR_grid.Rows.Clear()
        mrbox1.Items.Clear()
        mrbox2.Items.Clear()

        Dim n_r As Integer : n_r = 0
        Dim br_name As String : br_name = "Build Request"

        Try

            '--get latest revision --
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
            cmd41.CommandText = "SELECT distinct n_r from Build_request.build_r where job = @job order by n_r desc limit 1"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    n_r = reader41(0).ToString
                End While
            End If

            reader41.Close()
            '----------------------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
            cmd4.Parameters.AddWithValue("@n_r", n_r)
            cmd4.CommandText = "SELECT panel, panel_desc, panel_qty, need_date, br_name, ready_t from Build_request.build_r where job = @job and n_r = @n_r"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    PR_grid.Rows.Add(New String() {})
                    PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                    PR_grid.Rows(i).Cells(4).Value = If(String.Equals(reader4(5).ToString, "Y") = True, True, False)
                    br_name = reader4(4).ToString

                    i = i + 1
                End While

            End If

            reader4.Close()

            Me.Text = br_name

            '------------- get shipping ------------
            Dim cmd46 As New MySqlCommand
            cmd46.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
            cmd46.CommandText = "SELECT Shipping_address from management.projects where Job_number = @job"
            cmd46.Connection = Login.Connection
            Dim reader46 As MySqlDataReader
            reader46 = cmd46.ExecuteReader

            If reader46.HasRows Then
                While reader46.Read
                    shipping_ad = reader46(0).ToString
                End While
            End If

            reader46.Close()
            '--------------------------------------

            shipping_l.Text = "Shipping Address: " & shipping_ad

            '--fill dropboxes --------
            Dim cmdc As New MySqlCommand
            cmdc.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
            cmdc.CommandText = "SELECT distinct br_name from Build_request.build_r where job = @job"
            cmdc.Connection = Login.Connection
            Dim readerc As MySqlDataReader
            readerc = cmdc.ExecuteReader

            If readerc.HasRows Then
                While readerc.Read
                    mrbox1.Items.Add(readerc(0).ToString)
                    mrbox2.Items.Add(readerc(0).ToString)
                End While
            End If

            readerc.Close()


            Call detect_empty()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Build_Request.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Build Request exported successfully!")
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

        If PR_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(PR_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '---- Compare two Build Request
        If Not mrbox1.SelectedItem Is Nothing And Not mrbox2.SelectedItem Is Nothing Then

            compare_grid.Rows.Clear()

            Dim total_mr As New Dictionary(Of String, String)
            '-- store mr1 parts
            Dim dimen_table1 = New DataTable
            dimen_table1.Columns.Add("Part_No", GetType(String))
            dimen_table1.Columns.Add("Description", GetType(String))

            '-- store mr2 parts
            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("Description", GetType(String))

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@br_name", mrbox1.Text)
                cmd4.CommandText = "SELECT panel, panel_desc from Build_request.build_r where br_name = @br_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        dimen_table1.Rows.Add(reader4(0).ToString, reader4(1).ToString)
                    End While
                End If

                reader4.Close()

                '--- mr2
                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@br_name", mrbox2.Text)
                cmd41.CommandText = "SELECT panel, panel_desc from Build_request.build_r where br_name = @br_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        dimen_table2.Rows.Add(reader41(0).ToString, reader41(1).ToString)
                    End While
                End If

                reader41.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            For i = 0 To dimen_table1.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table1.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table1.Rows(i).Item(0), dimen_table1.Rows(i).Item(1))
                End If
            Next

            For i = 0 To dimen_table2.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table2.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table2.Rows(i).Item(0), dimen_table2.Rows(i).Item(1))
                End If
            Next

            '--- add to comp_Grid
            For Each kvp As KeyValuePair(Of String, String) In total_mr.ToArray
                compare_grid.Rows.Add(New String() {kvp.Key, total_mr(kvp.Key)})
            Next

            For i = 0 To compare_grid.Rows.Count - 1
                Try
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.Clear()
                    cmd41.Parameters.AddWithValue("@br_name", mrbox1.Text)
                    cmd41.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd41.CommandText = "SELECT panel_qty from Build_request.build_r where br_name = @br_name and panel = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            compare_grid.Rows(i).Cells(2).Value = reader41(0).ToString
                        End While
                    End If

                    reader41.Close()

                    Dim cmd42 As New MySqlCommand
                    cmd42.Parameters.Clear()
                    cmd42.Parameters.AddWithValue("@br_name", mrbox2.Text)
                    cmd42.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd42.CommandText = "SELECT panel_qty from Build_request.build_r where br_name = @br_name and panel = @part"
                    cmd42.Connection = Login.Connection
                    Dim reader42 As MySqlDataReader
                    reader42 = cmd42.ExecuteReader

                    If reader42.HasRows Then
                        While reader42.Read
                            compare_grid.Rows(i).Cells(3).Value = reader42(0).ToString
                        End While
                    End If

                    reader42.Close()

                    compare_grid.Rows(i).Cells(4).Value = If(IsNumeric(compare_grid.Rows(i).Cells(3).Value), compare_grid.Rows(i).Cells(3).Value, 0) - If(IsNumeric(compare_grid.Rows(i).Cells(2).Value), compare_grid.Rows(i).Cells(2).Value, 0)

                    If compare_grid.Rows(i).Cells(4).Value <> 0 Then
                        compare_grid.Rows(i).Cells(4).Style.BackColor = Color.CadetBlue
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Next



        End If


    End Sub

    Private Sub job_label_Click(sender As Object, e As EventArgs) Handles job_label.Click
        Job_description.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim result As DialogResult = MessageBox.Show("Do you really want to apply this changes?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then
            '--------- check boxes in Build request -----------
            Try

                Dim my_assemblies = New List(Of String)()

                '--------------  add to device -------------
                Dim cmd2 As New MySqlCommand
                cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
                cmd2.Connection = Login.Connection
                Dim reader2 As MySqlDataReader
                reader2 = cmd2.ExecuteReader

                If reader2.HasRows Then
                    While reader2.Read
                        my_assemblies.Add(reader2(0))
                    End While
                End If

                reader2.Close()
                '------------------------------------------------


                For i = 0 To PR_grid.Rows.Count - 1

                    If Convert.ToBoolean(PR_grid.Rows(i).Cells(4).Value) = True Then

                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@job", job_label.Text)
                        Create_cmd.Parameters.AddWithValue("@panel", PR_grid.Rows(i).Cells(0).Value)

                        Create_cmd.CommandText = "UPDATE Build_request.build_r  SET ready_t = 'Y' where job = @job and panel = @panel"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                    Else

                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@job", job_label.Text)
                        Create_cmd.Parameters.AddWithValue("@panel", PR_grid.Rows(i).Cells(0).Value)

                        Create_cmd.CommandText = "UPDATE Build_request.build_r  SET ready_t = ''  where job = @job and panel = @panel"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                    End If


                    '---------- panel done check -------------


                Next

                '////////////////   send APL notification to procurement and inventory and mfg
                If enable_mess = True And String.Equals(job_label.Text, "Open Project: ") = False Then


                    'write mail
                    Dim mail_n As String : mail_n = "Manufacturing update for Project " & job_label.Text & vbCrLf & vbCrLf _
                     & "Updated by: " & current_user & vbCrLf & vbCrLf _
                    & " -----  THE FOLLOWING PANELS ARE READY TO BE TESTED  -----" & vbCrLf & vbCrLf

                    For i = 0 To PR_grid.Rows.Count - 1
                        If PR_grid.Rows(i).Cells(4).Value = True Then
                            If my_assemblies.Contains(PR_grid.Rows(i).Cells(0).Value) = False Then
                                mail_n = mail_n & vbCrLf & " Panel Name:  " & PR_grid.Rows(i).Cells(0).Value & "  Description: " & PR_grid.Rows(i).Cells(1).Value & vbCrLf
                            End If
                        End If
                    Next

                    'add email addresses
                    Dim emails_addr As New List(Of String)()

                    emails_addr.Add("TBullard@atronixengineering.com")
                    ' emails_addr.Add("dshipman@atronixengineering.com")

                    ''procurement
                    emails_addr.Add("ecoy@atronixengineering.com")
                    emails_addr.Add("fvargas@atronixengineering.com")
                    ' emails_addr.Add("mmorris@atronixengineering.com")
                    '  emails_addr.Add("sowens@atronixengineering.com")

                    ''mfg
                    emails_addr.Add("shenley@atronixengineering.com")
                    emails_addr.Add("mowens@atronixengineering.com")



                    Try

                        '------- get email of the person that release this mbom ----
                        Dim cmd41 As New MySqlCommand
                        Dim email_user As String : email_user = "notfound"
                        cmd41.Parameters.AddWithValue("@job", job_label.Text)
                        cmd41.CommandText = "SELECT p1.email from parts.users as p1 INNER JOIN Material_Request.mr ON p1.username =  Material_Request.mr.released_by where Material_Request.mr.job = @job"
                        cmd41.Connection = Login.Connection
                        Dim reader41 As MySqlDataReader
                        reader41 = cmd41.ExecuteReader

                        If reader41.HasRows Then

                            While reader41.Read
                                If reader41(0).ToString.Contains("@") Then
                                    email_user = reader41(0)
                                End If
                            End While

                        End If

                        reader41.Close()


                        If String.Equals("notfound", email_user) = False Then
                            emails_addr.Add(email_user)
                        End If

                        '------------------------------------------------------


                        Dim e_mail As New MailMessage()
                        e_mail = New MailMessage()
                        e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                        For i = 0 To emails_addr.Count - 1
                            e_mail.To.Add(emails_addr.Item(i))
                        Next

                        e_mail.Subject = "Manufacturing update for Project: " & job_label.Text & " is ready"
                        e_mail.IsBodyHtml = False
                        e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

                        Smtp_Server.Send(e_mail)

                    Catch error_t As Exception
                        MsgBox(error_t.ToString)
                    End Try


                End If

                MessageBox.Show("Done")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        TabControl1.TabPages.Insert(1, TabPage1)
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.ForeColor = Color.DarkTurquoise
    End Sub

    Sub detect_empty()
        For i = 0 To PR_grid.Rows.Count - 1
            If String.Equals(PR_grid.Rows(i).Cells(2).Value, "0") = True Or String.IsNullOrEmpty(PR_grid.Rows(i).Cells(2).Value) = True Then
                PR_grid.Rows(i).DefaultCellStyle.BackColor = Color.Firebrick
            End If
        Next
    End Sub
End Class