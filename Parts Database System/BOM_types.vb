Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class BOM_types

    Public edit_panel As Boolean
    Public job_Selected As String
    Public need_by_date As String

    Public temp_panel_qty As Integer
    Public temp_panel_name As String
    Public temp_panel_desc As String

    Public shipping_ad As String

    Public Smtp_Server As New SmtpClient
    Public feature_s As Boolean
    Public feature_Field As Boolean


    Private Sub BOMDiagramToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BOMDiagramToolStripMenuItem.Click
        If String.Equals(Me.Text, "New BOM") = False Then
            Package_BOM_tree.Load_BOM_map(Me.Text, job_Selected)
            Package_BOM_tree.Visible = True
        Else
            MessageBox.Show("No BOMs found")
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Log_out = True Then
            Me.Visible = False
            BOM_init.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        '-- remove row from panel_grid
        Call Remove_row(Panel_grid)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        '----- remove row from field_grid
        Call Remove_row(field_grid)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        '----- remove row from assem_grid
        Call Remove_row(assem_grid)
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        '----- remove row from spare_grid
        Call Remove_row(sp_grid)
    End Sub

    Sub Remove_row(grid As DataGridView)
        If grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In grid.SelectedRows
                Try
                    grid.Rows.Remove(r)
                Catch ex As Exception
                    MessageBox.Show("This row cannot be deleted")
                End Try
            Next
        Else
            MessageBox.Show("Select the row you want to delete.")
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        edit_panel = False

        If Panel_grid.Rows.Count > 0 Then
            Call Save_progress(Panel_grid)
        End If


        Enter_Panel_info.ShowDialog()
        Call populate_panel_bom()

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        '-- changing selection panel
        ' Call Save_progress(Panel_grid)

        If String.Equals("Current Job:  ", job_label.Text) = True Then
            Job_assign.ShowDialog()
        End If

        If String.Equals(shipping_ad, "") = True Then
            Enter_Address.ShowDialog()
        End If

        If String.Equals("", job_Selected) = False Then
            Call Save_p_bom(job_Selected, True)
            Call Combine_bom()
            Call Save_p_bom(job_Selected, True)
            ' Call populate_panel_bom()
        End If


        '----------------------------------

        'populate panel_grid with selected panel bom
        Panel_grid.Rows.Clear()
        Panel_grid.AllowUserToAddRows = True


        Dim my_p As String : my_p = "xx"
        Dim panel_s As String : panel_s = ""

        Try

            panel_s = ComboBox1.Text

            '-- fill date from mr
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@MBOM", Me.Text)
            cmd4.Parameters.AddWithValue("@Panel_name", ComboBox1.Text)
            cmd4.CommandText = "SELECT Panel_name, Panel_desc, Panel_qty, mr_name from Material_Request.mr where MBOM = @MBOM and Panel_name = @Panel_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    temp_panel_name = reader4(0).ToString
                    temp_panel_desc = reader4(1).ToString
                    temp_panel_qty = reader4(2).ToString
                    my_p = reader4(3).ToString

                    Panel_n_l.Text = "Name:  " & reader4(0).ToString
                    qty_l.Text = "Qty: " & reader4(2).ToString
                End While
            End If

            reader4.Close()

            '-- fill data from mr_data

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", my_p)
            cmd41.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type,  Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                Dim i As Integer : i = 0
                While reader41.Read
                    Panel_grid.Rows.Add(New String() {})
                    Panel_grid.Rows(i).Cells(0).Value = reader41(0).ToString
                    Panel_grid.Rows(i).Cells(1).Value = reader41(1).ToString
                    Panel_grid.Rows(i).Cells(2).Value = reader41(2).ToString
                    Panel_grid.Rows(i).Cells(3).Value = reader41(3).ToString
                    Panel_grid.Rows(i).Cells(4).Value = reader41(4).ToString
                    Panel_grid.Rows(i).Cells(5).Value = reader41(5).ToString
                    Panel_grid.Rows(i).Cells(6).Value = reader41(6).ToString
                    Panel_grid.Rows(i).Cells(7).Value = reader41(7).ToString
                    Panel_grid.Rows(i).Cells(8).Value = reader41(8).ToString
                    Panel_grid.Rows(i).Cells(9).Value = reader41(9).ToString
                    Panel_grid.Rows(i).Cells(10).Value = reader41(10).ToString
                    Panel_grid.Rows(i).Cells(11).Value = reader41(11).ToString

                    i = i + 1
                End While
            End If

            reader41.Close()

            Call populate_panel_bom()



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Sub Save_progress(grid As DataGridView)

        Dim result As DialogResult = MessageBox.Show("Would you like to save your changes?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            '--- check if job has been assigned
            If String.Equals("Current Job:  ", job_label.Text) = True Then
                Job_assign.ShowDialog()
            End If

            If String.Equals(shipping_ad, "") = True Then
                Enter_Address.ShowDialog()
            End If

            If String.Equals("", job_Selected) = False Then

                ' Dim exist_c As Boolean : exist_c = False

                'Try
                '    Dim check_cmd As New MySqlCommand
                '    check_cmd.Parameters.AddWithValue("@mr_name", job_Selected & "_Panel_" & temp_panel_name)
                '    check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                '    check_cmd.Connection = Login.Connection
                '    check_cmd.ExecuteNonQuery()

                '    Dim reader As MySqlDataReader
                '    reader = check_cmd.ExecuteReader

                '    If reader.HasRows Then
                '        exist_c = True
                '    End If

                '    reader.Close()

                'Catch ex As Exception
                '    MessageBox.Show(ex.ToString)
                'End Try

                '' save BOM package
                'If exist_c = False Then
                '    Call Save_p_bom(job_Selected, True)
                '    Call Combine_bom()
                '    Call Save_p_bom(job_Selected, True)
                'Else
                '    MessageBox.Show("There is already a Panel with that name!")
                'End If
                Call Save_p_bom(job_Selected, True)
                Call Combine_bom()
                Call Save_p_bom(job_Selected, True)

            End If

        End If
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        If String.Equals(Panel_n_l.Text, "Name:   Unknown") = False Then
            edit_panel = True

            Enter_Panel_info.desc_box.Text = temp_panel_desc
            Enter_Panel_info.q_box.Text = temp_panel_qty

            Enter_Panel_info.ShowDialog()
            Call populate_panel_bom()
        End If

    End Sub

    Sub export_Excel(grid As DataGridView, type_t As String)

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
                For i As Integer = 0 To grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = grid.Columns(i).HeaderText
                    For j As Integer = 0 To grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                'Dim valid_mr_name As String = Me.Text

                'For Each c In Path.GetInvalidFileNameChars()
                '    If valid_mr_name.Contains(c) Then
                '        valid_mr_name = valid_mr_name.Replace(c, "")
                '    End If
                'Next

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Table_data_" & type_t & ".xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Table exported successfully!")
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call export_Excel(Panel_grid, "Panel")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Call export_Excel(field_grid, "Field")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Call export_Excel(assem_grid, "Assembly")
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Call export_Excel(sp_grid, "Spare_Parts")
    End Sub

    Private Sub ExportMasterBOMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMasterBOMToolStripMenuItem.Click

        Dim valid_mr_name As String = Me.Text

        For Each c In Path.GetInvalidFileNameChars()
            If valid_mr_name.Contains(c) Then
                valid_mr_name = valid_mr_name.Replace(c, "")
            End If
        Next


        Call export_Excel(m_grid, valid_mr_name)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("Panel")
        Part_Picker1.specify(Panel_grid)
        Part_Picker1.qty_box.Text = 0

    End Sub

    Private Sub BOM_types_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            'Part_Picker1.Width = 1
            ' Part_Picker1.Height = 1
            Me.Close()


        End If
    End Sub

    Sub Combine_bom()
        '------- Combine panel, field, assem and spare boms into one MBOM

        m_grid.Rows.Clear()

        '-- panel bom add
        Dim my_panels = New List(Of String)()

        Try

            '--------------  add to device
            Dim cmd12 As New MySqlCommand
            cmd12.Parameters.AddWithValue("@job", job_Selected)
            cmd12.Parameters.AddWithValue("@MBOM", Me.Text)
            cmd12.CommandText = "SELECT mr_name from Material_Request.mr where released is null and MBOM = @MBOM and BOM_type = 'Panel' and job = @job"
            cmd12.Connection = Login.Connection
            Dim reader12 As MySqlDataReader
            reader12 = cmd12.ExecuteReader

            If reader12.HasRows Then
                While reader12.Read
                    my_panels.Add(reader12(0))
                End While
            End If

            reader12.Close()

            For i = 0 To my_panels.Count - 1

                Dim qty_pa As Integer : qty_pa = 1

                '-----///////// if its a panel bom get the number /////////--------
                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", my_panels.Item(i).ToString)
                cmd419.CommandText = "SELECT Panel_qty ,BOM_type from Material_Request.mr where mr_name = @mr_name"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read
                        If String.Equals("Panel", reader419(1)) = True Then

                            qty_pa = If(reader419(0) Is DBNull.Value, 1, reader419(0))

                        End If
                    End While
                End If

                reader419.Close()


                '----------------------------------------------------

                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", my_panels.Item(i).ToString)
                cmd41.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        m_grid.Rows.Add(New String() {reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, reader41(3).ToString, reader41(4).ToString, reader41(5).ToString * qty_pa, reader41(6).ToString, reader41(7).ToString, reader41(8).ToString, reader41(9).ToString, reader41(10).ToString, reader41(11).ToString})
                    End While
                End If

                reader41.Close()
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        '--- field bom add
        For i = 0 To field_grid.Rows.Count - 1

            If String.IsNullOrEmpty(field_grid.Rows(i).Cells(0).Value) = False Then
                m_grid.Rows.Add(New String() {field_grid.Rows(i).Cells(0).Value, field_grid.Rows(i).Cells(1).Value, field_grid.Rows(i).Cells(2).Value, field_grid.Rows(i).Cells(3).Value, field_grid.Rows(i).Cells(4).Value, field_grid.Rows(i).Cells(5).Value, field_grid.Rows(i).Cells(6).Value, field_grid.Rows(i).Cells(7).Value, field_grid.Rows(i).Cells(8).Value, field_grid.Rows(i).Cells(9).Value, field_grid.Rows(i).Cells(10).Value, field_grid.Rows(i).Cells(11).Value})
            End If

        Next

        '--- assem bom add
        For i = 0 To assem_grid.Rows.Count - 1

            If String.IsNullOrEmpty(assem_grid.Rows(i).Cells(0).Value) = False Then
                m_grid.Rows.Add(New String() {assem_grid.Rows(i).Cells(0).Value, assem_grid.Rows(i).Cells(1).Value, assem_grid.Rows(i).Cells(2).Value, assem_grid.Rows(i).Cells(3).Value, assem_grid.Rows(i).Cells(4).Value, assem_grid.Rows(i).Cells(5).Value, assem_grid.Rows(i).Cells(6).Value, assem_grid.Rows(i).Cells(7).Value, assem_grid.Rows(i).Cells(8).Value, assem_grid.Rows(i).Cells(9).Value, assem_grid.Rows(i).Cells(10).Value, assem_grid.Rows(i).Cells(11).Value})
            End If

        Next

        '---- spare parts bom add
        For i = 0 To sp_grid.Rows.Count - 1

            If String.IsNullOrEmpty(sp_grid.Rows(i).Cells(0).Value) = False Then
                m_grid.Rows.Add(New String() {sp_grid.Rows(i).Cells(0).Value, sp_grid.Rows(i).Cells(1).Value, sp_grid.Rows(i).Cells(2).Value, sp_grid.Rows(i).Cells(3).Value, sp_grid.Rows(i).Cells(4).Value, sp_grid.Rows(i).Cells(5).Value, sp_grid.Rows(i).Cells(6).Value, sp_grid.Rows(i).Cells(7).Value, sp_grid.Rows(i).Cells(8).Value, sp_grid.Rows(i).Cells(9).Value, sp_grid.Rows(i).Cells(10).Value, sp_grid.Rows(i).Cells(11).Value})
            End If

        Next

        '------------- merge the parts of same name ---------
        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Subtotal", GetType(String))
        dimen_table.Columns.Add("mfg_type", GetType(String))
        dimen_table.Columns.Add("Part_status", GetType(String))
        dimen_table.Columns.Add("Part_Type", GetType(String))
        dimen_table.Columns.Add("Need_by_date", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))

        For i = 0 To m_grid.Rows.Count - 1
            dimen_table.Rows.Add(m_grid.Rows(i).Cells(0).Value, m_grid.Rows(i).Cells(1).Value, m_grid.Rows(i).Cells(2).Value, m_grid.Rows(i).Cells(3).Value, m_grid.Rows(i).Cells(4).Value, m_grid.Rows(i).Cells(5).Value, m_grid.Rows(i).Cells(6).Value, m_grid.Rows(i).Cells(7).Value, m_grid.Rows(i).Cells(8).Value, m_grid.Rows(i).Cells(9).Value, m_grid.Rows(i).Cells(10).Value, m_grid.Rows(i).Cells(11).Value)
        Next

        m_grid.Rows.Clear()

        For i = 0 To dimen_table.Rows.Count - 1
            '-- combine

            Dim found_t As Boolean : found_t = False

            For j = 0 To m_grid.Rows.Count - 1

                If String.IsNullOrEmpty(dimen_table.Rows(i).Item(0)) = False Then
                    If String.Equals(dimen_table.Rows(i).Item(0).ToString, m_grid.Rows(j).Cells(0).Value) = True Then
                        m_grid.Rows(j).Cells(5).Value = If(IsNumeric(m_grid.Rows(j).Cells(5).Value), m_grid.Rows(j).Cells(5).Value, 0) + CType(dimen_table.Rows(i).Item(5).ToString, Double)
                        found_t = True
                        Exit For
                    End If
                End If

            Next

            If found_t = False Then
                m_grid.Rows.Add(New String() {dimen_table.Rows(i).Item(0).ToString, dimen_table.Rows(i).Item(1).ToString, dimen_table.Rows(i).Item(2).ToString, dimen_table.Rows(i).Item(3).ToString, dimen_table.Rows(i).Item(4).ToString, dimen_table.Rows(i).Item(5).ToString, dimen_table.Rows(i).Item(6).ToString, dimen_table.Rows(i).Item(7).ToString, dimen_table.Rows(i).Item(8).ToString, dimen_table.Rows(i).Item(9).ToString, dimen_table.Rows(i).Item(10).ToString, dimen_table.Rows(i).Item(11).ToString})
            End If
        Next


    End Sub

    Private Sub m_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles m_grid.CellValueChanged
        For Each row As DataGridViewRow In m_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(4).Value) = True And IsNumeric(row.Cells(5).Value)) Then
                row.Cells(6).Value = row.Cells(4).Value * row.Cells(5).Value
            End If
        Next
    End Sub

    Private Sub sp_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles sp_grid.CellValueChanged
        For Each row As DataGridViewRow In sp_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(4).Value) = True And IsNumeric(row.Cells(5).Value)) Then
                row.Cells(6).Value = row.Cells(4).Value * row.Cells(5).Value
            End If
        Next

        '  Call Combine_bom()
    End Sub

    Private Sub assem_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles assem_grid.CellValueChanged
        For Each row As DataGridViewRow In assem_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(4).Value) = True And IsNumeric(row.Cells(5).Value)) Then
                row.Cells(6).Value = row.Cells(4).Value * row.Cells(5).Value
            End If
        Next

        ' Call Combine_bom()
    End Sub

    Private Sub field_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles field_grid.CellValueChanged
        For Each row As DataGridViewRow In field_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(4).Value) = True And IsNumeric(row.Cells(5).Value)) Then
                row.Cells(6).Value = row.Cells(4).Value * row.Cells(5).Value
            End If
        Next

        '  Call Combine_bom()
    End Sub

    Private Sub Panel_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Panel_grid.CellValueChanged
        For Each row As DataGridViewRow In Panel_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(4).Value) = True And IsNumeric(row.Cells(5).Value)) Then
                row.Cells(6).Value = row.Cells(4).Value * row.Cells(5).Value
            End If
        Next

        'Call Combine_bom()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'part selector field
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("Field")
        Part_Picker1.specify(field_grid)
        Part_Picker1.qty_box.Text = 0
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        '-part selector assem
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("Assembly")
        Part_Picker1.specify(assem_grid)
        Part_Picker1.qty_box.Text = 0
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        '--- part spare order
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("x")
        Part_Picker1.specify(sp_grid)
        Part_Picker1.qty_box.Text = 0
    End Sub

    Private Sub GenerateMBOMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateMBOMToolStripMenuItem.Click

        If String.Equals(Me.Text, "New BOM") = False Then
            Call Combine_bom()
        Else
            MessageBox.Show("Save BOMs to generate Master BOM")
        End If
    End Sub

    Private Sub Panel_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles Panel_grid.RowsAdded
        '  Call Combine_bom()
    End Sub

    Private Sub Panel_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles Panel_grid.RowsRemoved
        '   Call Combine_bom()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If m_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(m_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        need_by_date = DateTimePicker1.Value

    End Sub

    Private Sub BOM_types_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TabControl1.TabPages.Remove(TabPage3) 'remove cost tab


        feature_s = False
        feature_field = False

        job_Selected = ""
        temp_panel_desc = ""
        temp_panel_name = ""
        temp_panel_qty = 1
        need_by_date = DateTimePicker1.Value

        shipping_ad = ""

        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"

        Call EnableDoubleBuffered(sp_grid)
        Call EnableDoubleBuffered(m_grid)
        Call EnableDoubleBuffered(field_grid)
        Call EnableDoubleBuffered(Panel_grid)
        Call EnableDoubleBuffered(assem_grid)


    End Sub

    Private Sub AssignJobNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssignJobNumberToolStripMenuItem.Click
        Job_assign.ShowDialog()
    End Sub

    Sub Import_data_excel(mygrid As DataGridView, mfg_type As String)
        '----------- Import Material Request from excel file -------------

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*ods;"

            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)

                Dim i As Integer : i = 2

                While (wb.ActiveSheet.Cells(i, 1).Value IsNot Nothing)

                    Call Get_part_data(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 2).value, wb.Worksheets(1).Cells(i, 3).value, wb.Worksheets(1).Cells(i, 4).value, wb.Worksheets(1).Cells(i, 5).value, wb.Worksheets(1).Cells(i, 6).value, wb.Worksheets(1).Cells(i, 7).value, mfg_type, wb.Worksheets(1).Cells(i, 9).value, wb.Worksheets(1).Cells(i, 10).value, wb.Worksheets(1).Cells(i, 11).value, mygrid)
                    i = i + 1

                End While
                '---------------------------------------------
                'get latest cost
                For i = 0 To mygrid.Rows.Count - 1
                    If mygrid.Rows(i).DefaultCellStyle.BackColor = Color.CadetBlue Then
                        mygrid.Rows(i).Cells(4).Value = "$" & mygrid.Rows(i).Cells(4).Value
                    Else
                        mygrid.Rows(i).Cells(4).Value = "$" & Form1.Get_Latest_Cost(Login.Connection, mygrid.Rows(i).Cells(0).Value, mygrid.Rows(i).Cells(3).Value)
                    End If
                Next
                '-------------------------------------------


                wb.Close(False)
                MessageBox.Show("Done")

            End If
        End If

    End Sub

    Sub Get_part_data(part As String, desc As String, mfg As String, vendor As String, cost As String, qty As Double, sub_t As String, mfg_t As String, status As String, type_t As String, notes_t As String, mygrid As DataGridView)

        Dim index As Integer : index = mygrid.Rows.Count - 1

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part", part)
            cmd4.CommandText = "Select Part_Description,  Manufacturer, Primary_Vendor, MFG_type, Part_Status, Part_Type  from parts_table where Part_Name = @part"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            'if part is found copy the data from parts table else just copy data from excel file
            If reader4.HasRows Then
                While reader4.Read
                    mygrid.Rows.Add(New String() {})
                    mygrid.Rows(index).Cells(0).Value = part
                    mygrid.Rows(index).Cells(1).Value = reader4(0)
                    mygrid.Rows(index).Cells(2).Value = reader4(1)
                    mygrid.Rows(index).Cells(3).Value = reader4(2)
                    mygrid.Rows(index).Cells(5).Value = qty

                    If String.Equals(mfg_t, "none") = False Then
                        mygrid.Rows(index).Cells(7).Value = mfg_t
                    Else
                        mygrid.Rows(index).Cells(7).Value = reader4(3)
                    End If


                    mygrid.Rows(index).Cells(8).Value = reader4(4)
                    mygrid.Rows(index).Cells(9).Value = reader4(5)
                    mygrid.Rows(index).Cells(10).Value = notes_t
                End While

            Else
                mygrid.Rows.Add(New String() {})
                mygrid.Rows(index).DefaultCellStyle.BackColor = Color.CadetBlue
                mygrid.Rows(index).Cells(0).Value = part
                mygrid.Rows(index).Cells(1).Value = desc
                mygrid.Rows(index).Cells(2).Value = mfg
                mygrid.Rows(index).Cells(3).Value = vendor
                mygrid.Rows(index).Cells(4).Value = cost
                mygrid.Rows(index).Cells(5).Value = qty
                mygrid.Rows(index).Cells(6).Value = sub_t

                If String.Equals(mfg_t, "Panel") = True Or String.Equals(mfg_t, "Field") = True Or String.Equals(mfg_t, "Assembly") = True Or String.Equals(mfg_t, "Bulk") = True Then
                    mygrid.Rows(index).Cells(7).Value = mfg_t
                Else
                    mygrid.Rows(index).Cells(7).Value = "Field"
                End If

                mygrid.Rows(index).Cells(8).Value = "Special Order"
                mygrid.Rows(index).Cells(9).Value = "Other"
                mygrid.Rows(index).Cells(10).Value = notes_t

            End If

            reader4.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    '------ this subroutine check for special orders and add them to the main database if no error is found
    Function Special_order_check(mygrid As DataGridView) As Boolean

        '-- returns true if no invalid part was found

        Special_order_check = True

        Try
            For i = 0 To mygrid.Rows.Count - 1

                If mygrid.Rows(i).IsNewRow Then Continue For

                If String.Equals(mygrid.Rows(i).Cells(0).Value, "") = False Then

                    Dim exist_c As Boolean : exist_c = False
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part", mygrid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "Select * from parts_table where Part_Name = @part"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    'if part is found update boolean
                    If reader4.HasRows Then
                        While reader4.Read
                            exist_c = True
                        End While
                    End If

                    reader4.Close()

                    Dim first_pass As Boolean : first_pass = False
                    Dim last_pass As Boolean : last_pass = False

                    If exist_c = False Then

                        '-- check if its a valid part
                        If String.IsNullOrEmpty(mygrid.Rows(i).Cells(0).Value) = False And String.Equals(mygrid.Rows(i).Cells(0).Value, "") = False Then 'check for valid part name 
                            If String.IsNullOrEmpty(mygrid.Rows(i).Cells(2).Value) = False And String.Equals(mygrid.Rows(i).Cells(2).Value, "") = False Then 'check for valid manuf
                                first_pass = True
                            End If
                        End If

                        ' check if MFG Type is valid
                        If first_pass = True Then

                            If String.Equals(mygrid.Rows(i).Cells(7).Value, "Panel") = True Or String.Equals(mygrid.Rows(i).Cells(7).Value, "Field") = True Or String.Equals(mygrid.Rows(i).Cells(7).Value, "Assembly") = True Or String.Equals(mygrid.Rows(i).Cells(7).Value, "Bulk") = True Then
                                last_pass = True

                                'start inserting data
                                Dim Create_cmd As New MySqlCommand
                                Create_cmd.Parameters.AddWithValue("@Part_Name", mygrid.Rows(i).Cells(0).Value)
                                Create_cmd.Parameters.AddWithValue("@Manufacturer", mygrid.Rows(i).Cells(2).Value)
                                Create_cmd.Parameters.AddWithValue("@Part_Description", mygrid.Rows(i).Cells(1).Value)
                                Create_cmd.Parameters.AddWithValue("@Part_Status", "Special Order")
                                Create_cmd.Parameters.AddWithValue("@Part_Type", "Other")
                                Create_cmd.Parameters.AddWithValue("@Primary_Vendor", mygrid.Rows(i).Cells(3).Value)
                                Create_cmd.Parameters.AddWithValue("@MFG_type", mygrid.Rows(i).Cells(7).Value)


                                Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description,  Part_Status, Part_Type, Primary_Vendor, MFG_type) VALUES (@Part_Name, @Manufacturer, @Part_Description, @Part_Status, @Part_Type, @Primary_Vendor, @MFG_type)"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()

                                '--- insert into inventory
                                Dim Create_cmdi As New MySqlCommand
                                Create_cmdi.Parameters.AddWithValue("@part_name", mygrid.Rows(i).Cells(0).Value)
                                Create_cmdi.Parameters.AddWithValue("@description", mygrid.Rows(i).Cells(1).Value)
                                Create_cmdi.Parameters.AddWithValue("@manufacturer", mygrid.Rows(i).Cells(2).Value)
                                Create_cmdi.Parameters.AddWithValue("@MFG_type", mygrid.Rows(i).Cells(7).Value)


                                Create_cmdi.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, 0,0,0,0)"
                                Create_cmdi.Connection = Login.Connection
                                Create_cmdi.ExecuteNonQuery()

                                '----------- insert vendor and cost -----
                                If IsNumeric(mygrid.Rows(i).Cells(4).Value) = True And String.IsNullOrEmpty(mygrid.Rows(i).Cells(3).Value) = False And String.Equals(mygrid.Rows(i).Cells(3).Value, "") = False Then
                                    Dim main_cmd2 As New MySqlCommand
                                    main_cmd2.Parameters.AddWithValue("@Part_Name", mygrid.Rows(i).Cells(0).Value)
                                    main_cmd2.Parameters.AddWithValue("@Vendor_Name", mygrid.Rows(i).Cells(3).Value)
                                    main_cmd2.Parameters.AddWithValue("@Cost", If(IsNumeric(mygrid.Rows(i).Cells(4).Value.ToString.Replace("$", "")) = True, mygrid.Rows(i).Cells(4).Value.ToString.Replace("$", ""), 0))
                                    main_cmd2.CommandText = "INSERT INTO vendors_table(Part_Name, Vendor_Name, Cost, Purchase_Date) VALUES (@Part_Name, @Vendor_Name, @Cost, now())"
                                    main_cmd2.Connection = Login.Connection
                                    main_cmd2.ExecuteNonQuery()
                                End If

                            End If
                        End If

                        If last_pass = False Then
                            mygrid.Rows(i).DefaultCellStyle.BackColor = Color.Firebrick
                            Special_order_check = False
                        Else
                            If i Mod 2 = 0 Then
                                mygrid.Rows(i).DefaultCellStyle.BackColor = Color.WhiteSmoke
                            Else
                                mygrid.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
                            End If

                            Special_order_check = True
                        End If

                    Else  'if part exist
                        Special_order_check = True

                    End If

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '--- import data to panel table
        If Panel_grid.Rows.Count > 0 Then

            Call Import_data_excel(Panel_grid, "Panel")

            If Special_order_check(Panel_grid) = False Then
                MessageBox.Show("Invalid Parts found! Please check the red rows")
            End If

        Else
            MessageBox.Show("Please, select a Panel")
        End If
    End Sub

    Private Sub CheckInvalidPartsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CheckInvalidPartsToolStripMenuItem.Click

        If ok_to_save() = False Then
            MessageBox.Show("Invalid Parts found! Please check the Red rows")
        Else
            MessageBox.Show("No Invalid Parts were found")
        End If

    End Sub


    Function ok_to_save() As Boolean

        ok_to_save = True


        If Special_order_check(Panel_grid) = False Then
            ok_to_save = True
        End If

        If Special_order_check(field_grid) = False Then
            ok_to_save = True
        End If

        If Special_order_check(assem_grid) = False Then
            ok_to_save = True
        End If

        If Special_order_check(sp_grid) = False Then
            ok_to_save = True
        End If


    End Function



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '--- import field
        If field_grid.Rows.Count > 0 Then

            Call Import_data_excel(field_grid, "Field")

            If Special_order_check(field_grid) = False Then
                MessageBox.Show("Please check the red rows")
            End If

        Else
            MessageBox.Show("Please, select a Panel")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '--- import assembly
        If assem_grid.Rows.Count > 0 Then

            Call Import_data_excel(assem_grid, "Assembly")

            If Special_order_check(assem_grid) = False Then
                MessageBox.Show("Please check the red rows")
            End If

        Else
            MessageBox.Show("Please, select a Panel")
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '--- import spare parts
        If assem_grid.Rows.Count > 0 Then

            Call Import_data_excel(sp_grid, "none")

            If Special_order_check(sp_grid) = False Then
                MessageBox.Show("Please check the red rows")
            End If

        Else
            MessageBox.Show("Please, select a Panel")
        End If
    End Sub

    Private Sub field_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles field_grid.RowsAdded
        'Call Combine_bom()
    End Sub

    Private Sub field_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles field_grid.RowsRemoved
        ' Call Combine_bom()
    End Sub

    Private Sub assem_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles assem_grid.RowsAdded
        ' Call Combine_bom()
    End Sub

    Private Sub assem_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles assem_grid.RowsRemoved
        ' Call Combine_bom()
    End Sub

    Private Sub sp_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles sp_grid.RowsAdded
        ' Call Combine_bom()
    End Sub

    Private Sub sp_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles sp_grid.RowsRemoved
        ' Call Combine_bom()
    End Sub

    '====== This function save a bom to mr table and mr_data table

    Sub Save_bom(mygrid As DataGridView, mr_name As String, job As String, BOM_type As String, enable_check As Boolean, panel_name As String, panel_desc As String, panel_qty As Integer, MBOM As String)

        Dim proceed_m As Boolean : proceed_m = True
        Dim exist_c As Boolean : exist_c = False


        If enable_check = True Then
            proceed_m = ok_to_save()
        End If


        If proceed_m = True Then

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                    check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        exist_c = True
                    End If

                    reader.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                If exist_c = False Then

                    Try
                        '--- enter data to mr -------
                        Dim main_cmd As New MySqlCommand
                        main_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                        main_cmd.Parameters.AddWithValue("@created_by", current_user)
                        main_cmd.Parameters.AddWithValue("@job", job)
                        main_cmd.Parameters.AddWithValue("@BOM_type", BOM_type)
                        main_cmd.Parameters.AddWithValue("@Panel_name", panel_name)
                        main_cmd.Parameters.AddWithValue("@Panel_desc", panel_desc)
                        main_cmd.Parameters.AddWithValue("@Panel_qty", panel_qty)
                        main_cmd.Parameters.AddWithValue("@MBOM", MBOM)
                        main_cmd.Parameters.AddWithValue("@need_date", need_by_date)
                        main_cmd.Parameters.AddWithValue("@shipping_ad", shipping_ad)

                        main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, job, BOM_type, Panel_name, Panel_desc, Panel_qty, MBOM, need_date, shipping_ad) VALUES (@mr_name, now(), @created_by, @job, @BOM_type, @Panel_name, @Panel_desc, @Panel_qty, @MBOM, @need_date, @shipping_Ad)"
                        main_cmd.Connection = Login.Connection
                        main_cmd.ExecuteNonQuery()


                        '-------- enter data to mr_data
                        For i = 0 To mygrid.Rows.Count - 1

                        If IsNumeric(mygrid.Rows(i).Cells(5).Value) = True And mygrid.Rows(i).Cells(0).Value Is Nothing = False Then

                            Dim Create_cmd6 As New MySqlCommand
                            Create_cmd6.Parameters.Clear()
                            Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name)
                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(mygrid.Rows(i).Cells(0).Value Is Nothing, "unknown part", mygrid.Rows(i).Cells(0).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Description", If(mygrid.Rows(i).Cells(1).Value Is Nothing, "", mygrid.Rows(i).Cells(1).Value.ToString.Trim))
                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(mygrid.Rows(i).Cells(2).Value Is Nothing, "", mygrid.Rows(i).Cells(2).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(mygrid.Rows(i).Cells(3).Value Is Nothing, "", mygrid.Rows(i).Cells(3).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Price", If(mygrid.Rows(i).Cells(4).Value Is Nothing, "", mygrid.Rows(i).Cells(4).Value.ToString.Replace("$", "")))
                            Create_cmd6.Parameters.AddWithValue("@Qty", If(mygrid.Rows(i).Cells(5).Value Is Nothing, "", mygrid.Rows(i).Cells(5).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@subtotal", If(mygrid.Rows(i).Cells(6).Value Is Nothing, "", mygrid.Rows(i).Cells(6).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(mygrid.Rows(i).Cells(7).Value Is Nothing, "Panel", mygrid.Rows(i).Cells(7).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Part_status", If(mygrid.Rows(i).Cells(8).Value Is Nothing, "Special Order", mygrid.Rows(i).Cells(8).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Part_type", If(mygrid.Rows(i).Cells(9).Value Is Nothing, "Other", mygrid.Rows(i).Cells(9).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@Notes", If(mygrid.Rows(i).Cells(10).Value Is Nothing, "", mygrid.Rows(i).Cells(10).Value.ToString))
                            Create_cmd6.Parameters.AddWithValue("@need_by_date", If(mygrid.Rows(i).Cells(11).Value Is Nothing, "", mygrid.Rows(i).Cells(11).Value.ToString))

                            If String.Equals(BOM_type, "MB") = True Then
                                Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes,  need_by_date, latest_r) VALUES (@mr_name, @Part_No, @Description,  @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type,  @Part_status, @Part_type, @Notes,  @need_by_date, 'x')"
                            Else
                                Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes,  need_by_date) VALUES (@mr_name, @Part_No, @Description,  @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type,  @Part_status, @Part_type, @Notes,  @need_by_date)"
                            End If


                            ' Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes,  need_by_date) VALUES (@mr_name, @Part_No, @Description,  @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type,  @Part_status, @Part_type, @Notes,  @need_by_date)"
                            Create_cmd6.Connection = Login.Connection
                            Create_cmd6.ExecuteNonQuery()

                        End If
                    Next


                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                Else
                    MessageBox.Show("BOM already exist!")
                End If

            Else
                MessageBox.Show("Invalid Special Order Parts found!")
            End If

    End Sub

    '====== This function delete a bom to mr table and mr_data table
    Sub delete_bom(mr_name As String)

        Dim go_delete As Boolean : go_delete = True

        Try

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd2.CommandText = "delete from Material_Request.mr where mr_name = @mr_name"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim check_cmd3 As New MySqlCommand
            check_cmd3.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd3.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
            check_cmd3.Connection = Login.Connection
            check_cmd3.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    '------This function save an entire BOM package

    Sub Save_p_bom(job As String, enable_check As Boolean)

        '============ save MBOM ===================
        'first delete the existing MBOM

        Dim mbom_name As String : mbom_name = job & "_Master_BOM"
        Call delete_bom(mbom_name)
        Call Save_bom(m_grid, mbom_name, job, "MB", True, "", "", 0, "")

        '============ save Panel BOM ===============
        If String.Equals(temp_panel_name, "") = False Then

            Dim panel_bom_name As String : panel_bom_name = job & "_Panel_" & temp_panel_name
            Call delete_bom(panel_bom_name)
            Call Save_bom(Panel_grid, panel_bom_name, job, "Panel", True, temp_panel_name, temp_panel_desc, temp_panel_qty, mbom_name)

        End If

        '============ save field BOM ===============
        Dim field_bom_name As String : field_bom_name = job & "_Field"
        Call delete_bom(field_bom_name)
        Call Save_bom(field_grid, field_bom_name, job, "Field", True, "", "", 0, mbom_name)

        '============ save Assem BOMs ================

        Call ASSEM_bom(job, mbom_name)

        '============ save Special Order BOM ==========
        Dim sp_bom_name As String : sp_bom_name = job & "_Spare_Parts"
        Call delete_bom(sp_bom_name)
        Call Save_bom(sp_grid, sp_bom_name, job, "Spare_Parts", True, "", "", 0, mbom_name)

        Me.Text = mbom_name


    End Sub

    Sub ASSEM_bom(job As String, mbom_name As String)

        '--- create assemblies bom
        Try
            Dim my_assem_bom = New List(Of String)()

            '--------------  add to device
            Dim cmd2 As New MySqlCommand
            cmd2.Parameters.AddWithValue("@MBOM", mbom_name)
            cmd2.Parameters.AddWithValue("@job", job)
            cmd2.CommandText = "SELECT mr_name from Material_Request.mr where released is null and MBOM = @MBOM and BOM_type = 'Assembly' and job = @job"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assem_bom.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()

            '--- delete all assemblies bom not released and associated with mbom and job
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@MBOM", mbom_name)
            check_cmd2.Parameters.AddWithValue("@job", job)
            check_cmd2.CommandText = "delete from Material_Request.mr where released is null and MBOM = @MBOM and BOM_type = 'Assembly' and job = @job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()


            For i = 0 To my_assem_bom.Count - 1
                Dim check_cmd3 As New MySqlCommand
                check_cmd3.Parameters.Clear()
                check_cmd3.Parameters.AddWithValue("@mr_name", my_assem_bom.Item(i).ToString)
                check_cmd3.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
                check_cmd3.Connection = Login.Connection
                check_cmd3.ExecuteNonQuery()
            Next


            '--- loop through assembly grid and create an assemblie bom per row
            For i = 0 To assem_grid.Rows.Count - 1

                    If assem_grid.Rows(i).IsNewRow Then Continue For

                Dim exist_c As Boolean = False
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.Clear()
                    check_cmd.Parameters.AddWithValue("@mr_name", job & "_Assembly_" & assem_grid.Rows(i).Cells(0).Value)
                    check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        exist_c = True
                    End If

                reader.Close()

                If exist_c = False Then

                    '--- enter data to mr -------
                    Dim main_cmd As New MySqlCommand
                    main_cmd.Parameters.AddWithValue("@mr_name", job & "_Assembly_" & assem_grid.Rows(i).Cells(0).Value)
                    main_cmd.Parameters.AddWithValue("@created_by", current_user)
                    main_cmd.Parameters.AddWithValue("@job", job)
                    main_cmd.Parameters.AddWithValue("@BOM_type", "Assembly")
                    main_cmd.Parameters.AddWithValue("@Panel_name", assem_grid.Rows(i).Cells(0).Value)
                    main_cmd.Parameters.AddWithValue("@Panel_desc", assem_grid.Rows(i).Cells(1).Value)
                    main_cmd.Parameters.AddWithValue("@Panel_qty", assem_grid.Rows(i).Cells(5).Value)
                    main_cmd.Parameters.AddWithValue("@MBOM", mbom_name)
                    main_cmd.Parameters.AddWithValue("@need_by_date", need_by_date)
                    main_cmd.Parameters.AddWithValue("@shipping_ad", shipping_ad)

                    main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, job, BOM_type, Panel_name, Panel_desc, Panel_qty, MBOM, need_date, shipping_ad) VALUES (@mr_name, now(), @created_by, @job, @BOM_type, @Panel_name, @Panel_desc, @Panel_qty, @MBOM, @need_date, @shipping_ad)"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()


                    '-------- enter data to mr_data


                    If IsNumeric(assem_grid.Rows(i).Cells(5).Value) = True And assem_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                        Dim Create_cmd6 As New MySqlCommand
                        Create_cmd6.Parameters.Clear()
                        Create_cmd6.Parameters.AddWithValue("@mr_name", job & "_Assembly_" & assem_grid.Rows(i).Cells(0).Value)
                        Create_cmd6.Parameters.AddWithValue("@Part_No", If(assem_grid.Rows(i).Cells(0).Value Is Nothing, "unknown part", assem_grid.Rows(i).Cells(0).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Description", If(assem_grid.Rows(i).Cells(1).Value Is Nothing, "", assem_grid.Rows(i).Cells(1).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(assem_grid.Rows(i).Cells(2).Value Is Nothing, "", assem_grid.Rows(i).Cells(2).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Vendor", If(assem_grid.Rows(i).Cells(3).Value Is Nothing, "", assem_grid.Rows(i).Cells(3).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Price", If(assem_grid.Rows(i).Cells(4).Value Is Nothing, "", assem_grid.Rows(i).Cells(4).Value.ToString.Replace("$", "")))
                        Create_cmd6.Parameters.AddWithValue("@Qty", If(assem_grid.Rows(i).Cells(5).Value Is Nothing, "", assem_grid.Rows(i).Cells(5).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@subtotal", If(assem_grid.Rows(i).Cells(6).Value Is Nothing, "", assem_grid.Rows(i).Cells(6).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@mfg_type", If(assem_grid.Rows(i).Cells(7).Value Is Nothing, "Panel", assem_grid.Rows(i).Cells(7).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Part_status", If(assem_grid.Rows(i).Cells(8).Value Is Nothing, "Special Order", assem_grid.Rows(i).Cells(8).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Part_type", If(assem_grid.Rows(i).Cells(9).Value Is Nothing, "Other", assem_grid.Rows(i).Cells(9).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@Notes", If(assem_grid.Rows(i).Cells(10).Value Is Nothing, "", assem_grid.Rows(i).Cells(10).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@need_by_date", If(assem_grid.Rows(i).Cells(11).Value Is Nothing, "", assem_grid.Rows(i).Cells(11).Value.ToString))


                        Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes,  need_by_date) VALUES (@mr_name, @Part_No, @Description,  @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type,  @Part_status, @Part_type, @Notes,  @need_by_date)"
                        Create_cmd6.Connection = Login.Connection
                        Create_cmd6.ExecuteNonQuery()

                    End If

                End If
            Next

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

    End Sub



    Sub populate_panel_bom()
        '--- fill the panel bom dropbox
        Try
            ComboBox1.Items.Clear()

            Dim cmd_j As New MySqlCommand
            cmd_j.Parameters.AddWithValue("@MBOM", Me.Text)
            cmd_j.CommandText = "SELECT Panel_name from Material_Request.mr where BOM_type = 'Panel' and MBOM = @MBOM"
            cmd_j.Connection = Login.Connection

            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    ComboBox1.Items.Add(readerj(0))
                End While
            End If

            readerj.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub SaveBOMPackageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveBOMPackageToolStripMenuItem.Click

        If String.Equals(job_Selected, "") = False Then
            If String.Equals(shipping_ad, "") = True Then
                Enter_Address.ShowDialog()
            End If

            '--------- check for shipping address and need_by_date
            Dim weird_date As Boolean : weird_date = False

            Dim result As DialogResult = MessageBox.Show("Your Need by Date is " & need_by_date & ". Do you still want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then
                weird_date = False
            Else
                weird_date = True
            End If

            '--------------------------------------------------

            If weird_date = False Then
                Call Save_p_bom(job_Selected, True)
                Call Combine_bom()
                Call Save_p_bom(job_Selected, True)
                Call populate_panel_bom()
                MessageBox.Show("BOMs saved succesfully")
            End If
        Else
                MessageBox.Show("No Job Selected")
        End If
    End Sub



    Private Sub DeleteBOMPackageToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles DeleteBOMPackageToolStripMenuItem.Click
        If String.Equals("New BOM", Me.Text) = False And String.Equals(job_Selected, "") = False Then

            Dim result As DialogResult = MessageBox.Show("Are you sure, you want to delete this BOM package?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                Dim my_bom = New List(Of String)()

                Try
                    Dim cmd2 As New MySqlCommand
                    cmd2.Parameters.AddWithValue("@job", job_Selected)
                    cmd2.Parameters.AddWithValue("@MBOM", Me.Text)
                    cmd2.CommandText = "SELECT distinct mr_name from Material_Request.mr where released is null and MBOM = @MBOM and job = @job"
                    cmd2.Connection = Login.Connection
                    Dim reader2 As MySqlDataReader
                    reader2 = cmd2.ExecuteReader

                    If reader2.HasRows Then
                        While reader2.Read
                            my_bom.Add(reader2(0))
                        End While
                    End If

                    reader2.Close()

                    '----delete mrs related to MBOM

                    Dim check_cmd2 As New MySqlCommand
                    check_cmd2.Parameters.AddWithValue("@MBOM", Me.Text)
                    check_cmd2.Parameters.AddWithValue("@job", job_Selected)
                    check_cmd2.CommandText = "delete from Material_Request.mr where released is null and MBOM = @MBOM  and job = @job"
                    check_cmd2.Connection = Login.Connection
                    check_cmd2.ExecuteNonQuery()

                    '---------- delete MBOM itself
                    Dim check_cmd21 As New MySqlCommand
                    check_cmd21.Parameters.AddWithValue("@mr_name", Me.Text)
                    check_cmd21.Parameters.AddWithValue("@job", job_Selected)
                    check_cmd21.CommandText = "delete from Material_Request.mr where released is null and mr_name = @mr_name  and job = @job"
                    check_cmd21.Connection = Login.Connection
                    check_cmd21.ExecuteNonQuery()


                    Dim check_cmd22 As New MySqlCommand
                    check_cmd22.Parameters.AddWithValue("@mr_name", Me.Text)
                    check_cmd22.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
                    check_cmd22.Connection = Login.Connection
                    check_cmd22.ExecuteNonQuery()
                    '==============================


                    For i = 0 To my_bom.Count - 1
                        Dim check_cmd3 As New MySqlCommand
                        check_cmd3.Parameters.Clear()
                        check_cmd3.Parameters.AddWithValue("@mr_name", my_bom.Item(i))
                        check_cmd3.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
                        check_cmd3.Connection = Login.Connection
                        check_cmd3.ExecuteNonQuery()
                    Next


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                Me.Text = "New BOM"
                Panel_n_l.Text = "Name:   Unknown"
                qty_l.Text = "Qty: 1"
                ComboBox1.Items.Clear()
                job_label.Text = "Current Job:  "
                job_Selected = ""
                temp_panel_name = ""
                temp_panel_desc = ""
                temp_panel_qty = 1
                need_by_date = DateTimePicker1.Value
                shipping_ad = ""

                Panel_grid.Rows.Clear()
                field_grid.Rows.Clear()
                assem_grid.Rows.Clear()
                sp_grid.Rows.Clear()
                m_grid.Rows.Clear()

                MessageBox.Show("BOM package deleted succesfully")

            End If
        End If
    End Sub


    Sub Panel_delete(mr_name As String, job As String)
        '-- right now im using this function to delete a bom package cause the open sub is not ready yet

        Dim my_bom = New List(Of String)()

        Try
            Dim cmd2 As New MySqlCommand
            cmd2.Parameters.AddWithValue("@job", job)
            cmd2.Parameters.AddWithValue("@MBOM", mr_name)
            cmd2.CommandText = "SELECT distinct mr_name from Material_Request.mr where released is null and MBOM = @MBOM and job = @job"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_bom.Add(reader2(0))
                End While
            End If

            reader2.Close()

            '----delete mrs related to MBOM

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@MBOM", mr_name)
            check_cmd2.Parameters.AddWithValue("@job", job)
            check_cmd2.CommandText = "delete from Material_Request.mr where released is null and MBOM = @MBOM  and job = @job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            '---------- delete MBOM itself
            Dim check_cmd21 As New MySqlCommand
            check_cmd21.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd21.Parameters.AddWithValue("@job", job)
            check_cmd21.CommandText = "delete from Material_Request.mr where released is null and mr_name = @mr_name  and job = @job"
            check_cmd21.Connection = Login.Connection
            check_cmd21.ExecuteNonQuery()


            Dim check_cmd22 As New MySqlCommand
            check_cmd22.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd22.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
            check_cmd22.Connection = Login.Connection
            check_cmd22.ExecuteNonQuery()
            '==============================


            For i = 0 To my_bom.Count - 1
                Dim check_cmd3 As New MySqlCommand
                check_cmd3.Parameters.Clear()
                check_cmd3.Parameters.AddWithValue("@mr_name", my_bom.Item(i))
                check_cmd3.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
                check_cmd3.Connection = Login.Connection
                check_cmd3.ExecuteNonQuery()
            Next

            Call Combine_bom()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        MessageBox.Show("BOM package deleted succesfully")
    End Sub



    Private Sub NewBOMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewBOMToolStripMenuItem.Click

        '--- prepare for a new BOM package

        Me.Text = "New BOM"
        Panel_n_l.Text = "Name:   Unknown"
        qty_l.Text = "Qty: 1"
        ComboBox1.Items.Clear()
        job_label.Text = "Current Job:  "
        job_Selected = ""
        temp_panel_name = ""
        temp_panel_desc = ""
        temp_panel_qty = 1

        need_by_date = DateTimePicker1.Value
        shipping_ad = ""

        Panel_grid.Rows.Clear()
        field_grid.Rows.Clear()
        assem_grid.Rows.Clear()
        sp_grid.Rows.Clear()
        m_grid.Rows.Clear()

    End Sub

    Private Sub OpenBOMPackageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenBOMPackageToolStripMenuItem.Click
        Open_package.ShowDialog()
    End Sub

    Private Sub AddEditAddressToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddEditAddressToolStripMenuItem.Click
        Enter_Address.ShowDialog()
    End Sub

    Sub Open_BOM_package(mr_name As String, job As String)

        Panel_grid.Rows.Clear()
        field_grid.Rows.Clear()
        assem_grid.Rows.Clear()
        sp_grid.Rows.Clear()
        m_grid.Rows.Clear()

        '======= Open all BOMs =========
        Try
            Me.Text = mr_name
            Panel_n_l.Text = "Name:   Unknown"
            qty_l.Text = "Qty: 1"
            job_label.Text = "Current Job:  " & job
            job_Selected = job
            temp_panel_name = ""
            temp_panel_desc = ""
            temp_panel_qty = 1

            '---- get info about Master BOM
            Dim cmd2 As New MySqlCommand
            cmd2.Parameters.AddWithValue("@job", job)
            cmd2.Parameters.AddWithValue("@mr_name", mr_name)
            cmd2.CommandText = "SELECT shipping_ad, need_date from Material_Request.mr where released is null and mr_name = @mr_name and job = @job"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    shipping_ad = reader2(0).ToString
                    need_by_date = reader2(1).ToString  'it used to be reader2(1).tostring
                End While
            End If

            reader2.Close()

            '--- fill the panel combobox
            Call populate_panel_bom()
            DateTimePicker1.Value = need_by_date

            '-------- fill field bom ------------

            Dim mr_name_field As String : mr_name_field = "xxxx"

            Dim cmd31 As New MySqlCommand
            cmd31.Parameters.AddWithValue("@BOM_type", "Field")
            cmd31.Parameters.AddWithValue("@job", job)
            cmd31.Parameters.AddWithValue("@MBOM", mr_name)
            cmd31.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type and job = @job and MBOM = @MBOM"
            cmd31.Connection = Login.Connection
            Dim reader31 As MySqlDataReader
            reader31 = cmd31.ExecuteReader

            If reader31.HasRows Then
                While reader31.Read
                    mr_name_field = reader31(0).ToString
                End While
            End If

            reader31.Close()


            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@mr_name", mr_name_field)
            cmd3.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                Dim i As Integer : i = 0

                While reader3.Read
                    field_grid.Rows.Add(New String() {})
                    field_grid.Rows(i).Cells(0).Value = reader3(0).ToString
                    field_grid.Rows(i).Cells(1).Value = reader3(1).ToString
                    field_grid.Rows(i).Cells(2).Value = reader3(2).ToString
                    field_grid.Rows(i).Cells(3).Value = reader3(3).ToString
                    field_grid.Rows(i).Cells(4).Value = reader3(4).ToString
                    field_grid.Rows(i).Cells(5).Value = reader3(5).ToString
                    field_grid.Rows(i).Cells(6).Value = reader3(6).ToString
                    field_grid.Rows(i).Cells(7).Value = reader3(7).ToString
                    field_grid.Rows(i).Cells(8).Value = reader3(8).ToString
                    field_grid.Rows(i).Cells(9).Value = reader3(9).ToString
                    field_grid.Rows(i).Cells(10).Value = reader3(10).ToString
                    field_grid.Rows(i).Cells(11).Value = reader3(11).ToString

                    i = i + 1
                End While
            End If

            reader3.Close()

            '------- fill assembly bom ----------

            Dim my_assem_bom = New List(Of String)()

            '--------------  add to device
            Dim cmd12 As New MySqlCommand
            cmd12.Parameters.AddWithValue("@job", job)
            cmd12.Parameters.AddWithValue("@MBOM", mr_name)
            cmd12.CommandText = "SELECT distinct mr_name from Material_Request.mr where released is null and MBOM = @MBOM and BOM_type = 'Assembly' and job = @job"
            cmd12.Connection = Login.Connection
            Dim reader12 As MySqlDataReader
            reader12 = cmd12.ExecuteReader

            If reader12.HasRows Then
                While reader12.Read
                    my_assem_bom.Add(reader12(0))
                End While
            End If

            reader12.Close()

            For i = 0 To my_assem_bom.Count - 1

                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", my_assem_bom.Item(i).ToString)
                cmd41.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then


                    While reader41.Read
                        assem_grid.Rows.Add(New String() {reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, reader41(3).ToString, reader41(4).ToString, reader41(5).ToString, reader41(6).ToString, reader41(7).ToString, reader41(8).ToString, reader41(9).ToString, reader41(10).ToString, reader41(11).ToString})
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(0).Value = reader41(0).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(1).Value = reader41(1).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(2).Value = reader41(2).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(3).Value = reader41(3).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(4).Value = reader41(4).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(5).Value = reader41(5).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(6).Value = reader41(6).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(7).Value = reader41(7).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(8).Value = reader41(8).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(9).Value = reader41(9).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(10).Value = reader41(10).ToString
                        'assem_grid.Rows(assem_grid.Rows.Count - 1).Cells(11).Value = reader41(11).ToString

                    End While
                End If

                reader41.Close()
            Next



            '---------- fill spare parts -----------

            Dim mr_name_sp As String : mr_name_sp = "xxxx"

            Dim cmd32 As New MySqlCommand
            cmd32.Parameters.AddWithValue("@BOM_type", "Spare_Parts")
            cmd32.Parameters.AddWithValue("@job", job)
            cmd32.Parameters.AddWithValue("@MBOM", mr_name)
            cmd32.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type and job = @job and MBOM = @MBOM"
            cmd32.Connection = Login.Connection
            Dim reader32 As MySqlDataReader
            reader32 = cmd32.ExecuteReader

            If reader32.HasRows Then
                While reader32.Read
                    mr_name_sp = reader32(0).ToString
                End While
            End If

            reader32.Close()


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name_sp)
            cmd4.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0

                While reader4.Read
                    sp_grid.Rows.Add(New String() {})
                    sp_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    sp_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    sp_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    sp_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                    sp_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                    sp_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                    sp_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                    sp_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                    sp_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                    sp_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                    sp_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                    sp_grid.Rows(i).Cells(11).Value = reader4(11).ToString

                    i = i + 1
                End While
            End If

            reader4.Close()

            '---------- fill MBOM ----------------
            Dim cmd6 As New MySqlCommand
            cmd6.Parameters.AddWithValue("@mr_name", mr_name)
            cmd6.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd6.Connection = Login.Connection
            Dim reader6 As MySqlDataReader
            reader6 = cmd6.ExecuteReader

            If reader6.HasRows Then
                Dim i As Integer : i = 0

                While reader6.Read
                    m_grid.Rows.Add(New String() {})
                    m_grid.Rows(i).Cells(0).Value = reader6(0).ToString
                    m_grid.Rows(i).Cells(1).Value = reader6(1).ToString
                    m_grid.Rows(i).Cells(2).Value = reader6(2).ToString
                    m_grid.Rows(i).Cells(3).Value = reader6(3).ToString
                    m_grid.Rows(i).Cells(4).Value = reader6(4).ToString
                    m_grid.Rows(i).Cells(5).Value = reader6(5).ToString
                    m_grid.Rows(i).Cells(6).Value = reader6(6).ToString
                    m_grid.Rows(i).Cells(7).Value = reader6(7).ToString
                    m_grid.Rows(i).Cells(8).Value = reader6(8).ToString
                    m_grid.Rows(i).Cells(9).Value = reader6(9).ToString
                    m_grid.Rows(i).Cells(10).Value = reader6(10).ToString
                    m_grid.Rows(i).Cells(11).Value = reader6(11).ToString

                    i = i + 1
                End While
            End If

            reader6.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub RemoveAPanelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveAPanelToolStripMenuItem.Click

        Dim panel_name As String : panel_name = "Not_selected"
        Dim mr_name As String : mr_name = job_Selected & "_Panel_" & panel_name

        '  If Not ComboBox1.SelectedItem Is Nothing Then
        panel_name = temp_panel_name
        mr_name = job_Selected & "_Panel_" & panel_name
        '  End If

        Dim result As DialogResult = MessageBox.Show("Are you sure, you want to delete panel " & panel_name & "?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            Try
                '---------- delete ---------------
                Dim check_cmd21 As New MySqlCommand
                check_cmd21.Parameters.AddWithValue("@MBOM", Me.Text)
                check_cmd21.Parameters.AddWithValue("@job", job_Selected)
                check_cmd21.Parameters.AddWithValue("@mr_name", mr_name)
                check_cmd21.CommandText = "delete from Material_Request.mr where released is null and mr_name = @mr_name  and job = @job and MBOM = @MBOM"
                check_cmd21.Connection = Login.Connection
                check_cmd21.ExecuteNonQuery()


                Dim check_cmd22 As New MySqlCommand
                check_cmd22.Parameters.AddWithValue("@mr_name", mr_name)
                check_cmd22.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
                check_cmd22.Connection = Login.Connection
                check_cmd22.ExecuteNonQuery()
                '==============================

                temp_panel_name = ""
                temp_panel_desc = ""
                temp_panel_qty = 1

                Panel_grid.Rows.Clear()
                Call Combine_bom()
                Call Save_p_bom(job_Selected, True)
                Call populate_panel_bom()

                '--reset values ------
                Panel_n_l.Text = "Name:   Unknown"
                qty_l.Text = "Qty: 1"


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub ReleaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReleaseToolStripMenuItem.Click

        If String.Equals("New BOM", Me.Text) = False And String.Equals(job_Selected, "") = False And String.Equals(shipping_ad, "") = False Then

            Dim result As DialogResult = MessageBox.Show("Are you sure you want to release this BOM Package?  Releasing a BOM will notify Procurement and Inventory to start allocating parts for the job. Please, make sure the BOM package is correct before releasing it", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                wait_la.Visible = True
                Application.DoEvents()

                '-- assign id_bom to all boms related to MBOM

                Dim my_bom = New List(Of String)() 'all BOMs except MB
                Dim id_bom As Integer : id_bom = 2

                Try
                    '--------- get all BOMs related except MB itself
                    Dim cmd2 As New MySqlCommand
                    cmd2.Parameters.AddWithValue("@job", job_Selected)
                    cmd2.Parameters.AddWithValue("@MBOM", Me.Text)
                    cmd2.CommandText = "SELECT distinct mr_name from Material_Request.mr where released is null and MBOM = @MBOM and job = @job"
                    cmd2.Connection = Login.Connection
                    Dim reader2 As MySqlDataReader
                    reader2 = cmd2.ExecuteReader

                    If reader2.HasRows Then
                        While reader2.Read
                            my_bom.Add(reader2(0))
                        End While
                    End If

                    reader2.Close()

                    '----------------------------------
                    '----------------------------------------------------------------
                    'Update MBOM    
                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.Clear()
                    Create_cmd.Parameters.AddWithValue("@mr_name", Me.Text)
                    Create_cmd.Parameters.AddWithValue("@user", current_user)

                    Create_cmd.CommandText = "UPDATE Material_Request.mr SET released = 'Y', released_by = @user, release_date = now(), id_bom = 1 where mr_name = @mr_name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Dim Create_cmd2 As New MySqlCommand
                    Create_cmd2.Parameters.AddWithValue("@mr_name", Me.Text)

                    Create_cmd2.CommandText = "UPDATE Material_Request.mr_data SET released = 'Y' where mr_name = @mr_name"
                    Create_cmd2.Connection = Login.Connection
                    Create_cmd2.ExecuteNonQuery()

                    '----------------------------------

                    For i = 0 To my_bom.Count - 1

                        Dim Create_cmd1 As New MySqlCommand
                        Create_cmd1.Parameters.Clear()
                        Create_cmd1.Parameters.AddWithValue("@mr_name", my_bom(i).ToString)
                        Create_cmd1.Parameters.AddWithValue("@user", current_user)
                        Create_cmd1.Parameters.AddWithValue("@id_bom", id_bom)

                        Create_cmd1.CommandText = "UPDATE Material_Request.mr SET released = 'Y', released_by = @user, release_date = now(), id_bom = @id_bom where mr_name = @mr_name"
                        Create_cmd1.Connection = Login.Connection
                        Create_cmd1.ExecuteNonQuery()

                        Dim Create_cmd21 As New MySqlCommand
                        Create_cmd21.Parameters.AddWithValue("@mr_name", my_bom(i).ToString)

                        Create_cmd21.CommandText = "UPDATE Material_Request.mr_data SET released = 'Y' where mr_name = @mr_name"
                        Create_cmd21.Connection = Login.Connection
                        Create_cmd21.ExecuteNonQuery()

                        id_bom = id_bom + 1
                    Next

                    Call Create_build_request(job_Selected)  'Create Build request

                    Call Create_MPL(job_Selected)  'create mpl

                    '////////////////   send APL notification to procurement and inventory and mfg
                    If enable_mess = True Then


                        'write mail
                        Dim mail_n As String : mail_n = "Material Request and MPL for Project " & job_Selected & "  has been Released to Procurement" & vbCrLf & vbCrLf _
                 & "Released by: " & current_user & vbCrLf _
                 & "Released Date: " & Now & vbCrLf _
                 & "Project: " & job_Selected & vbCrLf _
                 & "BOM File Name: " & Me.Text & vbCrLf



                        Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request and MPL has been Released for Project " & job_Selected, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request and MPL has been Released for Project " & job_Selected, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request and MPL has been Released for Project " & job_Selected, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Manufacturing", "Material Request and MPL has been Released for Project " & job_Selected, mail_n)
                        Call Sent_mail.Sent_multiple_emails("General Management", "Material Request and MPL has been Released for Project " & job_Selected, mail_n)

                        'add email addresses
                        Dim emails_addr As New List(Of String)()

                        emails_addr.Add("TBullard@atronixengineering.com")
                        emails_addr.Add("dshipman@atronixengineering.com")

                        ''procurement
                        emails_addr.Add("ecoy@atronixengineering.com")
                        emails_addr.Add("fvargas@atronixengineering.com")
                        emails_addr.Add("mmorris@atronixengineering.com")
                        emails_addr.Add("sowens@atronixengineering.com")

                        ''mfg
                        emails_addr.Add("shenley@atronixengineering.com")
                        emails_addr.Add("mowens@atronixengineering.com")

                        ''inventory
                        emails_addr.Add("dnix@atronixengineering.com")
                        emails_addr.Add("dmoore@atronixengineering.com")
                        emails_addr.Add("bbullard@atronixengineering.com")


                        '  For i = 0 To emails_addr.Count - 1
                        Try

                            Dim e_mail As New MailMessage()
                            e_mail = New MailMessage()
                            e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                            For i = 0 To emails_addr.Count - 1
                                e_mail.To.Add(emails_addr.Item(i))
                            Next

                            e_mail.Subject = "Material Request for Project " & job_Selected & "  has been Released"
                            e_mail.IsBodyHtml = False
                            e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

                            Smtp_Server.Send(e_mail)

                        Catch error_t As Exception
                            MsgBox(error_t.ToString)
                        End Try


                    End If

                    '-------- confirmation message -------------'
                    Call send_confirmation_m(current_user, Me.Text)
                    '------------------------------------------


                    '----------------------------------
                    Call Inventory_manage.General_inv_cal()   'recalculate inventory values

                    wait_la.Visible = False
                    MessageBox.Show(Me.Text & " was released successfully!")

                    '--clear everything to avoid having a released BOM package open in this form
                    Me.Text = "New BOM"
                    Panel_n_l.Text = "Name:   Unknown"
                    qty_l.Text = "Qty: 1"
                    ComboBox1.Items.Clear()
                    job_label.Text = "Current Job:  "
                    job_Selected = ""
                    temp_panel_name = ""
                    temp_panel_desc = ""
                    temp_panel_qty = 1

                    need_by_date = DateTimePicker1.Value
                    shipping_ad = ""

                    Panel_grid.Rows.Clear()
                    field_grid.Rows.Clear()
                    assem_grid.Rows.Clear()
                    sp_grid.Rows.Clear()
                    m_grid.Rows.Clear()


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            End If

        Else
            MessageBox.Show("Please Open a BOM package and assign a job number")
        End If
    End Sub

    Sub Create_build_request(job As String)

        '--- build request = all panels bom and assembly boms

        Dim build_name As String : build_name = job & "_Build_Request"

        Dim unmerge_table = New DataTable
        unmerge_table.Columns.Add("Part_No", GetType(String))
        unmerge_table.Columns.Add("Description", GetType(String))
        unmerge_table.Columns.Add("qty", GetType(String))
        unmerge_table.Columns.Add("need_date", GetType(String))
        unmerge_table.Columns.Add("ready", GetType(String))

        Try
            '---------- get date ------
            Dim need_date As String : need_date = ""
            Dim check_cmdn As New MySqlCommand
            check_cmdn.Parameters.AddWithValue("@job", job)
            check_cmdn.CommandText = "select need_date from Material_Request.mr where job = @job and BOM_type = 'MB'"

            check_cmdn.Connection = Login.Connection
            check_cmdn.ExecuteNonQuery()

            Dim readern As MySqlDataReader
            readern = check_cmdn.ExecuteReader

            If readern.HasRows Then
                While readern.Read
                    need_date = readern(0)
                End While
            End If

            readern.Close()

            '--------------------------

            Dim all_list = New List(Of String)()
            Dim all_assem = New List(Of String)()  'contains assem mr

            '--------------------------------------------------

            Dim n_bom As Double : n_bom = 0
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", job)
            check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    n_bom = reader(0)
                End While
            End If

            reader.Close()
            '---------------------------------------------------

            '------------------ enter panels mr -------------
            For i = 2 To n_bom

                Dim check_cmd1 As New MySqlCommand
                check_cmd1.Parameters.Clear()
                check_cmd1.Parameters.AddWithValue("@job", job)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom  and BOM_type = 'Panel' order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    While reader1.Read
                        all_list.Add(reader1(0).ToString)
                    End While
                End If

                reader1.Close()
            Next


            '--------------- enter assem mr ----------------------
            For i = 2 To n_bom

                Dim check_cmd1 As New MySqlCommand
                check_cmd1.Parameters.Clear()
                check_cmd1.Parameters.AddWithValue("@job", job)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom  and BOM_Type = 'Assembly'  order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    While reader1.Read
                        all_assem.Add(reader1(0).ToString)
                    End While
                End If

                reader1.Close()
            Next


            '------enter panel data from mr----
            For i = 0 To all_list.Count - 1

                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", all_list.Item(i).ToString)
                cmd41.CommandText = "SELECT Panel_name, Panel_desc, Panel_qty from Material_Request.mr where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        unmerge_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, need_date)
                    End While
                End If

                reader41.Close()
            Next

            '---------enter assembly data from mr_data -----------
            For i = 0 To all_assem.Count - 1

                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", all_assem.Item(i).ToString)
                cmd41.CommandText = "SELECT Part_No, Description, Qty from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        unmerge_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, need_date, "")
                    End While
                End If

                reader41.Close()
            Next


            '------- check how many revisions so far

            Dim n_r As Double : n_r = 0
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@job", job)
            check_cmd2.CommandText = "select count(distinct n_r) from Build_request.build_r where job = @job"

            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    n_r = reader2(0)
                End While
            End If

            reader2.Close()

            If n_r > 0 Then
                build_name = build_name & "_rev_" & n_r
            End If

            '----------------------------------------------

            '--------- inherit ready box values from first build_r ------------

            For i = 0 To unmerge_table.Rows.Count - 1

                Dim cmd416 As New MySqlCommand
                cmd416.Parameters.Clear()
                cmd416.Parameters.AddWithValue("@panel", unmerge_table.Rows(i).Item(0))
                cmd416.Parameters.AddWithValue("@n_r", n_r - 1)
                cmd416.Parameters.AddWithValue("@job", job)
                cmd416.CommandText = "SELECT ready_t from Build_request.build_r where job = @job and n_r = @n_r and panel = @panel"
                cmd416.Connection = Login.Connection
                Dim reader416 As MySqlDataReader
                reader416 = cmd416.ExecuteReader

                If reader416.HasRows Then
                    While reader416.Read
                        unmerge_table.Rows(i).Item(4) = If(reader416.IsDBNull(0) = True, "", reader416(0).ToString)
                    End While
                End If

                reader416.Close()
            Next


            '--------------------------------------------
            For i = 0 To unmerge_table.Rows.Count - 1

                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.Clear()
                main_cmd.Parameters.AddWithValue("@job", job)
                main_cmd.Parameters.AddWithValue("@br_name", build_name)
                main_cmd.Parameters.AddWithValue("@panel", unmerge_table.Rows(i).Item(0).ToString)
                main_cmd.Parameters.AddWithValue("@panel_desc", unmerge_table.Rows(i).Item(1).ToString)
                main_cmd.Parameters.AddWithValue("@panel_qty", unmerge_table.Rows(i).Item(2).ToString)
                main_cmd.Parameters.AddWithValue("@ready_t", unmerge_table.Rows(i).Item(4).ToString)
                main_cmd.Parameters.AddWithValue("@need_date", need_date)
                main_cmd.Parameters.AddWithValue("@n_r", n_r)


                main_cmd.CommandText = "INSERT INTO Build_request.build_r(job, br_name, panel, panel_desc, panel_qty, need_date, n_r, ready_t) VALUES (@job, @br_name, @panel, @panel_desc, @panel_qty, @need_date, @n_r, @ready_t)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Create_MPL(job As String)
        '------- this function create a MPL --------------
        Dim mpl_name As String : mpl_name = job & "_Master_Packing_List"

        Dim unmerge_table = New DataTable
        unmerge_table.Columns.Add("Part_No", GetType(String))
        unmerge_table.Columns.Add("Description", GetType(String))
        unmerge_table.Columns.Add("qty", GetType(String))
        unmerge_table.Columns.Add("need_date", GetType(String))

        Try
            '---------- get date ------
            Dim need_date As String : need_date = ""
            Dim check_cmdn As New MySqlCommand
            check_cmdn.Parameters.AddWithValue("@job", job)
            check_cmdn.CommandText = "select need_date from Material_Request.mr where job = @job and BOM_type = 'MB'"

            check_cmdn.Connection = Login.Connection
            check_cmdn.ExecuteNonQuery()

            Dim readern As MySqlDataReader
            readern = check_cmdn.ExecuteReader

            If readern.HasRows Then
                While readern.Read
                    need_date = readern(0)
                End While
            End If

            readern.Close()

            '--------------------------
            '---- get latest build_request of job specified
            Dim cmd44 As New MySqlCommand
            Dim n_r1 As Integer : n_r1 = 0
            cmd44.Parameters.AddWithValue("@job", job)
            cmd44.CommandText = "SELECT distinct n_r from Build_request.build_r where job = @job order by n_r desc limit 1"
            cmd44.Connection = Login.Connection
            Dim reader44 As MySqlDataReader
            reader44 = cmd44.ExecuteReader

            If reader44.HasRows Then
                While reader44.Read
                    n_r1 = reader44(0).ToString
                End While
            End If

            reader44.Close()
            '----------------------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", job)
            cmd4.Parameters.AddWithValue("@n_r", n_r1)
            cmd4.CommandText = "SELECT panel, panel_desc, panel_qty, need_date from Build_request.build_r where job = @job and n_r = @n_r"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    unmerge_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, need_date)
                End While

            End If

            reader4.Close()


            '--------------- enter field, spare ----------------------
            Dim all_list = New List(Of String)()

            all_list.Add(job & "_Field")
            all_list.Add(job & "_Spare_Parts")


            For i = 0 To all_list.Count - 1

                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", all_list.Item(i).ToString)
                cmd41.CommandText = "SELECT Part_No, Description, Qty from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        unmerge_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, need_date)
                    End While
                End If

                reader41.Close()
            Next


            '------- check how many revisions so far

            Dim n_r As Double : n_r = 0
            Dim check_cmd2 As New MySqlCommand
            check_cmd2.Parameters.AddWithValue("@job", job)
            check_cmd2.CommandText = "select count(distinct n_r) from Master_Packing_List.packing_l where job = @job"

            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    n_r = reader2(0)
                End While
            End If

            reader2.Close()

            If n_r > 0 Then
                mpl_name = mpl_name & "_rev_" & n_r
            End If

            '--------------------------------------------
            For i = 0 To unmerge_table.Rows.Count - 1

                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.Clear()
                main_cmd.Parameters.AddWithValue("@job", job)
                main_cmd.Parameters.AddWithValue("@mpl_name", mpl_name)
                main_cmd.Parameters.AddWithValue("@part_name", unmerge_table.Rows(i).Item(0).ToString)
                main_cmd.Parameters.AddWithValue("@part_desc", unmerge_table.Rows(i).Item(1).ToString)
                main_cmd.Parameters.AddWithValue("@qty", unmerge_table.Rows(i).Item(2).ToString)
                main_cmd.Parameters.AddWithValue("@need_date", need_date)
                main_cmd.Parameters.AddWithValue("@n_r", n_r)


                main_cmd.CommandText = "INSERT INTO Master_Packing_List.packing_l(job, mpl_name, part_name, part_desc, qty, need_date, n_r) VALUES (@job, @mpl_name, @part_name, @part_desc, @qty, @need_date, @n_r)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub send_confirmation_m(username As String, BOM_name As String)

        Try

            Dim send_e As Boolean : send_e = False
            Dim Smtp_Server As New SmtpClient


            'host properties
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            '---- get user email --------'
            Dim cmd41 As New MySqlCommand
            Dim email_user As String : email_user = "notfound"
            cmd41.Parameters.AddWithValue("@user", username)
            cmd41.CommandText = "SELECT email from users where username = @user"
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

            If String.Equals(email_user, "notfound") = False Then

                Dim mail_n As String : mail_n = "=====  THE FOLLOWING IS A APL CONFIRMATION MESSAGE FOR PROJECT : " & job_label.Text & " ======" & vbCrLf & vbCrLf & vbCrLf

                mail_n = mail_n & vbCrLf & "Material Request and MPL for Project " & job_Selected & "  has been Released to Procurement" & vbCrLf & vbCrLf _
                 & "Released by: " & username & vbCrLf _
                 & "Released Date: " & Now & vbCrLf _
                 & "BOM File Name: " & BOM_name & vbCrLf & vbCrLf

                'Try
                '    Dim cmd4 As New MySqlCommand
                '    cmd4.Parameters.AddWithValue("@mr_name", BOM_name)
                '    cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Qty from Material_Request.mr_data where mr_name = @mr_name"
                '    cmd4.Connection = Login.Connection
                '    Dim reader4 As MySqlDataReader
                '    reader4 = cmd4.ExecuteReader

                '    If reader4.HasRows Then
                '        While reader4.Read
                '            mail_n = mail_n & reader4(0).ToString & "    Description: " & reader4(1) & "   Manufacturer: " & reader4(2) & "  Vendor: " & reader4(3) & "  Qty: " & reader4(4) & vbCrLf & vbCrLf
                '        End While
                '    End If

                '    reader4.Close()

                'Catch ex As Exception
                '    MessageBox.Show(ex.ToString)
                'End Try

                '---------------------------------
                Try
                    Dim e_mail As New MailMessage()
                    e_mail = New MailMessage()
                    e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                    e_mail.To.Add(email_user)
                    e_mail.Subject = "MATERIAL REQUEST RELEASE CONFIRMATION FOR PROJECT:  " & job_label.Text
                    e_mail.IsBodyHtml = False
                    e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                    Smtp_Server.Send(e_mail)

                    Catch error_t As Exception
                        MsgBox(error_t.ToString)
                    End Try

                End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub FeatureCodesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FeatureCodesToolStripMenuItem.Click

        feature_s = False
        Feature_sel.ShowDialog()

    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub ChangePanelDescriptionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangePanelDescriptionToolStripMenuItem.Click

        If String.Equals(Me.Text, "New BOM") = False Then

            Description_panels.Text = Me.Text
            Description_panels.ShowDialog()

        End If
    End Sub
End Class