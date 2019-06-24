Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail


Public Class Inventory_manage

    Public part_sel As String
    Public start_cal As Boolean



    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub Inventory_manage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        part_sel = ""
        start_cal = False

        EnableDoubleBuffered(fullfill_grid)


        Try
            Dim cmd_dt As New MySqlCommand
            Dim reader_panel2 As MySqlDataReader

            cmd_dt.CommandText = "SELECT * from inventory.inventory_qty"
            cmd_dt.Connection = Login.Connection
            reader_panel2 = cmd_dt.ExecuteReader

            If reader_panel2.HasRows Then
                While reader_panel2.Read
                    fullfill_grid.Rows.Add(New String() {reader_panel2(0).ToString, reader_panel2(1).ToString, reader_panel2(2).ToString, reader_panel2(3).ToString, reader_panel2(4).ToString, reader_panel2(5).ToString, reader_panel2(6).ToString, reader_panel2(7).ToString, reader_panel2(8).ToString, reader_panel2(9).ToString, reader_panel2(10).ToString})
                End While
            End If

            reader_panel2.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        start_cal = True

    End Sub

    Sub General_inv_cal()

        '-------------------------------------
        'Dim Smtp_Server As New SmtpClient

        ''host properties
        'Smtp_Server.UseDefaultCredentials = False
        'Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        'Smtp_Server.Port = 587
        'Smtp_Server.EnableSsl = True
        'Smtp_Server.Host = "smtp.gmail.com"

        '/////////// Main function that loop the inventory table check for qyt below min values calculate parts needed including demand take into consideration upcoming qty and fill the material order table along sending notification and email
        Dim dimen_table As DataTable
        Dim total As Double : total = 0
        Dim demand As Double : demand = 0


        '=================== create datatable ====================
        dimen_table = New DataTable
        dimen_table.Columns.Add("part_name", GetType(String))
        dimen_table.Columns.Add("min_qty", GetType(String))
        dimen_table.Columns.Add("max_qty", GetType(String))
        dimen_table.Columns.Add("current_qty", GetType(String))
        dimen_table.Columns.Add("Qty_in_order", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))

        Try
            'truncate material order

            Dim check_cmd3 As New MySqlCommand
            check_cmd3.CommandText = "truncate table inventory.Material_orders"
            check_cmd3.Connection = Login.Connection
            check_cmd3.ExecuteNonQuery()


            'populate datatable 
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select part_name, min_qty, max_qty, current_qty, Qty_in_order, description from inventory.inventory_qty where current_qty <  min_qty"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    dimen_table.Rows.Add(reader(0).ToString, reader(1).ToString, reader(2).ToString, reader(3).ToString, reader(4).ToString, reader(5).ToString)
                End While
            End If

            reader.Close()

            'calculate values
            For i = 0 To dimen_table.Rows.Count - 1
                'max - current_inv - coming_parts
                total = CType(dimen_table.Rows(i).Item(2), Double) - CType(dimen_table.Rows(i).Item(3), Double) - CType(dimen_table.Rows(i).Item(4), Double)
                demand = Cal_demand(dimen_table.Rows(i).Item(0)) 'function calculate demand of all active projects

                total = total + demand

                'insert into material order table 
                If total > 0 And ((CType(dimen_table.Rows(i).Item(3), Double) + CType(dimen_table.Rows(i).Item(4), Double) - demand) < CType(dimen_table.Rows(i).Item(1), Double) = True) Then   'after second AND is (on hnad, qty_transit - demand) >= min

                    Dim main_cmd As New MySqlCommand
                    main_cmd.Parameters.AddWithValue("@Part_No", dimen_table.Rows(i).Item(0))
                    main_cmd.Parameters.AddWithValue("@Qty_needed", If(total < 0, 0, total))
                    main_cmd.Parameters.AddWithValue("@description", dimen_table.Rows(i).Item(5))
                    main_cmd.CommandText = "INSERT INTO inventory.Material_orders(Part_No, Description, Qty_needed) VALUES (@Part_No, @description, @Qty_needed)"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()

                End If

            Next


            '--------- check if message was sent ---------
            'Dim message_t As String = "Inventory Demand Report " & Date.Today
            'Dim mail_sent As Boolean = True
            'Dim message As String = "An Inventory Report was sent the " & Now & vbCrLf & vbCrLf & "Please check the Inventory Demand window."


            'Dim check_cmd1 As New MySqlCommand
            'check_cmd1.Parameters.AddWithValue("@title", message_t)
            'check_cmd1.Parameters.AddWithValue("@receiver", "Erin")
            'check_cmd1.CommandText = "select Mail from management.Dropbox where Title = @title and Receiver = @receiver"
            'check_cmd1.Connection = Login.Connection
            'check_cmd1.ExecuteNonQuery()

            'Dim reader1 As MySqlDataReader
            'reader1 = check_cmd1.ExecuteReader

            'If reader1.HasRows Then
            '    While reader1.Read
            '        mail_sent = False
            '    End While
            'End If

            'reader1.Close()

            '------ sent notification -----
            ' If enable_mess = True Then

            'If mail_sent = True Then

            'Dim create_cmd As New MySqlCommand
            'create_cmd.Parameters.Clear()
            'create_cmd.Parameters.AddWithValue("@Sender", "APL System")
            'create_cmd.Parameters.AddWithValue("@role_s", "General Management")
            'create_cmd.Parameters.AddWithValue("@Receiver", "Erin")
            'create_cmd.Parameters.AddWithValue("@role_r", "Procurement Management")
            'create_cmd.Parameters.AddWithValue("@priority", "n")
            'create_cmd.Parameters.AddWithValue("@date_s", Now)
            'create_cmd.Parameters.AddWithValue("@read_m", "n")
            'create_cmd.Parameters.AddWithValue("@Mail", message)
            'create_cmd.Parameters.AddWithValue("@title", message_t)

            'create_cmd.CommandText = "INSERT INTO management.Dropbox(Sender, role_s, Receiver, role_r, priority, date_s, read_m, Mail, title) VALUES (@Sender,@role_s,@Receiver,@role_r,@priority,@date_s,@read_m,@Mail,@title)"
            'create_cmd.Connection = Login.Connection
            'create_cmd.ExecuteNonQuery()

            '-------- sent email ----------
            'Dim e_mail As New MailMessage()
            'e_mail = New MailMessage()
            'e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
            'e_mail.To.Add("ecoy@atronixengineering.com")
            'e_mail.Subject = message_t
            'e_mail.IsBodyHtml = False
            'e_mail.Body = message & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
            'Smtp_Server.Send(e_mail)

            '---------- sent email to many people ----------

            ' Dim emails_addr As New List(Of String)()

            'procurement
            ' emails_addr.Add("ecoy@atronixengineering.com")
            ' emails_addr.Add("fvargas@atronixengineering.com")
            ' emails_addr.Add("mmorris@atronixengineering.com")
            'emails_addr.Add("sowens@atronixengineering.com")

            'mfg
            'emails_addr.Add("shenley@atronixengineering.com")
            'emails_addr.Add("mowens@atronixengineering.com")


            '  For i = 0 To emails_addr.Count - 1
            'Try
            '    Dim e_mail2 As New MailMessage()
            '    e_mail2 = New MailMessage()
            '    e_mail2.From = New MailAddress("inventory.atronix.system@gmail.com")

            '    For i = 0 To emails_addr.Count - 1
            '        e_mail2.To.Add(emails_addr.Item(i))
            '    Next

            '    e_mail2.Subject = message_t
            '    e_mail2.IsBodyHtml = False
            '    e_mail2.Body = message & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
            '    Smtp_Server.Send(e_mail2)

            'Catch error_t As Exception
            '    MsgBox(error_t.ToString)
            'End Try
            '  Next

            '---------------------------------------------------
            'send APL email to inventory about assemblies

            'Try
            '    Dim got_assemblies = New List(Of String)()
            '    Dim cmd2 As New MySqlCommand
            '    cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            '    cmd2.Connection = Login.Connection
            '    Dim reader2 As MySqlDataReader
            '    reader2 = cmd2.ExecuteReader

            '    If reader2.HasRows Then
            '        While reader2.Read
            '            got_assemblies.Add(reader2(0))
            '        End While
            '    End If

            '    reader2.Close()

            '    Dim assb_t As Boolean : assb_t = False

            '    For i = 0 To got_assemblies.Count - 1

            '        Dim cmd3 As New MySqlCommand
            '        cmd3.Parameters.AddWithValue("@part_name", got_assemblies.Item(i).ToString)
            '        cmd3.CommandText = "SELECT Part_No, Qty_needed from inventory.Material_orders where Part_No = @part_name"
            '        cmd3.Connection = Login.Connection
            '        Dim reader3 As MySqlDataReader
            '        reader3 = cmd3.ExecuteReader

            '        If reader3.HasRows Then
            '            While reader3.Read
            '                assb_t = True
            '            End While
            '        End If

            '        reader3.Close()
            '    Next


            '    Dim mail_s As String : mail_s = "Please check the Assemblies Demand Menu for more information"

            '    '-- send APL to inventory and MFG
            '    If assb_t = True Then
            '        Call Sent_mail.Sent_multiple_emails("Manufacturing", "Assemblies Demand Update", mail_s)
            '        Call Sent_mail.Sent_multiple_emails("Inventory", "Assemblies Demand Update", mail_s)
            '    End If


            'Catch ex As Exception
            '    MessageBox.Show(ex.ToString)
            'End Try

            '---------------------------------------------------

            'End If
            'End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub


    Function cal_demand(part_name As String) As Double
        '------- This function calculate the demand of all active jobs'

        cal_demand = 0

        Try

            'Dim job_list As New List(Of String)()

            'Dim cmd4 As New MySqlCommand
            'cmd4.CommandText = "SELECT distinct job from Material_Request.mr where job is not null and finished is null"
            'cmd4.Connection = Login.Connection
            'Dim reader4 As MySqlDataReader
            'reader4 = cmd4.ExecuteReader

            'If reader4.HasRows Then
            '    While reader4.Read
            '        job_list.Add(reader4(0).ToString)
            '    End While
            'End If

            'reader4.Close()

            'For i = 0 To job_list.Count - 1

            '    '----- get number of boms of the job specified -----

            '    Dim n_bom As Double : n_bom = 0
            '    Dim check_cmd As New MySqlCommand
            '    check_cmd.Parameters.Clear()
            '    check_cmd.Parameters.AddWithValue("@job", job_list.Item(i).ToString)
            '    check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

            '    check_cmd.Connection = Login.Connection
            '    check_cmd.ExecuteNonQuery()

            '    Dim readerb As MySqlDataReader
            '    readerb = check_cmd.ExecuteReader

            '    If readerb.HasRows Then
            '        While readerb.Read
            '            n_bom = readerb(0)
            '        End While
            '    End If

            '    readerb.Close()

            '    '----------------------------------------------

            '    For j = 1 To n_bom

            '        Dim mr_name As String : mr_name = ""

            '        Dim cmd5 As New MySqlCommand
            '        cmd5.Parameters.Clear()
            '        cmd5.Parameters.AddWithValue("@job", job_list.Item(i).ToString)
            '        cmd5.Parameters.AddWithValue("@id_bom", j)
            '        cmd5.CommandText = "SELECT mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"
            '        cmd5.Connection = Login.Connection
            '        Dim reader5 As MySqlDataReader
            '        reader5 = cmd5.ExecuteReader

            '        If reader5.HasRows Then
            '            While reader5.Read
            '                mr_name = reader5(0).ToString
            '            End While
            '        End If

            '        reader5.Close()

            '        '------- find demand of part -----

            '        Dim cmd As New MySqlCommand
            '        cmd.Parameters.Clear()
            '        cmd.Parameters.AddWithValue("@mr_name", mr_name)
            '        cmd.Parameters.AddWithValue("@Part_No", part_name)
            '        cmd.CommandText = "SELECT Qty, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name and Part_No = @Part_No"
            '        cmd.Connection = Login.Connection
            '        Dim reader As MySqlDataReader
            '        reader = cmd.ExecuteReader

            '        If reader.HasRows Then
            '            While reader.Read
            '                cal_demand = cal_demand + If(IsNumeric(reader(0).ToString) = True, reader(0).ToString, 0) - If(IsNumeric(reader(1).ToString) = True, reader(1).ToString, 0)
            '            End While
            '        End If

            '        reader.Close()

            '    Next

            'Next

            '--- new procedure--
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_No", part_name)
            cmd.CommandText = "select sum(Qty), sum(qty_fullfilled) from Material_Request.mr_data where Part_No = @Part_No and latest_r = 'x' and released = 'Y'"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    cal_demand = If(IsDBNull(reader(0)), 0, reader(0)) - If(IsDBNull(reader(1)), 0, reader(1))
                End While
            End If

            reader.Close()
            '-------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function




    Private Sub EditPartInventoryQtysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditPartInventoryQtysToolStripMenuItem.Click

        If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

            Dim component As String : component = ""
            Dim index As Integer : index = 0
            component = fullfill_grid.CurrentCell.Value.ToString
            index = fullfill_grid.CurrentCell.RowIndex

            UR_inventory.part_name.Text = component
            UR_inventory.Location_t.Text = fullfill_grid.Rows(index).Cells(7).Value
            UR_inventory.Min_qty.Text = fullfill_grid.Rows(index).Cells(5).Value
            UR_inventory.max_qty.Text = fullfill_grid.Rows(index).Cells(6).Value
            UR_inventory.current_qty.Text = fullfill_grid.Rows(index).Cells(8).Value

            UR_inventory.ShowDialog()

        End If

    End Sub

    'Private Sub AddPartToInventoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddPartToInventoryToolStripMenuItem.Click

    '    'part_sel = "active"
    '    'Part_selector.ShowDialog()

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
                xlWorkSheet.Range("D:L").ColumnWidth = 20
                xlWorkSheet.Range("A:L").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To fullfill_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = fullfill_grid.Columns(i).HeaderText
                    For j As Integer = 0 To fullfill_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = fullfill_grid.Rows(j).Cells(i).Value
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

    'Private Sub RemovePartFromInventoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemovePartFromInventoryToolStripMenuItem.Click
    '    '-- Remove a part from the inventory table

    '    If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

    '        Dim component As String : component = ""
    '        component = fullfill_grid.CurrentCell.Value.ToString

    '        Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this part from Inventory? Please make sure the current inventory of this part is zero before proceed", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

    '        If (result = DialogResult.Yes) Then

    '            Try

    '                Dim check_cmd2 As New MySqlCommand
    '                check_cmd2.Parameters.AddWithValue("@part_name", component)
    '                check_cmd2.CommandText = "delete from inventory.inventory_qty where part_name = @part_name"
    '                check_cmd2.Connection = Login.Connection
    '                check_cmd2.ExecuteNonQuery()

    '                '  EnableDoubleBuffered(fullfill_grid)

    '                '---reload-------

    '                Call refresh_table()

    '                MessageBox.Show("Part has been removed from Inventory table successfully!")

    '            Catch ex As Exception
    '                MessageBox.Show(ex.ToString)
    '            End Try

    '        End If
    '    End If
    'End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        Call refresh_table()
    End Sub

    Sub refresh_table()
        '--refresh
        EnableDoubleBuffered(fullfill_grid)

        Try
            Dim cmd_dt As New MySqlCommand
            Dim reader_panel2 As MySqlDataReader
            fullfill_grid.Rows.Clear()

            cmd_dt.CommandText = "SELECT * from inventory.inventory_qty"
            cmd_dt.Connection = Login.Connection
            reader_panel2 = cmd_dt.ExecuteReader

            If reader_panel2.HasRows Then
                While reader_panel2.Read
                    fullfill_grid.Rows.Add(New String() {reader_panel2(0).ToString, reader_panel2(1).ToString, reader_panel2(2).ToString, reader_panel2(3).ToString, reader_panel2(4).ToString, reader_panel2(5).ToString, reader_panel2(6).ToString, reader_panel2(7).ToString, reader_panel2(8).ToString, reader_panel2(9).ToString, reader_panel2(10).ToString})
                End While
            End If

            reader_panel2.Close()



            Call General_inv_cal()
            MessageBox.Show("Table refreshed!")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


    'Private Sub fullfill_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles fullfill_grid.CellValueChanged

    '    'If start_cal = True Then

    '    '    Call General_inv_cal()
    '    'End If
    'End Sub

    Private Sub Inventory_manage_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged

        If Me.Visible = True Then

            EnableDoubleBuffered(fullfill_grid)

            Try
                Dim cmd_dt As New MySqlCommand
                Dim reader_panel2 As MySqlDataReader
                fullfill_grid.Rows.Clear()

                cmd_dt.CommandText = "SELECT * from inventory.inventory_qty"
                cmd_dt.Connection = Login.Connection
                reader_panel2 = cmd_dt.ExecuteReader

                If reader_panel2.HasRows Then
                    While reader_panel2.Read
                        fullfill_grid.Rows.Add(New String() {reader_panel2(0).ToString, reader_panel2(1).ToString, reader_panel2(2).ToString, reader_panel2(3).ToString, reader_panel2(4).ToString, reader_panel2(5).ToString, reader_panel2(6).ToString, reader_panel2(7).ToString, reader_panel2(8).ToString, reader_panel2(9).ToString, reader_panel2(10).ToString})
                    End While
                End If

                reader_panel2.Close()


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub AddSpecificPartToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        '-------- add specific part --------
        '  Spec_inv.ShowDialog()

    End Sub

    Sub add_specific(part_name As String, desc As String, mfg As String)

        '------ add specific part --------

        Try
            If String.Equals(part_name, "") = False Then

                '--check if part exist
                Dim check_cmd As New MySqlCommand
                Dim exist_c As Boolean : exist_c = False
                check_cmd.Parameters.AddWithValue("@part_name", part_name)
                check_cmd.CommandText = "select * from inventory.inventory_qty where part_name = @part_name"
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    exist_c = True
                End If

                reader.Close()
                '---------------------

                If exist_c = False Then

                    Dim main_cmd As New MySqlCommand
                    main_cmd.Parameters.AddWithValue("@part_name", part_name)
                    main_cmd.Parameters.AddWithValue("@description", desc)
                    main_cmd.Parameters.AddWithValue("@manufacturer", mfg)
                    main_cmd.Parameters.AddWithValue("@MFG_type", "Specific_part")
                    main_cmd.Parameters.AddWithValue("@units", "")
                    main_cmd.Parameters.AddWithValue("@min_qty", 0)
                    main_cmd.Parameters.AddWithValue("@max_qty", 1)
                    main_cmd.Parameters.AddWithValue("@location", "shelf01")
                    main_cmd.Parameters.AddWithValue("@current_qty", 0)
                    main_cmd.Parameters.AddWithValue("@Qty_in_order", 0)
                    main_cmd.Parameters.AddWithValue("@es_date_of_arrival", DBNull.Value)

                    main_cmd.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, units, min_qty, max_qty, location, current_qty, Qty_in_order, es_date_of_arrival) VALUES (@part_name, @description, @manufacturer, @MFG_type, @units, @min_qty, @max_qty, @location, @current_qty, @Qty_in_order, @es_date_of_arrival)"
                    main_cmd.Connection = Login.Connection
                    main_cmd.ExecuteNonQuery()

                End If
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If fullfill_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(fullfill_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In fullfill_grid.Rows
            If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
                If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text, True) = 0 Then
                    rowindex = row.Index
                    fullfill_grid.CurrentCell = fullfill_grid.Rows(rowindex).Cells(0)
                    found_po = True
                    Exit For
                End If

            End If
        Next

        If found_po = False Then
            MsgBox("Part not found!")
        End If
    End Sub

    Private Sub CalculateDemandToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateDemandToolStripMenuItem.Click

        If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

            Dim component As String : component = ""
            component = fullfill_grid.CurrentCell.Value.ToString
            MessageBox.Show(cal_demand(component))

        End If
    End Sub

    'Private Sub AssembliesMinAndMaxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssembliesMinAndMaxToolStripMenuItem.Click
    '    Update_Min_Max.Visible = True
    'End Sub

    'Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem.Click
    '    '--- once function to add
    '    Try
    '        Dim dimen_table = New DataTable
    '        dimen_table.Columns.Add("Part_No", GetType(String))
    '        dimen_table.Columns.Add("Description", GetType(String))
    '        dimen_table.Columns.Add("Manufacturer", GetType(String))
    '        dimen_table.Columns.Add("mfg_type", GetType(String))
    '        dimen_table.Columns.Add("units", GetType(String))

    '        Dim cmd2 As New MySqlCommand
    '        cmd2.CommandText = "SELECT Part_Name, Part_Description, Manufacturer, MFG_type, Units from parts_table"
    '        cmd2.Connection = Login.Connection
    '        Dim reader2 As MySqlDataReader
    '        reader2 = cmd2.ExecuteReader

    '        If reader2.HasRows Then
    '            While reader2.Read
    '                dimen_table.Rows.Add(reader2(0).ToString, reader2(1).ToString, reader2(2).ToString, reader2(3).ToString, reader2(4).ToString)
    '            End While
    '        End If

    '        reader2.Close()

    '        For i = 0 To dimen_table.Rows.Count - 1

    '            Dim exist_c As Boolean : exist_c = False

    '            Dim cmd21 As New MySqlCommand
    '            cmd21.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0).ToString)
    '            cmd21.CommandText = "SELECT * from inventory.inventory_qty where part_name = @part_name"
    '            cmd21.Connection = Login.Connection
    '            Dim reader21 As MySqlDataReader
    '            reader21 = cmd21.ExecuteReader

    '            If reader21.HasRows Then
    '                exist_c = True
    '            End If

    '            reader21.Close()
    '            '--------------------------

    '            If exist_c = False Then

    '                Dim Create_cmd As New MySqlCommand
    '                Create_cmd.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0))
    '                Create_cmd.Parameters.AddWithValue("@description", dimen_table.Rows(i).Item(1))
    '                Create_cmd.Parameters.AddWithValue("@manufacturer", dimen_table.Rows(i).Item(2))
    '                Create_cmd.Parameters.AddWithValue("@MFG_type", dimen_table.Rows(i).Item(3))
    '                Create_cmd.Parameters.AddWithValue("@units", dimen_table.Rows(i).Item(4))

    '                Create_cmd.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, units, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, @units,0,0,0,0)"
    '                Create_cmd.Connection = Login.Connection
    '                Create_cmd.ExecuteNonQuery()

    '            End If

    '        Next

    '        MessageBox.Show("done")
    '    Catch ex As Exception
    '        MessageBox.Show(ex.ToString)
    '    End Try
    'End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        '------ search as soon as text change automatically
        EnableDoubleBuffered(fullfill_grid)

        fullfill_grid.Rows.Clear()

        Try
            Dim cmd_dt As New MySqlCommand
            Dim reader_panel2 As MySqlDataReader

            cmd_dt.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")
            cmd_dt.CommandText = "SELECT * from inventory.inventory_qty where part_name LIKE @search"
            cmd_dt.Connection = Login.Connection
            reader_panel2 = cmd_dt.ExecuteReader

            If reader_panel2.HasRows Then
                While reader_panel2.Read
                    fullfill_grid.Rows.Add(New String() {reader_panel2(0).ToString, reader_panel2(1).ToString, reader_panel2(2).ToString, reader_panel2(3).ToString, reader_panel2(4).ToString, reader_panel2(5).ToString, reader_panel2(6).ToString, reader_panel2(7).ToString, reader_panel2(8).ToString, reader_panel2(9).ToString, reader_panel2(10).ToString})
                End While
            End If

            reader_panel2.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub ASSEMMinAndMaxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASSEMMinAndMaxToolStripMenuItem.Click
        Update_Min_Max.ShowDialog()

    End Sub
End Class