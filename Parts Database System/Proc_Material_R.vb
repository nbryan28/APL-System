Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class Proc_Material_R

    Public open_job As String
    Public toggle As Boolean
    Public start_e As Boolean


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ViewMateriaLTrackingReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewMateriaLTrackingReportToolStripMenuItem.Click
        Repor_track.Visible = True
    End Sub

    Private Sub CreateMRRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateMRRevisionToolStripMenuItem.Click
        Report_tracking_mn.Visible = True
    End Sub

    Private Sub OpenMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenMaterialRequestToolStripMenuItem.Click
        My_Material_r.mode = "pc_rel"
        Open_file.ShowDialog()
    End Sub

    Private Sub RelaseMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs)

        'If String.Equals(Me.Text, "My Material Requests") = False Then

        '    Try
        '        '----------- check if its already released -----------

        '        Dim exist_c As Boolean = False


        '        Dim cmd4 As New MySqlCommand
        '        cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
        '        cmd4.CommandText = "SELECT job from Material_Request.mr where mr_name = @mr_name and released_inv = 'Y'"
        '        cmd4.Connection = Login.Connection
        '        Dim reader4 As MySqlDataReader
        '        reader4 = cmd4.ExecuteReader

        '        If reader4.HasRows Then
        '            exist_c = True

        '        End If

        '        reader4.Close()

        '        If exist_c = False Then

        '            Dim result As DialogResult = MessageBox.Show("Are you sure you want to release this Material Request to the Inventory department?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        '            If (result = DialogResult.Yes) Then

        '                Dim Create_cmd As New MySqlCommand
        '                Create_cmd.Parameters.AddWithValue("@mr_name", Me.Text)

        '                Create_cmd.CommandText = "UPDATE Material_Request.mr SET released_inv = 'Y' where mr_name = @mr_name"
        '                Create_cmd.Connection = Login.Connection
        '                Create_cmd.ExecuteNonQuery()

        '                If enable_mess = True Then

        '                    '--- SEND NOTIICATION
        '                    Dim email_m As String
        '                    email_m = "Material Request for Project " & open_job & "  has been sent to Inventory" & vbCrLf & vbCrLf _
        '                & "Sent by: " & current_user & vbCrLf _
        '                & "Sent Date: " & Now & vbCrLf _
        '                & "Project: " & open_job & vbCrLf _
        '                & "Material Request File Name: " & Me.Text & vbCrLf

        '                    Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request for Project " & open_job & " Ready", email_m)
        '                End If


        '                MessageBox.Show(Me.Text & " was released to inventory successfully!")
        '            End If

        '        Else
        '            MessageBox.Show("Material Request already released to inventory!")
        '        End If

        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)
        '    End Try

        'Else
        '    MessageBox.Show("Please open a Project")
        'End If
    End Sub

    Private Sub Proc_Material_R_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        open_job = ""
        toggle = False
        start_e = False

        total_grid.Columns(total_grid.Columns.Count - 1).Visible = False

        Timer1.Interval = 13000
        Timer1.Start()

    End Sub

    Private Sub total_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles total_grid.CellValueChanged

        For Each row As DataGridViewRow In total_grid.Rows
            If row.IsNewRow Then Continue For

            If (IsNumeric(row.Cells(0).Value) = True And IsNumeric(row.Cells(8).Value)) Then

                row.Cells(9).Value = CType(row.Cells(0).Value, Double) - CType(row.Cells(8).Value, Double)

                If CType(row.Cells(9).Value, Double) > 0 Then
                    row.Cells(9).Style.BackColor = Color.Firebrick
                End If

            End If

            '--------

            If start_e = True Then

                Try
                    If String.IsNullOrEmpty(row.Cells(10).Value) = False Then

                        Dim Create_cmd As New MySqlCommand
                        Dim mr_n As String : mr_n = get_last_revision(Me.Text)


                        Create_cmd.Parameters.AddWithValue("@Part_No", row.Cells(1).Value)
                        Create_cmd.Parameters.AddWithValue("@mr_name", mr_n)
                        Create_cmd.Parameters.AddWithValue("@status_m", row.Cells(10).Value)

                        Create_cmd.CommandText = "UPDATE Material_Request.mr_data  SET status_m = @status_m where Part_No = @Part_No and mr_name = @mr_name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If

        Next



    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If String.Equals("My Material Requests", Me.Text) = False And Me.Visible = True And toggle = True Then

            start_e = False
            total_grid.Rows.Clear()

            Dim mr_name As String : mr_name = Me.Text

            Try
                'Dim cmd As New MySqlCommand
                'cmd.Parameters.AddWithValue("@job", open_job)
                'cmd.CommandText = "Select mr_name from Material_Request.mr where released = 'Y' and job = @job order by release_date desc limit 1"
                'cmd.Connection = Login.Connection
                'Dim reader As MySqlDataReader
                'reader = cmd.ExecuteReader

                'If reader.HasRows Then
                '    While reader.Read
                '        mr_name = reader(0).ToString
                '    End While
                'End If

                'reader.Close()



                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", mr_name)
                cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, mfg_type, ADA_Number, Vendor, Price, subtotal, Qty, qty_fullfilled, qty_needed, status_m, Assembly_name, Part_status, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name and (full_panel is null or full_panel <> 'Yes')"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    Dim i As Integer : i = 0
                    While reader4.Read
                        total_grid.Rows.Add(New String() {})
                        total_grid.Rows(i).Cells(0).Value = reader4(8).ToString
                        total_grid.Rows(i).Cells(1).Value = reader4(0).ToString
                        total_grid.Rows(i).Cells(2).Value = reader4(1).ToString
                        total_grid.Rows(i).Cells(3).Value = reader4(2).ToString
                        total_grid.Rows(i).Cells(4).Value = reader4(3).ToString
                        total_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                        total_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                        total_grid.Rows(i).Cells(7).Value = If(IsNumeric(reader4(8)) = True, reader4(8).ToString, 0) * If(IsNumeric(reader4(6)) = True, reader4(6).ToString, 0)
                        total_grid.Rows(i).Cells(8).Value = reader4(9).ToString
                        total_grid.Rows(i).Cells(9).Value = reader4(10).ToString
                        total_grid.Rows(i).Cells(10).Value = reader4(11).ToString
                        total_grid.Rows(i).Cells(11).Value = reader4(15).ToString
                        total_grid.Rows(i).Cells(12).Value = reader4(14).ToString
                        total_grid.Rows(i).Cells(13).Value = reader4(12).ToString
                        total_grid.Rows(i).Cells(14).Value = reader4(13).ToString
                        total_grid.Rows(i).Cells(15).Value = reader4(4).ToString

                        i = i + 1
                    End While

                End If

                reader4.Close()
                Me.Text = mr_name
                start_e = True

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        Dim mr_name As String = "xxxy6"

        Dim date_created As String : date_created = ""
        Dim created_by As String : created_by = ""
        Dim released_by As String : released_by = ""
        Dim release_date As String : release_date = ""
        Dim last_modified As String : last_modified = ""

        date_c_l.Text = "Date Created:"
        created_by_l.Text = "Created By:"
        release_by_l.Text = "Released By:"
        release_d_l.Text = "Released Date:"
        last_m_l.Text = "Last Modified:"

        If Not ComboBox2.SelectedItem Is Nothing Then
            mr_name = ComboBox2.SelectedItem.ToString
        Else
            mr_name = "xxxy6"
        End If

        Try

            PR_grid.Rows.Clear()
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name, Part_Status, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name and (full_panel is null or full_panel <> 'Yes')"
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
                    PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                    PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                    PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                    PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                    PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                    PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                    PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                    PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                    PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString

                    i = i + 1
                End While

            End If

            reader4.Close()

            '-------------------- dates ----------

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT Date_Created, created_by, released_by, release_date, last_modified from Material_Request.mr where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                Dim i As Integer : i = 0
                While reader41.Read

                    date_c_l.Text = "Date Created: " & reader41(0).ToString
                    created_by_l.Text = "Created By: " & reader41(1).ToString
                    release_by_l.Text = "Released By: " & reader41(2).ToString
                    release_d_l.Text = "Released Date: " & reader41(3).ToString
                    last_modified = "Last Modified:" & reader41(4).ToString

                End While

            End If

            reader41.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        '-- refresn on/off
        If String.Equals(Label6.Text, "Refresh off") = True Then

            Label6.Text = "Refresh on"
            toggle = True
        Else
            Label6.Text = "Refresh off"
            toggle = False
        End If
    End Sub

    Private Sub ProjectTotalsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProjectTotalsToolStripMenuItem.Click

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            '  If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

            Try

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
                For i As Integer = 0 To total_grid.ColumnCount - 1
                    xlWorkSheet.Cells(1, i + 1) = total_grid.Columns(i).HeaderText
                    For j As Integer = 0 To total_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = total_grid.Rows(j).Cells(i).Value
                    Next j
                Next i


                'xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Material_Request_" & open_job & ".xlsx")
                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Order exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            '  End If
        End If
    End Sub

    Private Sub MaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaterialRequestToolStripMenuItem.Click
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            'If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

            Try

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

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"
                '  xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Material_Request_" & ComboBox2.Text & ".xlsx")
                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                'xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Order exported successfully!")
                ' End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If total_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(total_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In total_grid.Rows
            If String.IsNullOrEmpty(row.Cells.Item("DataGridViewTextBoxColumn3").Value) = False Then
                If String.Compare(row.Cells.Item("DataGridViewTextBoxColumn3").Value.ToString, TextBox1.Text, True) = 0 Then
                    rowindex = row.Index
                    total_grid.CurrentCell = total_grid.Rows(rowindex).Cells(0)
                    found_po = True
                    Exit For
                End If

            End If
        Next

        If found_po = False Then
            MsgBox("Part not found!")
        End If
    End Sub

    Private Sub SendMaterialRequestPartSReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SendMaterialRequestPartSReportToolStripMenuItem.Click


        If String.Equals("My Material Requests", Me.Text) = False Then

            Dim send_e As Boolean : send_e = False
            Dim Smtp_Server As New SmtpClient


            'host properties
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"


            Dim emails_addr As New List(Of String)()

            emails_addr.Add("TBullard@atronixengineering.com")
            'emails_addr.Add("dshipman@atronixengineering.com")
            'emails_addr.Add("bdahlqvist@atronixengineering.com")
            emails_addr.Add("shenley@atronixengineering.com")

            Try
                '--get who released the MR----
                Dim released_by As String : released_by = ""

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd4.Parameters.AddWithValue("@job", job_label.Text)
                cmd4.CommandText = "SELECT released_by from Material_Request.mr where mr_name = @mr_name and job = @job"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        released_by = reader4(0).ToString
                    End While
                End If

                reader4.Close()

                '-- get user email ----



                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@user", released_by)
                cmd41.CommandText = "SELECT email from users where username = @user"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then

                    While reader41.Read
                        If reader41(0).ToString.Contains("@") Then
                            emails_addr.Add(reader41(0).ToString)
                        Else
                            MessageBox.Show(released_by & " has no email on record. Please add an email to this user")
                        End If
                    End While

                End If

                reader41.Close()

                '----------------------

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try



            Dim mail_n As String : mail_n = "=====  THE FOLLOWING IS A PARTS STATUS REPORT FOR PROJECT : " & job_label.Text & " ======" & vbCrLf & vbCrLf & vbCrLf

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd4.CommandText = "SELECT Part_No, Status_m from Material_Request.mr_data where mr_name = @mr_name and status_m is not null and status_m <> ''"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        mail_n = mail_n & reader4(0).ToString & "    STATUS: " & reader4(1) & vbCrLf & vbCrLf
                        send_e = True
                    End While
                End If

                reader4.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            If send_e = True Then

                For i = 0 To emails_addr.Count - 1
                    Try

                        Dim e_mail As New MailMessage()
                        e_mail = New MailMessage()
                        e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                        e_mail.To.Add(emails_addr.Item(i))
                        e_mail.Subject = "MATERIAL REQUEST PARTS STATUS REPORT FOR PROJECT:  " & job_label.Text
                        e_mail.IsBodyHtml = False
                        e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                        Smtp_Server.Send(e_mail)

                    Catch error_t As Exception
                        MsgBox(error_t.ToString)
                    End Try

                Next

            End If

            MessageBox.Show("Report sent succesfully!")
        Else
            MessageBox.Show("Please open a project")
        End If
    End Sub

    Function get_last_revision(mr_name As String) As String

        get_last_revision = "xskfskdnfmslkdk"
        Dim id_bom As String : id_bom = "9999999999"
        Dim job As String : job = ""

        Try
            '---------- get id_bom --------------
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT id_bom, job from Material_Request.mr where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    id_bom = reader4(0).ToString
                    job = reader4(1).ToString
                End While
            End If

            reader4.Close()

            '----------------------------------
            '-- get mr name ---
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@id_bom", id_bom)
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    get_last_revision = reader41(0).ToString
                End While
            End If

            reader41.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- compare select mr
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
                cmd4.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                cmd4.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
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
                cmd41.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                cmd41.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
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
                    cmd41.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                    cmd41.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd41.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            compare_grid.Rows(i).Cells(2).Value = reader41(0).ToString
                            compare_grid.Rows(i).Cells(4).Value = reader41(1).ToString
                        End While
                    End If

                    reader41.Close()

                    Dim cmd42 As New MySqlCommand
                    cmd42.Parameters.Clear()
                    cmd42.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                    cmd42.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd42.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd42.Connection = Login.Connection
                    Dim reader42 As MySqlDataReader
                    reader42 = cmd42.ExecuteReader

                    If reader42.HasRows Then
                        While reader42.Read
                            compare_grid.Rows(i).Cells(3).Value = reader42(0).ToString
                            compare_grid.Rows(i).Cells(5).Value = reader42(1).ToString
                        End While
                    End If

                    reader42.Close()

                    compare_grid.Rows(i).Cells(6).Value = If(IsNumeric(compare_grid.Rows(i).Cells(5).Value), compare_grid.Rows(i).Cells(5).Value, 0) - If(IsNumeric(compare_grid.Rows(i).Cells(4).Value), compare_grid.Rows(i).Cells(4).Value, 0)

                    If compare_grid.Rows(i).Cells(6).Value <> 0 Then
                        compare_grid.Rows(i).Cells(6).Style.BackColor = Color.CadetBlue
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Next



        End If
    End Sub

    Private Sub ComparisonTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComparisonTableToolStripMenuItem.Click
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            '  If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

            Try

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
                For i As Integer = 0 To compare_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = compare_grid.Columns(i).HeaderText
                    For j As Integer = 0 To compare_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = compare_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                'xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\" & mrbox1.Text & "_and_" & mrbox2.Text & ".xlsx")
                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                '  xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Order exported successfully!")
                ' End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub AssemblyBOMOverviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssemblyBOMOverviewToolStripMenuItem.Click
        If String.Equals(Me.Text, "My Material Requests") = False Then
            ASM_overview.Visible = True
        Else
            MessageBox.Show("Please Open a BOM")
        End If
    End Sub

    Private Sub AcknowledgeBOMSenderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcknowledgeBOMSenderToolStripMenuItem.Click
        '--- This menu send a APL message to the person that released the BOM
        If String.Equals(Me.Text, "My Material Requests") = False Then
            Try
                '--- get username of person that release the BOM 
                Dim user_n As String : user_n = ""
                Dim role_m As String : role_m = ""
                Dim role_s As String : role_s = "General Management"

                Dim cmd_v2 As New MySqlCommand

                cmd_v2.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd_v2.CommandText = "SELECT released_by from Material_Request.mr where mr_name = @mr_name"
                cmd_v2.Connection = Login.Connection
                Dim readerv2 As MySqlDataReader
                readerv2 = cmd_v2.ExecuteReader

                If readerv2.HasRows Then
                    While readerv2.Read
                        user_n = readerv2(0).ToString
                    End While
                End If

                readerv2.Close()
                '------------------------------------
                '-- get role of user -----
                Dim cmd_v21 As New MySqlCommand

                cmd_v21.Parameters.AddWithValue("@user", user_n)
                cmd_v21.CommandText = "SELECT Role from users where username = @user"
                cmd_v21.Connection = Login.Connection
                Dim readerv21 As MySqlDataReader
                readerv21 = cmd_v21.ExecuteReader

                If readerv21.HasRows Then
                    While readerv21.Read
                        role_m = readerv21(0).ToString
                    End While
                End If

                readerv21.Close()
                '---------------------------

                'send APL message
                Dim Create_cmd As New MySqlCommand

                Dim mail_n As String : mail_n = "BOM has been received successfully. The Procurement team will begin allocating and purchasing the parts you requested" & vbCrLf & vbCrLf _
                         & "Project: " & job_label.Text & vbCrLf _
                         & "BOM File Name: " & Me.Text & vbCrLf

                Create_cmd.Parameters.AddWithValue("@Sender", current_user)
                Create_cmd.Parameters.AddWithValue("@role_s", role_s)
                Create_cmd.Parameters.AddWithValue("@Receiver", user_n)
                Create_cmd.Parameters.AddWithValue("@role_r", role_m)
                Create_cmd.Parameters.AddWithValue("@priority", "n")
                Create_cmd.Parameters.AddWithValue("@date_s", Now)
                Create_cmd.Parameters.AddWithValue("@read_m", "n")
                Create_cmd.Parameters.AddWithValue("@Mail", mail_n)
                Create_cmd.Parameters.AddWithValue("@title", "BOM Confirmation Message")

                Create_cmd.CommandText = "INSERT INTO management.Dropbox(Sender, role_s, Receiver, role_r, priority, date_s, read_m, Mail, title) VALUES (@Sender,@role_s,@Receiver,@role_r,@priority,@date_s,@read_m,@Mail,@title)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()


                MessageBox.Show("Acknowldge Message sent!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Update_prices(mr_name As String)
        '-- update the prices of a mr
        Try
            For i = 0 To total_grid.Rows.Count - 1

                If String.IsNullOrEmpty(total_grid.Rows(i).Cells(6).Value) = False Then
                    If IsNumeric(total_grid.Rows(i).Cells(6).Value) = True Then

                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                        Create_cmd.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(1).Value)
                        Create_cmd.Parameters.AddWithValue("@Price", total_grid.Rows(i).Cells(6).Value)

                        Create_cmd.CommandText = "UPDATE Material_Request.mr_data  SET Price = @Price where Part_No = @Part_No and mr_name = @mr_name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                    End If
                End If

            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub UpdatePricesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdatePricesToolStripMenuItem.Click
        '---- update prices of the current MR and latest mr
        Dim latest_mr As String : latest_mr = "xxxx"
        latest_mr = get_last_revision(Me.Text)

        Call Update_prices(Me.Text)
        Call Update_prices(latest_mr)

        MessageBox.Show("Prices updated successfully")

    End Sub
End Class