Imports System.Net.Mail
Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop


Public Class Reports

    Public Smtp_Server As New SmtpClient
    Public Message_email As String

    Private Sub Reports_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"
        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try

        Try
            Dim cmd_list As New MySqlCommand
            cmd_list.CommandText = "SELECT distinct Employee_Name from management.assignments"

            cmd_list.Connection = Login.Connection
            Dim reader_l As MySqlDataReader
            reader_l = cmd_list.ExecuteReader

            If reader_l.HasRows Then
                While reader_l.Read
                    ListBox1.Items.Add(reader_l(0))
                End While
            End If

            reader_l.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub


    Sub Write_message(Job As String, start_date As String, end_date As String)


        Try
            Dim rec_parts As New MySqlCommand
            rec_parts.CommandText = "SELECT PO, Company, Part_Description, QTY_Received, Date_Received from jobs.job_" & Job & " where Date_Received <= '" & end_date & "' and Date_Received >= '" & start_date & "' and QTY_Received > 0"

            rec_parts.Connection = Login.Connection
            Dim reader_r As MySqlDataReader
            reader_r = rec_parts.ExecuteReader

            If reader_r.HasRows Then

                Message_email = Message_email & "RECEIVED PARTS REPORT" & vbCrLf & vbCrLf & "JOB: " & Job & vbCrLf & vbCrLf

                While reader_r.Read
                    Message_email = Message_email & " PO: " & reader_r(0) & "  Company:  " & reader_r(1) & "   Part Description: " & Replace(reader_r(2), vbTab, " ") & "    QTY Received: " & reader_r(3) & "   Date: " & reader_r(4) & vbCrLf & vbCrLf
                End While

                Message_email = Message_email & vbCrLf & "===================================================================================" & vbCrLf & vbCrLf


            End If

            reader_r.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim project_managers As New List(Of String)()
        Dim jobs_of_employee As New List(Of String)()
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to send the Received Parts Report to all this recipients", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message


        If (result = DialogResult.Yes) Then

            Call Add_POs() 'Make sure the latest POs in the Received POs folder are in the database

            Try
                Dim assi_list As New MySqlCommand
                assi_list.CommandText = "SELECT distinct Employee_Name from management.assignments"
                assi_list.Connection = Login.Connection

                Dim reader_l As MySqlDataReader
                reader_l = assi_list.ExecuteReader

                If reader_l.HasRows Then
                    While reader_l.Read
                        project_managers.Add(reader_l(0))
                    End While
                End If

                reader_l.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try




            For i = 0 To project_managers.Count - 1

                Message_email = ""  'initialize message
                jobs_of_employee.Clear()

                Try
                    '-------------------------- Collect jobs -----------------------------
                    Dim job_assi As New MySqlCommand
                    job_assi.CommandText = "SELECT distinct Job_number from management.assignments where Employee_Name =  '" & project_managers.Item(i) & "'"
                    job_assi.Connection = Login.Connection

                    Dim reader_m As MySqlDataReader
                    reader_m = job_assi.ExecuteReader

                    If reader_m.HasRows Then
                        While reader_m.Read
                            jobs_of_employee.Add(reader_m(0))
                        End While
                    End If

                    reader_m.Close()

                    '-------------------------  Generate email --------------------------

                    For z = 0 To jobs_of_employee.Count - 1

                        Call Write_message(jobs_of_employee(z), DateTime.Today, DateTime.Today)  'write job message  DateTime.Today     

                    Next



                    '-------------------------- get the email address ----------------------------
                    Dim email_query As New MySqlCommand
                    Dim email_address As String : email_address = ""
                    email_query.CommandText = "SELECT Email_address from management.employees where Employee_name = '" & project_managers.Item(i) & "'"
                    email_query.Connection = Login.Connection

                    Dim reader_em As MySqlDataReader
                    reader_em = email_query.ExecuteReader

                    If reader_em.HasRows Then
                        While reader_em.Read
                            email_address = reader_em(0)
                        End While
                    End If

                    reader_em.Close()



                    '-------- email sending code ---------
                    If String.Equals(email_address, "") = False And String.Equals(Message_email, "") = False Then

                        Try

                            Dim e_mail As New MailMessage()
                            e_mail = New MailMessage()
                            e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                            e_mail.To.Add(email_address)
                            e_mail.Subject = "RECEIVED PARTS REPORT"
                            e_mail.IsBodyHtml = False
                            e_mail.Body = Message_email & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                            Smtp_Server.Send(e_mail)

                        Catch error_t As Exception
                            MsgBox(error_t.ToString)
                        End Try

                    End If
                    '-----------------------------------

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Next

            If project_managers.Count > 0 Then
                MsgBox("Reports Sent")
            End If

        End If

    End Sub

    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected

        'select start date and display it in textbox
        start_date.Text = MonthCalendar1.SelectionRange.Start.ToShortDateString()

    End Sub

    Private Sub MonthCalendar2_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar2.DateSelected

        'select end date and display it in textbox
        end_date.Text = MonthCalendar2.SelectionRange.Start.ToShortDateString()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'add employee to temporary assignment list

        Dim temp_S As Integer
        temp_S = ListBox1.SelectedItems.Count
        For i = 0 To ListBox1.SelectedItems.Count - 1
            If ListBox1.SelectedItems.Count > 0 And Is_on_list(ListBox1.SelectedItem.ToString, ListBox2) = False Then

                ListBox2.Items.Add(ListBox1.SelectedItem)

            End If
        Next i

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'remove employee from listbox 2
        Dim temp_S As Integer
        temp_S = ListBox2.SelectedItems.Count
        For i = 0 To ListBox2.SelectedItems.Count - 1
            If ListBox2.SelectedItems.Count > 0 Then

                ListBox2.Items.Remove(ListBox2.SelectedItem)
            End If
        Next i

    End Sub

    Function Is_on_list(element As String, myList As ListBox) As Boolean

        Is_on_list = False

        For Each person In myList.Items

            If (String.Equals(person.ToString, element) = True) Then

                Is_on_list = True

            End If
        Next

        Return Is_on_list

    End Function

    'extract Employees from listbox 2

    Function MyListofEmployees() As List(Of String)

        MyListofEmployees = New List(Of String)


        For Each person In ListBox2.Items
            MyListofEmployees.Add(person.ToString)
        Next

        Return MyListofEmployees

    End Function



    '----------------- change date to yyyy-mm-dd format -------------

    Function myDate_format(date_t As String) As String

        myDate_format = ""
        Dim new_date() As String

        new_date = date_t.Split("/")
        myDate_format = new_date(2) & "-" & new_date(0) & "-" & new_date(1)
        Return myDate_format


    End Function




    '------- This section is a copy of the methods in Enter PO and auto launcher with slightly variations. It pretty much enter all the POs received today
    Sub Add_POs()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim path As String : path = "O:\atlanta\Users\Steve Henley\Parts Tracking\Received PO's"

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            ' Choose the folder where the PO are located


            For Each file As String In My.Computer.FileSystem.GetFiles(path)

                Dim strLastModified As String
                strLastModified = System.IO.File.GetLastWriteTime(path & "\" & IO.Path.GetFileName(file)).ToShortDateString()

                If InStr(IO.Path.GetFileName(file), ".xls") > 0 Then

                    If DateTime.Compare(strLastModified, DateTime.Now.ToShortDateString) >= 0 Then
                        Dim wb As Excel.Workbook = xlApp.Workbooks.Open(path & "\" & IO.Path.GetFileName(file))
                        Call PO_parse(wb, IO.Path.GetFileName(file), strLastModified)

                    End If
                End If
            Next


        End If



    End Sub

    Sub PO_parse(wb As Excel.Workbook, file_po As String, file_date As String)

        Try

            Dim Job As String : Job = "Unknown"              'job number 117182
            Dim job_table As String : job_table = "job_"
            Dim PO_n As String : PO_n = "Unknown"            'po number 117182:03-P001
            Dim PO_date As Date : PO_date = file_date
            Dim Company As String : Company = "Unknown"      'Company ex:MCM


            If String.Equals("Project ID", wb.Worksheets(1).Cells(1, 1).value) = True And String.Equals("Description", wb.Worksheets(1).Cells(1, 3).value) = True Then

                ' ----------- Parsing po file name -----------------

                Dim is_PO_new As Boolean : is_PO_new = False  'use for revision
                    Dim compare_t As Boolean : compare_t = False 'use for revision loop

                    Dim entire_file As String() = file_po.Split(New Char() {"."c})
                    Dim po_name_company As String() = entire_file(0).Split(New Char() {" "c})

                    If po_name_company.Length > 1 Then
                        Company = po_name_company(1)  'store company name
                    Else
                        Company = "Unknown"
                    End If



                    Dim po_name_only As String : po_name_only = po_name_company(0)  'store po name as in file name  ex: 117182-03-P002
                    PO_n = Strings.Replace(po_name_only, "-", ":", 1, 1)
                    Dim dIndex = PO_n.IndexOf("R")

                    If (dIndex > -1) Then  'If there is an R, get rid of it

                        PO_n = PO_n.Substring(0, dIndex)

                    End If



                    Dim components As String() = po_name_only.Split(New Char() {"-"c})
                    Job = components(0)
                    job_table = "job_" & Job  'job_xxxxxx, use for mysql insertion

                    '----calculate number of parts in PO

                    Dim numRow As Integer : numRow = 2

                    While (wb.ActiveSheet.Cells(numRow, 3).Value IsNot Nothing)
                        numRow = numRow + 1
                    End While

                '-------------------------------------------------
                '---- Verify if job exist otherwise create job table
                '-----CREATE TABLE job_xxxxx(PO varchar(80), Company varchar(80), Part_Description varchar(90), QTY_Ordered int, QTY_Received int, Date_Received DATE);

                Try
                    Dim cmd As New MySqlCommand
                    cmd.CommandText = "SELECT * from jobs." & job_table
                    cmd.Connection = Login.Connection
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader
                    reader.Close()

                Catch ex As Exception

                    'If table does not exist then create it
                    Try
                        Dim cmd2 As New MySqlCommand
                        cmd2.CommandText = "CREATE TABLE jobs." & job_table & "(PO varchar(80), Company varchar(80), Part_Description varchar(220), QTY_Ordered int, QTY_Received int, Date_Received DATE)"
                        cmd2.Connection = Login.Connection
                        cmd2.ExecuteNonQuery()
                        is_PO_new = True

                    Catch ex2 As Exception
                        MessageBox.Show(ex2.ToString)
                    End Try

                End Try

                '------------------------ Check for revisions -----------------------------

                If (is_PO_new = False) Then

                        Dim counter As Integer : counter = 1

                        Do While compare_t = False

                        Try

                            Dim rev_cmd As New MySqlCommand
                            rev_cmd.Parameters.AddWithValue("@PO", PO_n) 'PO                
                            rev_cmd.CommandText = "SELECT distinct Part_Description from jobs." & job_table & " where PO = @PO order by Part_Description"
                            rev_cmd.Connection = Login.Connection
                            Dim reader_r As MySqlDataReader
                            reader_r = rev_cmd.ExecuteReader
                            Dim partList As New List(Of String)()
                            Dim excelList_t As New List(Of String)()

                            'fill partList with the PO parts
                            If reader_r.HasRows Then

                                While reader_r.Read

                                    partList.Add(reader_r(0).ToString)

                                End While

                                'fill excelList with excel parts
                                For j = 2 To numRow - 1
                                    excelList_t.Add(wb.Worksheets(1).Cells(j, 3).value)
                                Next

                                Dim excelList As List(Of String) = excelList_t.Distinct().ToList  'remove all duplicates

                                partList.Sort()
                                excelList.Sort()  'ascending order

                                compare_t = Compare_PO(partList, excelList)

                                If (compare_t = False) Then

                                    If counter = 1 Then
                                        PO_n = PO_n & "R" & counter
                                    Else
                                        Dim dIndex2 = PO_n.IndexOf("R")
                                        PO_n = PO_n.Substring(0, dIndex2)
                                        PO_n = PO_n & "R" & counter
                                    End If


                                End If

                                counter = counter + 1 'increase Revision counter ex 117181:03-P002Ri where i = 1....?
                            Else
                                compare_t = True  'end the loop

                            End If

                            reader_r.Close()

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                            End Try
                        Loop

                    End If

                '-----------------------------------------------------


                For i = 2 To numRow - 1
                    Call Insert_data(PO_n, Company, wb.Worksheets(1).Cells(i, 3).value, wb.Worksheets(1).Cells(i, 4).value, wb.Worksheets(1).Cells(i, 5).value, job_table, PO_date)
                Next

            End If


        Catch ex As Exception
        End Try


        wb.Close(False)


    End Sub

    Function Compare_PO(Parts_po As List(Of String), Parts_excel As List(Of String)) As Boolean
        'This function compare two list of parts. If the list is different then it will return false else it will return true

        Compare_PO = True

        Dim Total_Po As Integer : Total_Po = Parts_po.Count - 1
        Dim Total_Excel As Integer : Total_Excel = Parts_excel.Count - 1

        If Total_Po = Total_Excel Then

            For i As Integer = 0 To Total_Po
                If String.Equals(Parts_po(i), Parts_excel(i)) = False Then
                    Compare_PO = False 'discrepancy found, break loop
                    Exit For
                End If
            Next


        Else
            Compare_PO = False
        End If

    End Function


    '-------- This function is in charge to enter the PO to the right table and update qty received values --------------

    Sub Insert_data(PO As String, Company As String, description As String, qty_or As Integer, on_hand As Integer, job As String, date_po As Date)

        Dim qty_received As Integer : qty_received = 0

        'convert mm/dd/yy to yy-mm-dd
        If String.Equals(date_po, "") = False Then
            Dim left_o() As String
            Dim new_date() As String
            new_date = date_po.ToString.Split("/")
            left_o = new_date(2).Split(" ")
            date_po = left_o(0) & "-" & new_date(0) & "-" & new_date(1)
        End If


        'get rid of quotes in PLC components

        description = description.Replace("""", "")


        '--get the amount received so far and substract it to the on hand qty  (ON hand - qty received so far = Qty received)
        Try

            Dim enter_record_zero As Boolean : enter_record_zero = True
            Dim Create_cmd2 As New MySqlCommand
            Create_cmd2.Parameters.AddWithValue("@PO", PO) 'PO
            Create_cmd2.Parameters.AddWithValue("@Company", Company) 'Company
            Create_cmd2.Parameters.AddWithValue("@description", description) 'description
            Create_cmd2.Parameters.AddWithValue("@qty_or", qty_or) 'qty_ord
            Create_cmd2.Parameters.AddWithValue("@date_po", date_po)  'date

            Create_cmd2.CommandText = "Select SUM(QTY_Received) from jobs." & job & " where PO = @PO and Company = @Company and Part_Description = @description and QTY_Ordered = @qty_or and  Date_Received <= @date_po "
            Create_cmd2.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = Create_cmd2.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    If IsDBNull(reader(0)) = False Then
                        qty_received = on_hand - CType(reader(0), Integer)  'calculate received qty
                    Else
                        qty_received = on_hand   'very first qty received
                    End If
                End While
            End If

            reader.Close()

            If qty_received = 0 Then

                'make sure there is no more than one PO entry with qty_received of 0. We only need one to display the entire PO but more than one is a waste of space
                Dim Create_cmd3 As New MySqlCommand
                Create_cmd3.Parameters.AddWithValue("@PO", PO) 'PO
                Create_cmd3.Parameters.AddWithValue("@Company", Company) 'Company
                Create_cmd3.Parameters.AddWithValue("@description", description) 'description
                Create_cmd3.Parameters.AddWithValue("@qty_or", qty_or) 'qty_ord

                Create_cmd3.CommandText = "Select * from jobs." & job & " where PO = @PO and Company = @Company and Part_Description = @description and QTY_Ordered = @qty_or and QTY_Received = 0"
                Create_cmd3.Connection = Login.Connection
                Dim reader_zero As MySqlDataReader
                reader_zero = Create_cmd3.ExecuteReader


                If reader_zero.HasRows Then
                    enter_record_zero = False
                End If

                reader_zero.Close()

            End If

            If enter_record_zero = True Then

                '---------- insert record ---------------------------
                Dim insert_cmd As New MySqlCommand
                insert_cmd.Parameters.AddWithValue("@PO", PO)
                insert_cmd.Parameters.AddWithValue("@Company", Company)
                insert_cmd.Parameters.AddWithValue("@description", description)
                insert_cmd.Parameters.AddWithValue("@qty_or", qty_or.ToString)
                insert_cmd.Parameters.AddWithValue("@qty_received", qty_received.ToString)
                insert_cmd.Parameters.AddWithValue("@Date", date_po)

                insert_cmd.CommandText = "INSERT INTO jobs." & job & " VALUES (@PO, @Company, @description,  @qty_or, @qty_received, @Date)"
                insert_cmd.Connection = Login.Connection
                insert_cmd.ExecuteNonQuery()


                '------------ update record adjacent to the one inserted if exist ----------------
                Dim date_up As String : date_up = ""
                Dim qty_up As Integer : qty_up = -1

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@PO", PO) 'PO
                cmd4.Parameters.AddWithValue("@Company", Company) 'Company
                cmd4.Parameters.AddWithValue("@description", description) 'description
                cmd4.Parameters.AddWithValue("@qty_or", qty_or) 'qty_ord
                cmd4.Parameters.AddWithValue("@date_po", date_po)  'date

                cmd4.CommandText = "SELECT QTY_Received, Date_Received from jobs." & job & " where PO = @PO and Company = @Company and Part_Description = @description and QTY_Ordered = @qty_or and Date_Received > @date_po order by Date_Received limit 1"
                cmd4.Connection = Login.Connection
                Dim reader2 As MySqlDataReader
                reader2 = cmd4.ExecuteReader

                If reader2.HasRows Then
                    While reader2.Read
                        qty_up = reader2(0) - qty_received
                        date_up = reader2(1)
                    End While
                End If

                reader2.Close()

                'if found rows, update record
                If qty_up >= 0 And String.Equals("", date_up) = False Then

                    Dim left_o() As String
                    Dim new_date() As String
                    new_date = date_up.ToString.Split("/")
                    left_o = new_date(2).Split(" ")
                    date_up = left_o(0) & "-" & new_date(0) & "-" & new_date(1)

                    Dim upd_cmd As New MySqlCommand
                    upd_cmd.Parameters.AddWithValue("@PO", PO) 'PO
                    upd_cmd.Parameters.AddWithValue("@Company", Company) 'Company
                    upd_cmd.Parameters.AddWithValue("@description", description) 'description
                    upd_cmd.Parameters.AddWithValue("@qty_or", qty_or) 'qty_ord
                    upd_cmd.Parameters.AddWithValue("@date_up", date_up)  'date
                    upd_cmd.Parameters.AddWithValue("@qty_rec", qty_up)  'update received date

                    upd_cmd.CommandText = "UPDATE jobs." & job & " SET QTY_Received = @qty_rec where PO = @PO and Company = @Company and Part_Description = @description and QTY_Ordered = @qty_or and  Date_Received =  @date_up"
                    upd_cmd.Connection = Login.Connection
                    upd_cmd.ExecuteNonQuery()

                End If

            End If

        Catch ex2 As Exception
            MessageBox.Show(ex2.ToString)
        End Try


    End Sub


    '----------------------------- CUSTOM REPORT SECTION ----------------------------
    '-----------------------------------------------------------------------------------------------------------------

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click


        Dim project_managers As List(Of String) = New List(Of String)(MyListofEmployees())
        Dim jobs_of_employee As New List(Of String)()
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to send the Custom Received Parts Report to all this recipients", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            If (String.Equals(start_date.Text, "") = False And String.Equals(end_date.Text, "") = False) Then 'make sure dates are non empty

                Call Add_POs() 'Make sure the latest POs in the Received POs folder are in the database

                For i = 0 To project_managers.Count - 1

                    Message_email = ""  'initialize message
                    jobs_of_employee.Clear()

                    Try
                        '-------------------------- Collect jobs -----------------------------
                        Dim job_assi As New MySqlCommand
                        job_assi.CommandText = "SELECT distinct Job_number from management.assignments where Employee_Name =  '" & project_managers.Item(i) & "'"
                        job_assi.Connection = Login.Connection

                        Dim reader_m As MySqlDataReader
                        reader_m = job_assi.ExecuteReader

                        If reader_m.HasRows Then
                            While reader_m.Read
                                jobs_of_employee.Add(reader_m(0))
                            End While
                        End If

                        reader_m.Close()

                        '-------------------------  Generate email --------------------------

                        For z = 0 To jobs_of_employee.Count - 1

                            Call Write_message(jobs_of_employee(z), myDate_format(start_date.Text), myDate_format(end_date.Text))  'write job message  DateTime.Today     

                        Next



                        '-------------------------- get the email address ----------------------------
                        Dim email_query As New MySqlCommand
                        Dim email_address As String : email_address = ""
                        email_query.CommandText = "SELECT Email_address from management.employees where Employee_name = '" & project_managers.Item(i) & "'"
                        email_query.Connection = Login.Connection

                        Dim reader_em As MySqlDataReader
                        reader_em = email_query.ExecuteReader

                        If reader_em.HasRows Then
                            While reader_em.Read
                                email_address = reader_em(0)
                            End While
                        End If

                        reader_em.Close()



                        '-------- email sending code ---------
                        If String.Equals(email_address, "") = False And String.Equals(Message_email, "") = False Then

                            Try

                                Dim e_mail As New MailMessage()
                                e_mail = New MailMessage()
                                e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                                e_mail.To.Add(email_address)
                                e_mail.Subject = "RECEIVED PARTS REPORT"
                                e_mail.IsBodyHtml = False
                                e_mail.Body = Message_email & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                                Smtp_Server.Send(e_mail)

                            Catch error_t As Exception
                                MsgBox(error_t.ToString)
                            End Try

                        End If
                        '-----------------------------------

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                Next

                If project_managers.Count > 0 Then
                    MsgBox("Custom Reports Sent")
                End If

            Else
                    MessageBox.Show("Please select date range")

            End If



        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '======================== Generate Report ===========================

        Dim project_managers As List(Of String) = New List(Of String)(MyListofEmployees())
        Dim jobs_of_employee As New List(Of String)()

        If (String.Equals(start_date.Text, "") = False And String.Equals(end_date.Text, "") = False) Then 'make sure dates are non empty

            Call Add_POs() 'Make sure the latest POs in the Received POs folder are in the database

            For i = 0 To project_managers.Count - 1

                jobs_of_employee.Clear()
                FolderBrowserDialog1.Description = "Select " & project_managers.Item(i) & " report folder"

                Try
                    '-------------------------- Collect jobs -----------------------------
                    Dim job_assi As New MySqlCommand
                    job_assi.CommandText = "SELECT distinct Job_number from management.assignments where Employee_Name =  '" & project_managers.Item(i) & "'"
                    job_assi.Connection = Login.Connection

                    Dim reader_m As MySqlDataReader
                    reader_m = job_assi.ExecuteReader

                    If reader_m.HasRows Then
                        While reader_m.Read
                            jobs_of_employee.Add(reader_m(0))
                        End While
                    End If

                    reader_m.Close()

                    '-------------------------  Generate email --------------------------

                    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                        For z = 0 To jobs_of_employee.Count - 1
                            Call Generate_report(jobs_of_employee(z), myDate_format(start_date.Text), myDate_format(end_date.Text), FolderBrowserDialog1.SelectedPath, project_managers.Item(i))  'write job message  DateTime.Today     
                        Next

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Next

            If project_managers.Count > 0 Then
                MsgBox("Reports created successfully")
            End If

        Else
            MessageBox.Show("Please select date range")

        End If


    End Sub

    '============== This function generate an excel report of all the parts received in the date range specified ================================
    Sub Generate_report(Job As String, start_date As String, end_date As String, path As String, worker As String)

        If String.Equals(path, "") = False Then

            Try
                Dim rec_parts As New MySqlCommand
                rec_parts.CommandText = "SELECT PO, Company, Part_Description, QTY_Received, Date_Received from jobs.job_" & Job & " where Date_Received <= '" & end_date & "' and Date_Received >= '" & start_date & "' and QTY_Received > 0"

                rec_parts.Connection = Login.Connection
                Dim reader_r As MySqlDataReader
                reader_r = rec_parts.ExecuteReader

                If reader_r.HasRows Then

                    Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

                    If xlApp Is Nothing Then
                        MessageBox.Show("Excel is not properly installed!!")
                    Else

                        Dim i As Integer : i = 4
                        Dim xlWorkBook As Excel.Workbook
                        Dim xlWorkSheet As Excel.Worksheet
                        Dim misValue As Object = System.Reflection.Missing.Value
                        xlWorkBook = xlApp.Workbooks.Add(misValue)
                        xlWorkSheet = xlWorkBook.Sheets("sheet1")

                        '=================  Format of the spreadsheet ======================
                        xlWorkSheet.Cells(1, 1).Font.Bold = True

                        xlWorkSheet.Cells(2, 1).Font.Bold = True
                        xlWorkSheet.Cells(2, 2).Font.Bold = True
                        xlWorkSheet.Cells(2, 3).Font.Bold = True
                        xlWorkSheet.Cells(2, 4).Font.Bold = True
                        xlWorkSheet.Cells(2, 5).Font.Bold = True

                        xlWorkSheet.Cells(1, 1) = Job
                        xlWorkSheet.Range("A:A").ColumnWidth = 33
                        xlWorkSheet.Range("B:B").ColumnWidth = 43
                        xlWorkSheet.Range("C:C").ColumnWidth = 43
                        xlWorkSheet.Range("D:D").ColumnWidth = 23
                        xlWorkSheet.Range("E:E").ColumnWidth = 23

                        xlWorkSheet.Range("A:A").HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Range("B:B").HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Range("C:C").HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Range("D:D").HorizontalAlignment = Excel.Constants.xlCenter
                        xlWorkSheet.Range("E:E").HorizontalAlignment = Excel.Constants.xlCenter

                        xlWorkSheet.Cells(2, 1) = "PO (Purchase Order)"
                        xlWorkSheet.Cells(2, 2) = "COMPANY"
                        xlWorkSheet.Cells(2, 3) = "PART DESCRIPTION"
                        xlWorkSheet.Cells(2, 4) = "QTY RECEIVED"
                        xlWorkSheet.Cells(2, 5) = "DATE RECEIVED"

                        While reader_r.Read

                            xlWorkSheet.Cells(i, 1) = reader_r(0)
                            xlWorkSheet.Cells(i, 2) = reader_r(1)
                            xlWorkSheet.Cells(i, 3) = reader_r(2)
                            xlWorkSheet.Cells(i, 4) = reader_r(3)
                            xlWorkSheet.Cells(i, 5) = reader_r(4)

                            i = i + 1

                        End While

                        xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\" & worker & "_" & Job & ".xlsx")
                        xlWorkBook.Close(True)
                        ' xlApp.Quit()

                        'end excel app in system
                        releaseObject(xlWorkSheet)
                        releaseObject(xlWorkBook)
                        releaseObject(xlApp)

                    End If

                End If

                reader_r.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try


        End If

    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch ex As Exception

        Finally
            obj = Nothing
        End Try
    End Sub

End Class