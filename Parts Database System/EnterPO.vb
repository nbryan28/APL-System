Imports System.IO
Imports Microsoft.Office.Interop
Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class EnterPO

    Dim process_ls As Integer

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '-------------------- CHECK TYPE OF DATA AND SELECT RIGHT FUNCTIONS -------------------
        Dim type_data As String : type_data = ""

        If Not ComboBox1.SelectedItem Is Nothing Then
            type_data = ComboBox1.SelectedItem.ToString
        Else
            type_data = ""
        End If

        Select Case type_data

            'Enter Purchase Orders data (part, qty_ord, date ,etc)
            Case "Purchase Order"

                Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
                Dim found_ls As Integer : found_ls = 0
                process_ls = 0

                If xlApp Is Nothing Then
                    MessageBox.Show("Excel is not properly installed!!")
                Else
                    ' Choose the folder where the PO are located
                    If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                        found_ls = 0
                        process_ls = 0
                        TextBox1.Clear()
                        TextBox2.Clear()
                        process_l.Text = "File Process: 0"

                        For Each file As String In My.Computer.FileSystem.GetFiles(FolderBrowserDialog1.SelectedPath)

                            Dim strLastModified As String
                            strLastModified = System.IO.File.GetLastWriteTime(FolderBrowserDialog1.SelectedPath & "\" & IO.Path.GetFileName(file)).ToShortDateString()

                            TextBox1.AppendText(IO.Path.GetFileName(file) & vbCrLf)
                            If InStr(IO.Path.GetFileName(file), ".xls") > 0 Then
                                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(FolderBrowserDialog1.SelectedPath & "\" & IO.Path.GetFileName(file))
                                Call PO_parse(wb, IO.Path.GetFileName(file), strLastModified)

                            End If
                            found_ls = found_ls + 1
                        Next
                    End If

                    found_l.Text = "Files Found: " & found_ls
                    MessageBox.Show("Data Insertion Finished Succesfully")

                End If

            Case "Parts Name"
              '  Call Extract_Parts_PO()  'call function to get part names
            Case "Rates"
              '  Call Extract_Rates_PO() 'call function to get rates
            Case "Shipping data"
                MessageBox.Show("Shipping data under construction")
            Case Else
                MessageBox.Show("Please select a type of data")
        End Select
    End Sub

    Sub PO_parse(wb As Excel.Workbook, file_po As String, file_date As String)

        Dim Job As String : Job = "Unknown"              'job number 117182
        Dim job_table As String : job_table = "job_"
        Dim PO_n As String : PO_n = "Unknown"            'po number 117182:03-P001
        Dim PO_date As Date : PO_date = file_date
        Dim Company As String : Company = "Unknown"      'Company ex:MCM



        If String.Equals("Project ID", wb.Worksheets(1).Cells(1, 1).value) = True And String.Equals("Description", wb.Worksheets(1).Cells(1, 3).value) = True And String.Equals("On Hand", wb.Worksheets(1).Cells(1, 5).value) = True Then


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


            process_ls = process_ls + 1
            process_l.Text = "Files Processed " & process_ls  'count processed files
            TextBox2.AppendText(file_po & "--- " & file_date & vbCrLf)

        End If

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

            If enter_record_zero = True And qty_received >= 0 Then  'last conditional avoid negative received numbers

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

    '---------------------<<<<<<<< Enter Parts from PO >>>>>>>>>----------------------------->

    Sub Extract_Parts_PO()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim found_ls As Integer : found_ls = 0
        process_ls = 0

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            ' Choose the folder where the PO are located
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                found_ls = 0
                process_ls = 0
                TextBox1.Clear()
                TextBox2.Clear()
                process_l.Text = "File Process: 0"

                For Each file As String In My.Computer.FileSystem.GetFiles(FolderBrowserDialog1.SelectedPath)

                    TextBox1.AppendText(IO.Path.GetFileName(file) & vbCrLf)
                    If InStr(IO.Path.GetFileName(file), ".xls") > 0 Then
                        Dim wb As Excel.Workbook = xlApp.Workbooks.Open(FolderBrowserDialog1.SelectedPath & "\" & IO.Path.GetFileName(file))
                        Call PO_part_parse(wb, IO.Path.GetFileName(file))  'process po xls file

                    End If
                    found_ls = found_ls + 1
                Next
            End If

            found_l.Text = "Files Found: " & found_ls
            MessageBox.Show("Data Insertion Finished Succesfully")

        End If


    End Sub

    'This function will extract the part names and company from the PO spreadsheet

    Sub PO_part_parse(wb As Excel.Workbook, file_po As String)

        Dim Company As String : Company = ""      'Company ex:MCM

        If String.Equals("Project ID", wb.Worksheets(1).Cells(1, 1).value) = True And String.Equals("Description", wb.Worksheets(1).Cells(1, 3).value) = True Then


            ' ----------- Parsing po file name -----------------
            Dim entire_file As String() = file_po.Split(New Char() {"."c})
            Dim po_name_company As String() = entire_file(0).Split(New Char() {" "c})

            If po_name_company.Length > 1 Then
                Company = po_name_company(1)  'store company name           
            End If


            '----calculate number of parts in PO
            Dim numRow As Integer : numRow = 2

            While (wb.ActiveSheet.Cells(numRow, 3).Value IsNot Nothing)
                numRow = numRow + 1
            End While

            '-------------------------------------------------

            For i = 2 To numRow - 1
                Call Insert_part_data(Company, wb.Worksheets(1).Cells(i, 3).value)
            Next


            process_ls = process_ls + 1
            process_l.Text = "Files Processed " & process_ls  'count processed files
            TextBox2.AppendText(file_po & vbCrLf)

        End If

        wb.Close(False)

    End Sub

    '-------------- enter the parts (TYPE: OTHER) to the part database -------------------
    Sub Insert_part_data(Company As String, part_name As String)


        'get rid of quotes in PLC components, extra spaces and replace tabs for spaces

        part_name = part_name.Replace("""", "")
        part_name = Replace(part_name, vbTab, " ")
        part_name = Regex.Replace(part_name, "\s{2,}", " ")


        '----------- MAKE SURE THE PART DOES NOT EXIST ------------
        Try
            Dim cmd As New MySqlCommand
            Dim Create_cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_Name", part_name)
            cmd.CommandText = "SELECT Part_Name from parts_table where Part_Name = @Part_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            '
            If String.Compare("MCMC", Company, True) = True Or String.Compare("McM", Company, True) = True Or String.Compare("MMC", Company, True) = True Or String.Compare("Mc-Mc", Company) = True Or
                String.Compare("MCM", Company, True) = True Or String.Compare("Macm", Company, True) = True Or String.Compare("MCNMC", Company, True) = True Then

                Company = "McNaughton-McKay"

            End If



            If Not reader.HasRows Then

                'IF IT DOES NOT EXIST THEN CREATE IT
                reader.Close()

                Create_cmd.Parameters.AddWithValue("@Part_Name", part_name)
                Create_cmd.Parameters.AddWithValue("@Manufacturer", "")
                Create_cmd.Parameters.AddWithValue("@Part_Description", "")
                Create_cmd.Parameters.AddWithValue("@Notes", "Part Extracted from PO")
                Create_cmd.Parameters.AddWithValue("@Part_Status", "Preferred")
                Create_cmd.Parameters.AddWithValue("@Part_Type", "Other")
                Create_cmd.Parameters.AddWithValue("@Units", "")
                Create_cmd.Parameters.AddWithValue("@Min_Order_Qty", 1)
                Create_cmd.Parameters.AddWithValue("@Legacy_ADA_Number", "")
                Create_cmd.Parameters.AddWithValue("@Primary_Vendor", Company)


                Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Legacy_ADA_Number, Primary_Vendor) VALUES (@Part_Name, @Manufacturer, @Part_Description, @Notes, @Part_Status, @Part_Type, @Units, @Min_Order_Qty, @Legacy_ADA_Number, @Primary_Vendor)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Else
                'DO NOT FORGET TO CLOSE THE READER
                reader.Close()
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    '---------------------<<<<<<<< Enter Rates from PO >>>>>>>>>----------------------------->

    Sub Extract_Rates_PO()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim found_ls As Integer : found_ls = 0
        process_ls = 0

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            ' Choose the folder where the PO are located
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                found_ls = 0
                process_ls = 0
                TextBox1.Clear()
                TextBox2.Clear()
                process_l.Text = "File Process: 0"

                For Each file As String In My.Computer.FileSystem.GetFiles(FolderBrowserDialog1.SelectedPath)

                    Dim strLastModified As String
                    strLastModified = System.IO.File.GetLastWriteTime(FolderBrowserDialog1.SelectedPath & "\" & IO.Path.GetFileName(file)).ToShortDateString()

                    TextBox1.AppendText(IO.Path.GetFileName(file) & vbCrLf)
                    If InStr(IO.Path.GetFileName(file), ".xls") > 0 Then
                        Dim wb As Excel.Workbook = xlApp.Workbooks.Open(FolderBrowserDialog1.SelectedPath & "\" & IO.Path.GetFileName(file))
                        Call PO_rate_parse(wb, IO.Path.GetFileName(file), strLastModified)

                    End If
                    found_ls = found_ls + 1
                Next
            End If

            found_l.Text = "Files Found: " & found_ls
            MessageBox.Show("Data Insertion Finished Succesfully")

        End If


    End Sub

    'This function will extract the Rates, company, date and part name from the PO spreadsheet

    Sub PO_rate_parse(wb As Excel.Workbook, file_po As String, file_date As String)

        Dim Company As String : Company = ""      'Company ex:MCM


        If String.Equals("Project ID", wb.Worksheets(1).Cells(1, 1).value) = True And String.Equals("Description", wb.Worksheets(1).Cells(1, 3).value) = True And String.Equals("Rate", wb.Worksheets(1).Cells(1, 7).value) = True Then


            ' ----------- Parsing po file name -----------------

            Dim entire_file As String() = file_po.Split(New Char() {"."c})
            Dim po_name_company As String() = entire_file(0).Split(New Char() {" "c})

            If po_name_company.Length > 1 Then
                Company = po_name_company(1)  'store company name           
            End If


            '-----------  calculate number of parts in PO ---------------
            Dim numRow As Integer : numRow = 2

            While (wb.ActiveSheet.Cells(numRow, 3).Value IsNot Nothing)
                numRow = numRow + 1
            End While
            '-------------------------------------------------

            For i = 2 To numRow - 1
                Call Insert_rate_data(Company, wb.Worksheets(1).Cells(i, 3).value, file_date, wb.Worksheets(1).Cells(i, 7).value)
            Next


            process_ls = process_ls + 1
            process_l.Text = "Files Processed " & process_ls  'count processed files
            TextBox2.AppendText(file_po & vbCrLf)

        End If

        wb.Close(False)


    End Sub

    '-------------- enter the rates in vendor table -------------------
    Sub Insert_rate_data(Company As String, part_name As String, date_po As Date, rate As String)

        'convert mm/dd/yy to yy-mm-dd
        If String.Equals(date_po, "") = False Then
            Dim left_o() As String
            Dim new_date() As String
            new_date = date_po.ToString.Split("/")
            left_o = new_date(2).Split(" ")
            date_po = left_o(0) & "-" & new_date(0) & "-" & new_date(1)
        End If


        'get rid of quotes in PLC components, extra spaces and replace tabs for spaces in all parts

        part_name = part_name.Replace("""", "")
        part_name = Replace(part_name, vbTab, " ")
        part_name = Regex.Replace(part_name, "\s{2,}", " ")


        'Make sure McNaughton-McKay is display instead of an abbreviation
        If String.Compare("MCMC", Company, True) = True Or String.Compare("McM", Company, True) = True Or String.Compare("MMC", Company, True) = True Or String.Compare("Mc-Mc", Company) = True Or
                String.Compare("MCM", Company, True) = True Or String.Compare("Macm", Company, True) = True Or String.Compare("MCNMC", Company, True) = True Then

            Company = "McNaughton-McKay"

        End If

        '----------- MAKE SURE THE PART EXIST ------------
        Try
            Dim cmd As New MySqlCommand
            Dim Create_cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_Name", part_name)
            cmd.CommandText = "SELECT Part_Name from parts_table where Part_Name = @Part_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader


            If reader.HasRows Then
                'IF IT DOES EXIST THEN ADD RATES TO THE VENDOR TABLE WITH THE CORRESPONDING PART. 
                reader.Close()

                Create_cmd.Parameters.AddWithValue("@Part_Name", part_name)
                Create_cmd.Parameters.AddWithValue("@Vendor_Name", Company)
                Create_cmd.Parameters.AddWithValue("@Vendor_Number", "")
                Create_cmd.Parameters.AddWithValue("@Cost", rate)
                Create_cmd.Parameters.AddWithValue("@Purchase_Date", date_po)

                Create_cmd.CommandText = "INSERT INTO vendors_table(Part_Name, Vendor_Name, Vendor_Number, Cost, Purchase_Date) VALUES (@Part_Name, @Vendor_Name, @Vendor_Number, @Cost, @Purchase_Date)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Else
                'DO NOT FORGET TO CLOSE THE READER
                reader.Close()
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub EnterPO_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '--- turning on and off inventory table experimental
        If String.Equals(Button2.Text, "On") = True Then

            Button2.Text = "Off"
            enable_mess = False
        Else
            Button2.Text = "On"
            enable_mess = True
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Connections_win.Visible = True
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
End Class