Imports System.Text
Imports MySql.Data.MySqlClient


Public Class ImportCSV

    Public fName As String
    Public vars As String
    Public myRows As String
    Public skips As New Collection
    Public status As New Collection
    Public types As New Collection
    Public skips_b As Boolean

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ImpCSV.Click

        '====================== IMPORT CSV FILE BUTTON ===========================

        'Display selected CSV file to DataGridView
        myRows = 0
        Dim fName As String : fName = ""
        Dim TextLine As String = ""
        Dim SplitLine() As String
        Dim SplitLine2() As String
        Dim TextLine2 As String = ""
        Dim count_C As Integer
        Dim temp_count_C As Integer : temp_count_C = 1

        OpenFileDialog1.Filter = "CSV files(*.csv)|*.csv"


        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            fName = OpenFileDialog1.FileName
        End If

        If System.IO.File.Exists(fName) = True Then

            Try

                Dim objReader As New System.IO.StreamReader(fName, Encoding.ASCII)
                Dim objReader2 As New System.IO.StreamReader(fName, Encoding.ASCII)

                'Get number of columns for the datagridview
                Do While objReader2.Peek() <> -1
                    TextLine2 = objReader2.ReadLine()
                    SplitLine2 = Split(TextLine2, ",")
                    count_C = SplitLine2.Length

                    If count_C > temp_count_C Then
                        temp_count_C = count_C
                    End If
                Loop

                'Add columns to datagridview
                For i = 1 To temp_count_C '- 1
                    DataGridView1.Columns.Add(New DataGridViewTextBoxColumn())
                    myRows = myRows + 1
                Next


                Do While objReader.Peek() <> -1
                    TextLine = objReader.ReadLine()
                    SplitLine = Split(TextLine, ",")
                    Me.DataGridView1.Rows.Add(SplitLine)
                Loop

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            MessageBox.Show("File does not exist")
        End If

    End Sub

    Private Sub Button7_MouseEnter(sender As Object, e As EventArgs) Handles EnterR.MouseEnter
        EnterR.BackColor = Color.DarkCyan
    End Sub

    Private Sub Button7_MouseLeave(sender As Object, e As EventArgs) Handles EnterR.MouseLeave
        EnterR.BackColor = Color.Maroon
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '========================= CLEAR BUTTON ==========================

        skips = New Collection

        'Remove columns from Datagridview (leave only one)
        For i = 1 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns.RemoveAt(0)
        Next
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        '======================= CREATE TEXT FILE FROM DATAGRID =========================
        'Creates a Text file from DataGridView

        DataGridView1.Refresh()
        Dim rows = From row As DataGridViewRow In DataGridView1.Rows.Cast(Of DataGridViewRow)()
                   Where Not row.IsNewRow
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        Using sw As New IO.StreamWriter("csv.txt")
            For Each r In rows
                sw.WriteLine(String.Join(",", r))
            Next

        End Using
        Process.Start("csv.txt")



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '==================== DELETE ROWS BUTTON ===========================

        If DataGridView1.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(r)
            Next
        Else
            MessageBox.Show("Select 1 row before you hit Delete")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        '-------SKIP SOME ROWS--------------------------

        If skips_b = True Then
            MessageBox.Show("No Rows will be skipped")

            'Clear color of skip rows
            For Each r As DataGridViewRow In DataGridView1.Rows
                DataGridView1.Rows(r.Index).DefaultCellStyle.BackColor = SystemColors.Window
            Next


            skips.Clear()
            skips_b = False
        Else
            'add rows to be skipped to skip collection

            skips = New Collection
            TextBox1.Text = "Skip Rows: "

            If DataGridView1.SelectedRows.Count > 0 Then
                For Each r As DataGridViewRow In DataGridView1.SelectedRows
                    skips.Add(CType((r.Index + 1), String))
                    'color row
                    DataGridView1.Rows(r.Index).DefaultCellStyle.BackColor = Color.Teal
                Next

                For i = 1 To skips.Count
                    TextBox1.AppendText(skips.Item(i).ToString & ", ")
                Next
                skips_b = True
            Else
                MessageBox.Show("Select 1 row before you hit Skip Rows")
            End If


        End If
    End Sub


    Function Skip_row(skips As Collection, index As Integer) As Boolean

        'return true if the index is contained in the collection
        Skip_row = False

        For Each r In skips
            If String.Equals(r, index.ToString) = True Then
                Skip_row = True
            End If
        Next
    End Function

    Private Sub Button8_MouseEnter(sender As Object, e As EventArgs) Handles Button8.MouseEnter
        Button8.BackColor = Color.DarkCyan
    End Sub

    Private Sub Button8_MouseLeave(sender As Object, e As EventArgs) Handles Button8.MouseLeave
        Button8.BackColor = Color.Maroon
    End Sub

    Private Sub ImportCSV_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged

        'Reset skip collection and clean datagrid
        skips = New Collection

        'Remove columns from Datagridview (leave only one)
        For i = 1 To DataGridView1.ColumnCount - 1
            DataGridView1.Columns.RemoveAt(0)
        Next
        DataGridView1.Rows.Clear()
        TextBox1.Text = ""
    End Sub

    Private Sub EnterR_Click(sender As Object, e As EventArgs) Handles EnterR.Click

        'Dim myTable As String : myTable = ""

        '' Get Table to insert data
        'If Not Table_box.SelectedItem Is Nothing Then
        '    myTable = Table_box.SelectedItem.ToString
        'Else
        '    myTable = ""
        'End If


        'If String.Equals(myTable, "") = False Then

        '    Select Case myTable

        '        Case "Parts Table"
        '            'Insert data in parts and vendor table
        '            Call Insert_Data_parts()

        '            'Case "AKN Table"
        '            'Insert data in AKN and Kits table
        '            ' Call Insert_Data_AKN()
        '            ' Case "ADV Table"
        '            'Insert data in ADV and Device table
        '            ' Call Insert_Data_ADV()
        '            ' Case "Vendors Table"
        '            '  Call Insert_Vendor()

        '    End Select

        'Else
        '    MessageBox.Show("No Table Selected!")
        'End If
        Call Insert_Data_parts()


    End Sub

    Sub Insert_Data_parts()

        '=========== INSERT DATA IN PARTS AND VENDORS TABLE ================


        Dim Vendor_data(27) As String
        Dim mes As String : mes = "Unknown Error. Check all values"
        Dim Part_Name As String : Part_Name = ""
        Dim Manufacturer As String : Manufacturer = ""
        Dim Part_Description As String : Part_Description = ""
        Dim Notes As String : Notes = ""
        Dim Part_Status As String : Part_Status = ""
        Dim Part_Type As String : Part_Type = ""
        Dim Units As String : Units = ""
        Dim Min_Qty As String : Min_Qty = 1
        Dim Ada As String : Ada = ""
        Dim p_vendor As String : p_vendor = ""
        Dim flag_pass As Boolean : flag_pass = True


        If myRows >= 10 Then

            '================================================== ENTER DATA to table from datagriview ===============================================
            'Go through every row
            For rowIndex = 0 To DataGridView1.RowCount - 2

                If Skip_row(skips, rowIndex + 1.ToString) = True Then Continue For   'Skips the rows selected

                Try
                    Dim command As New MySqlCommand
                    Dim myCommand As MySqlCommand = New MySqlCommand("SHOW WARNINGS")

                    '-- get part data values
                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(0).Value) = False Then
                        Part_Name = DataGridView1.Rows(rowIndex).Cells(0).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(1).Value) = False Then
                        Manufacturer = DataGridView1.Rows(rowIndex).Cells(1).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(2).Value) = False Then
                        Part_Description = DataGridView1.Rows(rowIndex).Cells(2).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(3).Value) = False Then
                        Notes = DataGridView1.Rows(rowIndex).Cells(3).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(4).Value) = False Then
                        Part_Status = DataGridView1.Rows(rowIndex).Cells(4).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(5).Value) = False Then
                        Part_Type = DataGridView1.Rows(rowIndex).Cells(5).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(6).Value) = False Then
                        Units = DataGridView1.Rows(rowIndex).Cells(6).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(7).Value) = False Then
                        Min_Qty = DataGridView1.Rows(rowIndex).Cells(7).Value.ToString
                        If String.Equals(Min_Qty, "") Then
                            Min_Qty = 1
                        End If
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(8).Value) = False Then
                        Ada = DataGridView1.Rows(rowIndex).Cells(8).Value.ToString
                    End If

                    If IsNothing(DataGridView1.Rows(rowIndex).Cells(9).Value) = False Then
                        p_vendor = DataGridView1.Rows(rowIndex).Cells(9).Value.ToString
                    Else
                        p_vendor = ""
                    End If

                    'initialize values of the array to avoid reference errors
                    For j = 0 To Vendor_data.Length - 1
                        Vendor_data(j) = ""
                    Next

                    '========================= Get vendor data =========================
                    For i = 9 To DataGridView1.ColumnCount - 1
                        If (IsNothing(DataGridView1.Rows(rowIndex).Cells(i).Value) = False) Then
                            If DataGridView1.Rows(rowIndex).Cells(i).Value.ToString.Contains("$") = True Then
                                Vendor_data(i - 9) = DataGridView1.Rows(rowIndex).Cells(i).Value.ToString.Replace("$", "")
                            Else
                                Vendor_data(i - 9) = DataGridView1.Rows(rowIndex).Cells(i).Value.ToString
                            End If
                        End If
                    Next

                    '================ CHECK VALID PART STATUS ======================
                    If IsOntheList(status, Part_Status) = False Then
                        mes = "Invalid Part Status: " & Part_Status
                        flag_pass = False
                    End If

                    '================ CHECK VALID PART TYPE ======================
                    If IsOntheList(types, Part_Type) = False Then
                        mes = "Invalid Part Type: " & Part_Type
                        flag_pass = False
                    End If

                    '================ CHECK VALID Min_QTY ======================
                    If Not IsNumeric(Min_Qty) Then
                        mes = "Invalid Min Qty: " & Min_Qty
                        flag_pass = False
                    End If
                    '----------------------------------------------------------

                    If flag_pass = True Then
                        'Insert into Parts table
                        command.Parameters.AddWithValue("@Part_Name", Part_Name) 'Part name
                        command.Parameters.AddWithValue("@Manuf", Manufacturer)  'manufacturer
                        command.Parameters.AddWithValue("@Desc", Part_Description)  'description
                        command.Parameters.AddWithValue("@Notes", Notes)  'notes
                        command.Parameters.AddWithValue("@Status", Part_Status)  'status
                        command.Parameters.AddWithValue("@Type", Part_Type)  'type
                        command.Parameters.AddWithValue("@Units", Units)  'units
                        command.Parameters.AddWithValue("@Min", Min_Qty)   'min_qty
                        command.Parameters.AddWithValue("@ada", Ada)   'ada
                        command.Parameters.AddWithValue("@pvendor", p_vendor)   'p_vendor


                        command.CommandText = "INSERT INTO parts_table values(@Part_Name, @Manuf, @Desc, @Notes, @Status, @Type, @Units, @Min, @ada, @pvendor)"
                        command.Connection = Login.Connection
                        command.ExecuteNonQuery()
                        myCommand.Connection = Login.Connection
                        Dim reader As MySqlDataReader = myCommand.ExecuteReader()
                        TextBox1.Text = "MySQL: Parts inserted succesfully!" & vbCrLf

                        'Write the warnings to the textbox
                        While reader.Read()
                            TextBox1.AppendText(reader(0).ToString() & " " & reader(1).ToString() & " " & reader(2).ToString() & vbCrLf)
                        End While
                        reader.Close()

                        For i = 0 To Vendor_data.Length - 1
                            If Vendor_data(i).Equals("") = False Or Vendor_data(i + 1).Equals("") = False Or Vendor_data(i + 2).Equals("") = False Or Vendor_data(i + 3).Equals("") = False Then

                                Try
                                    ' ============================ INSERT TO VENDOR TABLE ==========================

                                    Dim command2 As New MySqlCommand
                                    Dim myCommand2 As MySqlCommand = New MySqlCommand("SHOW WARNINGS")

                                    command2.CommandText = "INSERT INTO vendors_table values(@Part_Name, @vendor_n, @vendor_num, @cost, @date)"
                                    command2.Parameters.AddWithValue("@Part_Name", Part_Name)
                                    command2.Parameters.AddWithValue("@vendor_n", Vendor_data(i))
                                    command2.Parameters.AddWithValue("@vendor_num", Vendor_data(i + 1))
                                    command2.Parameters.AddWithValue("@cost", Vendor_data(i + 2))
                                    command2.Parameters.AddWithValue("@date", Vendor_data(i + 3))

                                    command2.Connection = Login.Connection
                                    command2.ExecuteNonQuery()
                                    myCommand2.Connection = Login.Connection
                                    Dim reader2 As MySqlDataReader = myCommand2.ExecuteReader()

                                    TextBox1.AppendText("  Vendors inserted succesfully!" & vbCrLf)

                                    'Write the warnings to the textbox
                                    While reader2.Read()
                                        TextBox1.AppendText(reader2(0).ToString() & " " & reader2(1).ToString() & " " & reader2(2).ToString() & vbCrLf)
                                    End While
                                    reader2.Close()

                                Catch ex As Exception
                                    DataGridView1.Rows(rowIndex).Selected = True  'Highlight row with the issue
                                    TextBox1.Text = ex.ToString                   'Write the error to the textbox
                                    rowIndex = DataGridView1.ColumnCount + 100
                                    Exit For  'End the for loop
                                End Try

                            End If
                            i = i + 3
                        Next

                    Else

                        'Error regarding data type,staus,min_qty
                        DataGridView1.Rows(rowIndex).Selected = True  'Highlight row with the issue
                        TextBox1.Text = mes
                        rowIndex = DataGridView1.RowCount + 1000
                        Exit For

                    End If

                    'Clear variables
                    Part_Name = ""
                    Manufacturer = ""
                    Part_Description = ""
                    Notes = ""
                    Part_Status = ""
                    Part_Type = ""
                    Units = ""
                    Min_Qty = 1
                    Ada = ""
                    p_vendor = ""

                    For j = 0 To Vendor_data.Length - 1
                        Vendor_data(j) = ""
                    Next

                Catch ex As MySqlException
                    DataGridView1.Rows(rowIndex).Selected = True  'Highlight row with the issue
                    TextBox1.Text = ex.ToString                   'Write the error to the textbox
                    TextBox1.AppendText(vbCrLf & "DO NOT USE THE $ SIGN IN ANY FIELD EXCEPT COST ")
                    Exit For  'End the for loop
                End Try

            Next

        Else
            MessageBox.Show("The table must have at least 10 columns before action can be executed")
        End If

    End Sub

    Sub Delete_data()

        '================ Delete data from table ===========================
        For rowIndex = 0 To DataGridView1.RowCount - 2
            If Skip_row(skips, rowIndex + 1.ToString) = True Then Continue For
            Try
                Dim command_e As New MySqlCommand
                Dim command_v As New MySqlCommand
                Dim myCommand_e As MySqlCommand = New MySqlCommand("SHOW WARNINGS")


                If IsNothing(DataGridView1.Rows(rowIndex).Cells(0).Value) = False And DataGridView1.Rows(rowIndex).Cells(0).Value.Equals("") = False Then
                    command_e.Parameters.AddWithValue("@Part_Name", DataGridView1.Rows(rowIndex).Cells(0).Value.ToString)
                    command_e.CommandText = "DELETE FROM parts_table where Part_Name = @Part_Name"
                    command_e.Connection = Login.Connection
                    command_e.ExecuteNonQuery()

                    '========= delete vendors attached to that part =================
                    command_v.Parameters.AddWithValue("@Part_Name", DataGridView1.Rows(rowIndex).Cells(0).Value.ToString)
                    command_v.CommandText = "DELETE FROM vendors_table where Part_Name = @Part_Name"
                    command_v.Connection = Login.Connection
                    command_v.ExecuteNonQuery()


                    myCommand_e.Connection = Login.Connection
                    Dim reader As MySqlDataReader = myCommand_e.ExecuteReader()

                    While reader.Read()
                        TextBox1.AppendText(reader(0).ToString() & " " & reader(1).ToString() & " " & reader(2).ToString() & vbCrLf)
                    End While
                    reader.Close()

                    TextBox1.Text = "Parts Removed"

                End If

            Catch ex As Exception
                DataGridView1.Rows(rowIndex).Selected = True
                TextBox1.Text = ex.ToString
                Exit For
            End Try
        Next

    End Sub

    Sub Insert_Data_ACN()

    End Sub
    Sub Insert_Data_AKN()

    End Sub
    Sub Insert_Data_ADV()

    End Sub

    Sub Insert_Vendor()

    End Sub

    Private Sub ImportCSV_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        types = New Collection
        status = New Collection

        'Populate part status
        status.Add("Preferred")
        status.Add("Alternate")
        status.Add("Custom")
        status.Add("Kanbon")


        'Populate Part Types
        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT Part_Type from parts_type_table"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    types.Add(readerv(0).ToString)
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            TextBox1.Text = ex.ToString
        End Try

    End Sub

    Function IsOntheList(c1 As Collection, value As String) As Boolean

        '========== CHECK IF VALUE FOUND ON THE LIST ====================

        IsOntheList = False
        For i = 1 To c1.Count
            If CType(c1.Item(i), String).Equals(value) Then
                IsOntheList = True
            End If
        Next i
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        '=============== DELETE ===============

        'Dim myTable As String : myTable = ""

        '' Get Table to insert data
        'If Not Table_box.SelectedItem Is Nothing Then
        '    myTable = Table_box.SelectedItem.ToString
        'Else
        '    myTable = ""
        'End If


        'If String.Equals(myTable, "") = False Then

        '    Select Case myTable

        '        Case "Parts Table"
        '            'Delete and Insert data in parts and vendor table
        Call Delete_data()

        '        Case "ACN Table"

        '        Case "AKN Table"

        '        Case "ADV Table"

        '        Case "Vendors Table"

        '    End Select

        'Else
        '    MessageBox.Show("No Table Selected!")
        'End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        ' =================== SHOW WARNINGS ======================

        'Displays the MySQL warnings
        Dim myCommand As MySqlCommand = New MySqlCommand("SHOW WARNINGS")
        myCommand.Connection = Login.Connection
        Dim reader As MySqlDataReader = myCommand.ExecuteReader()
        TextBox1.Text = ""

        If reader.HasRows Then
            While reader.Read()
                TextBox1.AppendText(reader(0).ToString() & " " & reader(1).ToString() & " " & reader(2).ToString() & vbCrLf)
            End While
        Else
            MessageBox.Show("No warnings")
        End If

        reader.Close()

    End Sub


End Class