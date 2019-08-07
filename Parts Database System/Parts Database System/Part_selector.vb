Imports MySql.Data.MySqlClient

Public Class Part_selector


    Private Sub Part_selector_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Part_Description, Manufacturer, Part_Type, Legacy_ADA_Number, Primary_Vendor, MFG_type, Part_Status from parts_table order by Part_Name", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            allParts.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill


            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210
            allParts.Columns(7).Width = 210

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Select Case Inventory_manage.part_sel

            Case "active"
                '------- Add part to inventory table
                Dim exist_c As Boolean : exist_c = False

                Try
                    If (allParts.CurrentCell.ColumnIndex = 0) Then

                        If String.Equals(allParts.CurrentCell.Value.ToString, "") = False Then

                            '--check if part exist
                            Dim check_cmd As New MySqlCommand
                            check_cmd.Parameters.AddWithValue("@part_name", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value)
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
                                main_cmd.Parameters.AddWithValue("@part_name", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value)
                                main_cmd.Parameters.AddWithValue("@description", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(1).Value)
                                main_cmd.Parameters.AddWithValue("@manufacturer", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(2).Value)
                                main_cmd.Parameters.AddWithValue("@MFG_type", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value)
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




                                MessageBox.Show("Part added to inventory table successfully!")
                                Call Inventory_manage.refresh_table() 'refresh table in inventory management
                            Else
                                MessageBox.Show("Part already exist in inventory!")
                            End If
                        End If

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Case "revision"

                '-- add part to material request grid

                If (allParts.CurrentCell.ColumnIndex = 0) Then

                    If String.Equals(allParts.CurrentCell.Value.ToString, "") = False Then

                        Dim my_assemblies = New List(Of String)()

                        '--------------  add to device
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
                        '-----------------------------------


                        Dim qty_change As Integer
                        Dim qty_test As Integer

                        Try
                            qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString & " to add")
                            If Integer.TryParse(qty_change, qty_test) Then

                                '---
                                'If my_assem.ContainsKey(allParts.CurrentCell.Value.ToString) = True Then

                                '    Call break_part_r(allParts.CurrentCell.Value.ToString, qty_change)
                                'Else

                                '-- see if part already exist if it does update qty otherwise add new row
                                Dim found_po As Boolean : found_po = False
                                Dim rowindex As Integer
                                Dim k As Integer : k = 0

                                For Each row As DataGridViewRow In My_Material_r.rev_grid.Rows
                                    If String.IsNullOrEmpty(My_Material_r.rev_grid.Rows(k).Cells(0).Value) = False Then
                                        If String.Compare(My_Material_r.rev_grid.Rows(k).Cells(0).Value, allParts.CurrentCell.Value.ToString) = 0 And String.IsNullOrEmpty(My_Material_r.rev_grid.Rows(k).Cells(10).Value) = True Then
                                            rowindex = row.Index
                                            My_Material_r.rev_grid.CurrentCell = My_Material_r.rev_grid.Rows(rowindex).Cells(0)
                                            My_Material_r.rev_grid.Rows(rowindex).Cells(9).Value = qty_change  ' My_Material_r.rev_grid.Rows(rowindex).Cells(8).Value = qty_change
                                            found_po = True
                                            Exit For
                                        End If

                                    End If
                                    k = k + 1
                                Next


                                If found_po = False Then
                                    Dim cost As Decimal : cost = 0

                                    If my_assemblies.Contains(allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value) = True Then
                                        cost = myQuote.Cost_of_Assem(allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value)
                                    Else
                                        cost = Form1.Get_Latest_Cost(Login.Connection, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value)
                                    End If

                                    My_Material_r.rev_grid.Rows.Add({allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(1).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(4).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(2).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value, cost, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value, 0, qty_change, "", "", "", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(7).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(3).Value})
                                End If
                                '-----------------------------------------------------
                                ' End If
                            Else
                                MsgBox("Please input an integer.")
                            End If
                        Catch ex As Exception
                            MsgBox("Please input a whole number.")
                            ' MessageBox.Show(ex.ToString)
                        End Try

                    End If

                End If


            Case Else

                '-- add part to material request grid
                If allParts.Rows.Count > 0 Then

                    If (allParts.CurrentCell.ColumnIndex = 0) Then

                        If String.Equals(allParts.CurrentCell.Value.ToString, "") = False Then

                            Dim my_assemblies = New List(Of String)()

                            '--------------  add to device
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
                            '-----------------------------------

                            Dim qty_change As Integer
                            Dim qty_test As Integer

                            Try
                                qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString & " to add")
                                If Integer.TryParse(qty_change, qty_test) Then

                                    '---
                                    'If my_assem.ContainsKey(allParts.CurrentCell.Value.ToString) = True Then

                                    '    Call break_part(allParts.CurrentCell.Value.ToString, qty_change)
                                    'Else

                                    '-- see if part already exist if it does update qty otherwise add new row
                                    Dim found_po As Boolean : found_po = False
                                    Dim rowindex As Integer
                                    Dim k As Integer : k = 0

                                    For Each row As DataGridViewRow In My_Material_r.PR_grid.Rows
                                        If String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(k).Cells(0).Value) = False Then
                                            If String.Compare(My_Material_r.PR_grid.Rows(k).Cells(0).Value, allParts.CurrentCell.Value.ToString) = 0 And String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(k).Cells(9).Value) = True Then
                                                rowindex = row.Index
                                                My_Material_r.PR_grid.CurrentCell = My_Material_r.PR_grid.Rows(rowindex).Cells(0)
                                                My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value = If(IsNumeric(My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value), My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value, 0) + qty_change
                                                found_po = True
                                                Exit For
                                            End If

                                        End If
                                        k = k + 1
                                    Next


                                    If found_po = False Then
                                        Dim cost As Decimal : cost = 0

                                        If my_assemblies.Contains(allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value) = True Then
                                            cost = myQuote.Cost_of_Assem(allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value)
                                        Else
                                            cost = Form1.Get_Latest_Cost(Login.Connection, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value)
                                        End If


                                        My_Material_r.PR_grid.Rows.Add({allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(1).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(4).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(2).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value, cost, qty_change, "", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value, "", "", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(7).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(3).Value})
                                    End If
                                    '-----------------------------------------------------
                                    ' End If
                                Else
                                    MsgBox("Please input an integer.")
                                End If
                            Catch ex As Exception
                                MsgBox("Please input a whole number.")  '"Please input a whole number."
                            End Try

                        End If

                    End If

                End If

        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '============== Partial match ============================
        Try
            Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Part_Description, Manufacturer, Part_Type,  Legacy_ADA_Number, Primary_Vendor, MFG_type, Part_status  from parts_table where Part_Name LIKE  @search"
            Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
            cmd.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")


            Dim table_s As New DataTable
            Dim adapter_s As New MySqlDataAdapter(cmd)

            adapter_s.Fill(table_s)
            allParts.DataSource = table_s

            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Part_selector_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Inventory_manage.part_sel = ""
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'show all parts
        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Part_Description, Manufacturer, Part_Type, Legacy_ADA_Number, Primary_Vendor, MFG_type, Part_Status from parts_table order by Part_Name", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            allParts.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill


            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub break_part(component As String, qt As Double)

        '--------- break assembly -------------------
        Dim Atronix_n As String : Atronix_n = ""

        Try
            Dim cmd_a As New MySqlCommand
            cmd_a.Parameters.AddWithValue("@assembly", component)
            cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @assembly"
            cmd_a.Connection = Login.Connection

            Dim reader_k As MySqlDataReader
            reader_k = cmd_a.ExecuteReader


            If reader_k.HasRows Then
                While reader_k.Read
                    Atronix_n = reader_k(0)
                End While
            End If

            reader_k.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '-------------------------------------------------------------------------

        Try
            Dim cmd_pd As New MySqlCommand
            cmd_pd.Parameters.AddWithValue("@adv_n", Atronix_n)
            cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, p1.Manufacturer,  p1.Primary_Vendor, adv.Qty, p1.MFG_type, p1.Part_Status, p1.Part_Type from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number =  @adv_n"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read

                    Dim found_po As Boolean : found_po = False
                    Dim rowindex As Integer
                    Dim i As Integer : i = 0

                    For Each row As DataGridViewRow In My_Material_r.PR_grid.Rows
                        If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
                            If String.Compare(row.Cells.Item("Column10").Value.ToString, readerv(0).ToString) = 0 And String.Equals(My_Material_r.PR_grid.Rows(i).Cells(9).Value, component) = True Then
                                rowindex = row.Index
                                My_Material_r.PR_grid.CurrentCell = My_Material_r.PR_grid.Rows(rowindex).Cells(0)
                                My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value = If(IsNumeric(My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value), My_Material_r.PR_grid.Rows(rowindex).Cells(6).Value, 0) + (readerv(4) * qt)
                                found_po = True
                                Exit For
                            End If

                        End If
                        i = i + 1
                    Next


                    If found_po = False Then

                        My_Material_r.PR_grid.Rows.Add({readerv(0).ToString, readerv(1).ToString, "", readerv(2).ToString, readerv(3).ToString, 0, readerv(4) * qt, "", readerv(5).ToString, component, "", readerv(6).ToString, readerv(7).ToString})

                    End If
                End While
            End If

            readerv.Close()

            '--- read only and color assemblies components
            Dim j As Integer : j = 0
            Dim cost As Decimal : cost = 0

            For Each row As DataGridViewRow In My_Material_r.PR_grid.Rows
                If My_Material_r.PR_grid.Rows(j).Cells(9).Value Is Nothing = False Then

                    If String.Equals(My_Material_r.PR_grid.Rows(j).Cells(9).Value, "") = False Then

                        cost = Form1.Get_Latest_Cost(Login.Connection, My_Material_r.PR_grid.Rows(j).Cells(0).Value, My_Material_r.PR_grid.Rows(j).Cells(4).Value)
                        My_Material_r.PR_grid.Rows(j).Cells(5).Value = cost
                        My_Material_r.PR_grid.Rows(j).DefaultCellStyle.BackColor = Color.Tan
                        My_Material_r.PR_grid.Rows(j).ReadOnly = True
                    End If
                End If
                j = j + 1
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


    Sub break_part_r(component As String, qt As Double)

        '--------- break assembly -------------------
        Dim Atronix_n As String : Atronix_n = ""

        Try
            Dim cmd_a As New MySqlCommand
            cmd_a.Parameters.AddWithValue("@assembly", component)
            cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @assembly"
            cmd_a.Connection = Login.Connection

            Dim reader_k As MySqlDataReader
            reader_k = cmd_a.ExecuteReader


            If reader_k.HasRows Then
                While reader_k.Read
                    Atronix_n = reader_k(0)
                End While
            End If

            reader_k.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '-------------------------------------------------------------------------

        Try
            Dim cmd_pd As New MySqlCommand
            cmd_pd.Parameters.AddWithValue("@adv_n", Atronix_n)
            cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, p1.Manufacturer,  p1.Primary_Vendor, adv.Qty, p1.MFG_type, p1.Part_Status, p1.Part_Type from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number =  @adv_n"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read

                    Dim found_po As Boolean : found_po = False
                    Dim rowindex As Integer
                    Dim i As Integer : i = 0

                    For Each row As DataGridViewRow In My_Material_r.rev_grid.Rows
                        If String.IsNullOrEmpty(row.Cells.Item("DataGridViewTextBoxColumn1").Value) = False Then
                            If String.Compare(row.Cells.Item("DataGridViewTextBoxColumn1").Value.ToString, readerv(0).ToString) = 0 And String.Equals(My_Material_r.rev_grid.Rows(i).Cells(10).Value, component) = True Then
                                rowindex = row.Index
                                My_Material_r.rev_grid.CurrentCell = My_Material_r.rev_grid.Rows(rowindex).Cells(0)
                                My_Material_r.rev_grid.Rows(rowindex).Cells(8).Value = (readerv(4) * qt)

                                found_po = True
                                Exit For
                            End If

                        End If
                        i = i + 1
                    Next


                    If found_po = False Then

                        My_Material_r.rev_grid.Rows.Add({readerv(0).ToString, readerv(1).ToString, "", readerv(2).ToString, readerv(3).ToString, 0, readerv(5).ToString, 0, readerv(4) * qt, "", component, "", readerv(6).ToString, readerv(7).ToString})

                    End If
                End While
            End If

            readerv.Close()

            Dim j As Integer : j = 0
            Dim cost As Decimal : cost = 0

            For Each row As DataGridViewRow In My_Material_r.rev_grid.Rows
                If My_Material_r.rev_grid.Rows(j).Cells(10).Value Is Nothing = False Then

                    If String.Equals(My_Material_r.rev_grid.Rows(j).Cells(10).Value, "") = False Then

                        cost = Form1.Get_Latest_Cost(Login.Connection, My_Material_r.rev_grid.Rows(j).Cells(0).Value, My_Material_r.rev_grid.Rows(j).Cells(4).Value)
                        My_Material_r.rev_grid.Rows(j).Cells(5).Value = cost

                    End If
                End If
                j = j + 1
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
End Class