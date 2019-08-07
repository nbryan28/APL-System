Imports MySql.Data.MySqlClient

Public Class Part_Picker

    Private mfg_t As String
    Private mygrid As DataGridView
    Private cables_l As List(Of String)



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mfg_t = "x"

    End Sub

    Sub load_data()

        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Part_Description, Manufacturer, Part_Type, Legacy_ADA_Number, Primary_Vendor, MFG_type, Part_Status from parts_table order by Part_Name", Login.Connection)

            adapter.Fill(table)     'DataViewGrid1 fill
            allParts.DataSource = table

            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210
            allParts.Columns(7).Width = 210

            For i = 0 To 7
                allParts.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Me.Width = 1
        Me.Height = 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '============== Partial match ============================
        Try
            Dim type_f As String : type_f = "Part_Name"

            If RadioButton1.Checked = True Then
                type_f = "Part_Name"
            ElseIf RadioButton2.Checked = True Then
                type_f = "Part_Description"
            ElseIf RadioButton3.Checked = True Then
                type_f = "Manufacturer"
            ElseIf RadioButton4.Checked = True Then
                type_f = "Primary_Vendor"
            ElseIf RadioButton5.Checked = True Then
                type_f = "Legacy_ADA_Number"

            End If



            Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Part_Description, Manufacturer, Part_Type,  Legacy_ADA_Number, Primary_Vendor, MFG_type, Part_status  from parts_table where " & type_f & " LIKE  @search"
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

    Private Sub allParts_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles allParts.CellContentClick

        Dim index_k = allParts.CurrentCell.RowIndex

        part_s.Text = allParts.Rows(index_k).Cells(0).Value.ToString

    End Sub

    Sub specify(ByRef data As DataGridView)
        '-- specify datagridview to enter data
        mygrid = data
    End Sub

    Sub mfg_type_Set(mfg_type)
        mfg_t = mfg_type
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- add part to grid specified
        If String.Equals(qty_box.Text, "0") = True Then
            MessageBox.Show("Warning! Qty set to zero")

        Else

            If String.Equals(mygrid.Name, "rev_grid") = False Then
                Call enter_data(mygrid)
            Else
                Call enter_data_rev(mygrid)
            End If

            '--check if its a cable---
            'If String.Equals(part_s.Text.ToString, "ROB-1096") = True Or String.Equals(part_s.Text.ToString, "10386-08547") = True Then

            '    Cables_package.Text = part_s.Text.ToString
            '    Cables_package.feet_qty.Text = qty_box.Text
            '    Cables_package.ShowDialog()

            'End If
            '-----------------------

        End If

    End Sub

    Sub enter_data(my_grid As DataGridView)
        '-- add part to grid specified

        If my_grid.Rows.Count > 0 Then

            If String.Equals(part_s.Text.ToString, "Part Selected: ") = False And IsNumeric(qty_box.Text) = True Then

                '-----------------------------------
                If qty_box.Text > 0 Then

                    Try

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


                        '-- see if part already exist if it does update qty otherwise add new row
                        Dim found_po As Boolean : found_po = False

                        Dim rowindex As Integer
                        Dim k As Integer : k = 0

                        For Each row As DataGridViewRow In my_grid.Rows
                            If String.IsNullOrEmpty(my_grid.Rows(k).Cells(0).Value) = False Then
                                If String.Compare(my_grid.Rows(k).Cells(0).Value, part_s.Text.ToString) = 0 Then
                                    rowindex = row.Index
                                    my_grid.CurrentCell = my_grid.Rows(rowindex).Cells(0)
                                    my_grid.Rows(rowindex).Cells(5).Value = If(IsNumeric(my_grid.Rows(rowindex).Cells(5).Value), my_grid.Rows(rowindex).Cells(5).Value, 0) + CType(qty_box.Text, Double)
                                    my_grid.Rows(rowindex).DefaultCellStyle.BackColor = Color.LightBlue 'color new edited row
                                    found_po = True
                                    Exit For
                                End If

                            End If
                            k = k + 1
                        Next


                        If found_po = False Then
                            Dim cost As Decimal : cost = 0

                            If my_assemblies.Contains(part_s.Text) = True Then
                                cost = myQuote.Cost_of_Assem(part_s.Text)
                            Else
                                cost = Form1.Get_Latest_Cost(Login.Connection, part_s.Text, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value)
                            End If

                            '--mfg_type
                            Dim mfg_v As String : mfg_v = allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value

                            If String.Equals(mfg_t, "x") = False Then
                                mfg_v = mfg_t

                            End If

                            my_grid.Rows.Add({allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(1).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(2).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value, cost, qty_box.Text, cost * qty_box.Text, mfg_v, "Preferred", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(3).Value})
                            my_grid.Rows(my_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.LightBlue 'color new edited row
                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If

            Else
                MessageBox.Show("Please select a Part and enter a Qty to add")
            End If
        Else
            MessageBox.Show("No Panel Selected")
        End If
    End Sub


    Sub enter_data_rev(my_grid As DataGridView)
        '--- enter data for rev table
        If my_grid.Rows.Count > 0 Then

            If String.Equals(part_s.Text.ToString, "Part Selected: ") = False And IsNumeric(qty_box.Text) = True Then

                '-----------------------------------
                If qty_box.Text > 0 Then

                    Try

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


                        '-- see if part already exist if it does update qty otherwise add new row
                        Dim found_po As Boolean : found_po = False

                        Dim rowindex As Integer
                        Dim k As Integer : k = 0

                        For Each row As DataGridViewRow In my_grid.Rows
                            If String.IsNullOrEmpty(my_grid.Rows(k).Cells(0).Value) = False Then
                                If String.Compare(my_grid.Rows(k).Cells(0).Value, part_s.Text.ToString) = 0 Then
                                    rowindex = row.Index
                                    my_grid.CurrentCell = my_grid.Rows(rowindex).Cells(0)
                                    my_grid.Rows(rowindex).Cells(9).Value = CType(qty_box.Text, Double) + If(IsNumeric(my_grid.Rows(rowindex).Cells(9).Value) = True, my_grid.Rows(rowindex).Cells(9).Value, 0)
                                    my_grid.Rows(rowindex).DefaultCellStyle.BackColor = Color.LightBlue 'color new edited row
                                    found_po = True
                                    Exit For
                                End If

                            End If
                            k = k + 1
                        Next


                        If found_po = False Then
                            Dim cost As Decimal : cost = 0

                            If my_assemblies.Contains(part_s.Text) = True Then
                                cost = myQuote.Cost_of_Assem(part_s.Text)
                            Else
                                cost = Form1.Get_Latest_Cost(Login.Connection, part_s.Text, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value)
                            End If

                            '--mfg_type
                            Dim mfg_v As String : mfg_v = allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value

                            If String.Equals(mfg_t, "x") = False Then
                                mfg_v = mfg_t

                            End If

                            my_grid.Rows.Add({allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(1).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(4).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(2).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(5).Value, cost, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(6).Value, 0, "", CType(qty_box.Text, Double), "", "", allParts.Rows(allParts.CurrentCell.RowIndex).Cells(7).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(3).Value})
                            my_grid.Rows(my_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.LightBlue 'color new edited row
                        End If

                        Call My_Material_r.recal_rev()

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If

            Else
                MessageBox.Show("Please select a Part and enter a Qty to add")
            End If
        Else
            MessageBox.Show("No Panel Selected")
        End If
    End Sub

    Private Sub allParts_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles allParts.CellClick

        Dim index_k = allParts.CurrentCell.RowIndex

        part_s.Text = allParts.Rows(index_k).Cells(0).Value.ToString
    End Sub

    Private Sub allParts_SelectionChanged(sender As Object, e As EventArgs) Handles allParts.SelectionChanged
        Dim index_k = allParts.CurrentCell.RowIndex

        part_s.Text = allParts.Rows(index_k).Cells(0).Value.ToString
    End Sub

    Private Sub MoreInfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoreInfoToolStripMenuItem.Click
        Extra_info.Text = allParts.CurrentCell.Value.ToString
        Extra_info.ShowDialog()
    End Sub

    Private Sub Part_Picker_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
