Imports System.Reflection
Imports MySql.Data.MySqlClient

Public Class Feature_sel

    Public dontdo As Boolean


    Private Sub Feature_sel_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dontdo = True
        sol_box.Items.Clear()
        Dim type_p As String : type_p = "MB"

        Try
            '--------- load solutions ---------------
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT solution_name from quote_table.feature_solutions order by solution_name"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    sol_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

            sol_box.Text = "EIP-MS/EIP-RIO" '"SWDMS/EIPRIO"

            '-------------------------------------

            If BOM_types.feature_s = False Then

                If BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage1 Then  'Panel bom

                    If String.Equals("Name:   Unknown", BOM_types.Panel_n_l.Text) = False Then

                        dontdo = False
                        BOM_types.feature_Field = False
                        Call load_features(BOM_types.feature_Field)
                        bom_name.Text = BOM_types.Panel_n_l.Text


                    Else
                        dontdo = True
                        bom_name.Text = "Not BOM found"
                    End If

                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage2 Then  'Field bom
                    dontdo = False
                    BOM_types.feature_Field = True
                    Call load_features(BOM_types.feature_Field)
                    bom_name.Text = "Field BOM"

                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage3 Then   'assem bom
                    dontdo = True
                    bom_name.Text = "ASSEM BOM. Not allowed"
                    BOM_types.feature_Field = False

                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage4 Then   'spare bom
                    dontdo = False
                    BOM_types.feature_Field = True
                    Call load_features(BOM_types.feature_Field)
                    bom_name.Text = "Spare BOM"


                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage5 Then   'master bom
                    dontdo = True
                    bom_name.Text = "Master BOM. Not allowed"
                    BOM_types.feature_Field = False
                End If

            Else
                '---- check what type of panel is and display hte right type of feature code
                Try
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                    cmd4.CommandText = "Select BOM_type from Material_Request.mr where mr_name = @mr_name"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        While reader4.Read
                            type_p = reader4(0).ToString
                        End While
                    End If

                    reader4.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                If String.Equals(type_p, "Panel") = True Then

                    dontdo = False
                    BOM_types.feature_Field = False
                    Call load_features(BOM_types.feature_Field)
                    bom_name.Text = type_p & " BOM"


                ElseIf String.Equals(type_p, "Field") = True Or String.Equals(type_p, "old_BOM") = True Or String.Equals(type_p, "Spare_Parts") = True Then

                    dontdo = False
                    BOM_types.feature_Field = True
                    Call load_features(BOM_types.feature_Field)
                    bom_name.Text = type_p & " BOM"


                End If


            End If


            If dontdo = False Then

                For i = 0 To Feature_Table.ColumnCount - 1
                    With Feature_Table.Columns(i)
                        .Width = 600
                    End With
                Next i

                Feature_Table.Columns(1).Width = 100

            End If

            Call EnableDoubleBuffered(Feature_Table)
            Call EnableDoubleBuffered(assem_grid)



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub fea_box_TextChanged(sender As Object, e As EventArgs) Handles fea_box.TextChanged

        If dontdo = False Then

            Dim cmdstr As String : cmdstr = "SELECT description, type, specific_type , Feature_code from quote_table.feature_codes where description LIKE  @search and Solution = @sol "

            If BOM_types.feature_Field = True Then
                cmdstr = cmdstr & " and type = 'Field'"
            Else
                cmdstr = cmdstr & " and type not like 'Field'"
            End If

            Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
            cmd.Parameters.AddWithValue("@search", "%" & fea_box.Text & "%")
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)


            Dim table_s As New DataTable
            Dim adapter_s As New MySqlDataAdapter(cmd)

            adapter_s.Fill(table_s)
            Feature_Table.DataSource = table_s

            For i = 0 To Feature_Table.ColumnCount - 1
                With Feature_Table.Columns(i)
                    .Width = 600
                End With
            Next i
            Feature_Table.Columns(1).Width = 100

        End If
    End Sub

    Private Sub sol_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles sol_box.SelectedValueChanged
        '----change solution --

        If dontdo = False Then

            assem_grid.Rows.Clear()
            '-------- load feature code tables -----------
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT description, type, specific_type, Feature_code from quote_table.feature_codes where Solution = @sol"

            If BOM_types.feature_Field = True Then
                cmd.CommandText = cmd.CommandText & " and type = 'Field'"
            Else
                cmd.CommandText = cmd.CommandText & " and type not like 'Field'"
            End If

            cmd.Connection = Login.Connection

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter(cmd)    'DataViewGrid1 fill
            adapter.Fill(table)
            Feature_Table.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Feature_Table.ColumnCount - 1
                With Feature_Table.Columns(i)
                    .Width = 600
                End With
            Next i

            Feature_Table.Columns(1).Width = 100

        End If
    End Sub

    Private Sub Feature_Table_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Feature_Table.CellClick

        If dontdo = False Then

            assem_grid.Rows.Clear()

            Dim index_k = Feature_Table.CurrentCell.RowIndex

            fc.Text = Feature_Table.Rows(index_k).Cells(0).Value.ToString

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@feature_code", Feature_Table.Rows(index_k).Cells(3).Value.ToString)
                cmd4.Parameters.AddWithValue("@solution", sol_box.Text)
                cmd4.CommandText = "Select qty, part_name from quote_table.feature_parts where Feature_code = @feature_code and solution = @solution"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        assem_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString})
                    End While
                End If

                reader4.Close()

                '--- parts ---

                For i = 0 To assem_grid.Rows.Count - 1
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.AddWithValue("@part", assem_grid.Rows(i).Cells(1).Value)
                    cmd41.CommandText = "Select Part_Description, Manufacturer, Primary_Vendor, MFG_type, Part_Type, Part_Status from parts_table where Part_Name = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            assem_grid.Rows(i).Cells(2).Value = reader41(0).ToString
                            assem_grid.Rows(i).Cells(3).Value = reader41(1).ToString
                            assem_grid.Rows(i).Cells(4).Value = reader41(2).ToString
                            assem_grid.Rows(i).Cells(5).Value = reader41(3).ToString
                            assem_grid.Rows(i).Cells(6).Value = reader41(4).ToString
                            assem_grid.Rows(i).Cells(7).Value = reader41(5).ToString
                        End While
                    End If

                    reader41.Close()
                Next


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        If Feature_Table.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(Feature_Table.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub


    Sub load_features(field As Boolean)
        '-- load either panel, control, plc or field feature codes ----

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd.CommandText = "SELECT description, type, specific_type, Feature_code from quote_table.feature_codes where Solution = @sol"

            '-- add more to the query
            If field = True Then
                cmd.CommandText = cmd.CommandText & " and type = 'Field'"
            Else
                cmd.CommandText = cmd.CommandText & " and type not like 'Field'"
            End If


            cmd.Connection = Login.Connection

            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter(cmd)    'DataViewGrid1 fill
            adapter.Fill(table)
            Feature_Table.DataSource = table


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Feature_sel_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then

            Feature_Table.DataSource = Nothing

            assem_grid.Rows.Clear()

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.Equals(qty_box.Text, "0") = True Then
            MessageBox.Show("Warning! Qty set to zero")

        Else

            If BOM_types.feature_s = False Then

                If BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage1 Then  'Panel bom
                    For i = 0 To assem_grid.Rows.Count - 1
                        Call enter_data(BOM_types.Panel_grid, assem_grid.Rows(i).Cells(1).Value, assem_grid.Rows(i).Cells(0).Value, assem_grid.Rows(i).Cells(2).Value, assem_grid.Rows(i).Cells(3).Value, assem_grid.Rows(i).Cells(5).Value, assem_grid.Rows(i).Cells(6).Value, assem_grid.Rows(i).Cells(7).Value, assem_grid.Rows(i).Cells(4).Value)
                    Next

                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage2 Then  'Field bom
                    For i = 0 To assem_grid.Rows.Count - 1
                        Call enter_data(BOM_types.field_grid, assem_grid.Rows(i).Cells(1).Value, assem_grid.Rows(i).Cells(0).Value, assem_grid.Rows(i).Cells(2).Value, assem_grid.Rows(i).Cells(3).Value, assem_grid.Rows(i).Cells(5).Value, assem_grid.Rows(i).Cells(6).Value, assem_grid.Rows(i).Cells(7).Value, assem_grid.Rows(i).Cells(4).Value)
                    Next

                ElseIf BOM_types.TabControl1.SelectedTab Is BOM_types.TabPage4 Then   'spare bom
                    For i = 0 To assem_grid.Rows.Count - 1
                        Call enter_data(BOM_types.sp_grid, assem_grid.Rows(i).Cells(1).Value, assem_grid.Rows(i).Cells(0).Value, assem_grid.Rows(i).Cells(2).Value, assem_grid.Rows(i).Cells(3).Value, assem_grid.Rows(i).Cells(5).Value, assem_grid.Rows(i).Cells(6).Value, assem_grid.Rows(i).Cells(7).Value, assem_grid.Rows(i).Cells(4).Value)
                    Next
                End If

            Else
                For i = 0 To assem_grid.Rows.Count - 1
                    Call enter_data_rev(My_Material_r.rev_grid, assem_grid.Rows(i).Cells(1).Value, assem_grid.Rows(i).Cells(0).Value, assem_grid.Rows(i).Cells(2).Value, assem_grid.Rows(i).Cells(3).Value, assem_grid.Rows(i).Cells(5).Value, assem_grid.Rows(i).Cells(6).Value, assem_grid.Rows(i).Cells(7).Value, assem_grid.Rows(i).Cells(4).Value)
                Next
            End If

        End If
    End Sub

    Sub enter_data(my_grid As DataGridView, part_s As String, f_Qty As String, desc As String, mfg As String, mfg_type As String, part_Type As String, preff As String, vendor As String)
        '-- add part to grid specified

        If my_grid.Rows.Count > 0 Then

            If String.Equals(fc.Text.ToString, "Not selected") = False And IsNumeric(qty_box.Text) = True Then

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
                                If String.Compare(my_grid.Rows(k).Cells(0).Value, part_s) = 0 Then
                                    rowindex = row.Index
                                    my_grid.CurrentCell = my_grid.Rows(rowindex).Cells(0)
                                    my_grid.Rows(rowindex).Cells(5).Value = If(IsNumeric(my_grid.Rows(rowindex).Cells(5).Value), my_grid.Rows(rowindex).Cells(5).Value, 0) + CType(qty_box.Text, Double) * f_Qty
                                    my_grid.Rows(rowindex).DefaultCellStyle.BackColor = Color.Bisque 'color new edited row
                                    found_po = True
                                    Exit For
                                End If

                            End If
                            k = k + 1
                        Next


                        If found_po = False Then
                            Dim cost As Decimal : cost = 0

                            If my_assemblies.Contains(part_s) = True Then
                                cost = myQuote.Cost_of_Assem(part_s)
                            Else
                                cost = Form1.Get_Latest_Cost(Login.Connection, part_s, vendor)
                            End If

                            my_grid.Rows.Add({part_s, desc, mfg, vendor, cost, f_Qty * qty_box.Text, cost * (f_Qty * qty_box.Text), mfg_type, preff, part_Type})  'add row
                            my_grid.Rows(my_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.Bisque 'color new edited row
                        End If

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If

            Else
                MessageBox.Show("Please select a Feature code and enter a Qty to add")
            End If
        Else
            MessageBox.Show("No Panel Selected")
        End If
    End Sub


    Sub enter_data_rev(my_grid As DataGridView, part_s As String, f_Qty As String, desc As String, mfg As String, mfg_type As String, part_Type As String, preff As String, vendor As String)
        '--- enter data for rev table
        If my_grid.Rows.Count > 0 Then

            If String.Equals(part_s, "Not selected") = False And IsNumeric(qty_box.Text) = True Then

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
                                If String.Compare(my_grid.Rows(k).Cells(0).Value, part_s) = 0 Then
                                    rowindex = row.Index
                                    my_grid.CurrentCell = my_grid.Rows(rowindex).Cells(0)
                                    my_grid.Rows(rowindex).Cells(9).Value = (CType(qty_box.Text, Double) * f_Qty) + If(IsNumeric(my_grid.Rows(rowindex).Cells(9).Value), my_grid.Rows(rowindex).Cells(9).Value, 0)
                                    my_grid.Rows(rowindex).DefaultCellStyle.BackColor = Color.Bisque 'color new edited row
                                    found_po = True
                                    Exit For
                                End If

                            End If
                            k = k + 1
                        Next


                        If found_po = False Then
                            Dim cost As Decimal : cost = 0

                            If my_assemblies.Contains(part_s) = True Then
                                cost = myQuote.Cost_of_Assem(part_s)
                            Else
                                cost = Form1.Get_Latest_Cost(Login.Connection, part_s, vendor)
                            End If


                            my_grid.Rows.Add({part_s, desc, "", mfg, vendor, cost, mfg_type, 0, "", qty_box.Text * f_Qty, "", "", preff, part_Type})
                            my_grid.Rows(my_grid.Rows.Count - 2).DefaultCellStyle.BackColor = Color.Bisque 'color new edited row
                        End If

                        Call My_Material_r.recal_rev()

                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If

            Else
                MessageBox.Show("Please select a feature code and enter a Qty to add")
            End If
        Else
            MessageBox.Show("No Panel Selected")
        End If
    End Sub
End Class