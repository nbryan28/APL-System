Imports MySql.Data.MySqlClient

Public Class SSLAW_m
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.BackColor = Color.Teal
    End Sub

    Private Sub Label41_MouseEnter(sender As Object, e As EventArgs) Handles Label41.MouseEnter
        Label41.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label41_MouseLeave(sender As Object, e As EventArgs) Handles Label41.MouseLeave
        Label41.BackColor = Color.Teal
    End Sub

    Private Sub Label1_MouseEnter(sender As Object, e As EventArgs) Handles Label1.MouseEnter
        Label1.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs) Handles Label1.MouseLeave
        Label1.BackColor = Color.Teal
    End Sub


    Private Sub SSLAW_m_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        field_c.Text = "Tracking"
        date_field.Text = "Date_Received"
        Try

            '-------- load feature code tables -----------
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT * from parts.SSLAW_RMA", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            rma_grid.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To rma_grid.ColumnCount - 1
                With rma_grid.Columns(i)
                    .Width = 190
                End With
            Next i

            rma_grid.Columns(18).Width = 700

            '--- my parts --
            Dim table_dev As New DataTable
            Dim adapter_dev As New MySqlDataAdapter("SELECT part_name from  SSLAW_parts", Login.Connection)

            adapter_dev.Fill(table_dev)
            allParts.DataSource = table_dev
            allParts.DataSource = table_dev

            allParts.Columns(0).Width = 700

            '------ my sslaw
            Dim table_dev2 As New DataTable
            Dim adapter_dev2 As New MySqlDataAdapter("SELECT * from  SSLAW_parts", Login.Connection)

            adapter_dev2.Fill(table_dev2)
            sslaw_g.DataSource = table_dev2
            sslaw_g.DataSource = table_dev2

            sslaw_g.Columns(0).Width = 1500
            sslaw_g.Columns(1).Width = 200
            sslaw_g.Columns(2).Width = 200

            '------ my sslaw systems
            Dim table_dev3 As New DataTable
            Dim adapter_dev3 As New MySqlDataAdapter("SELECT * from  Systems_SSLAW", Login.Connection)

            adapter_dev3.Fill(table_dev3)
            system_grid.DataSource = table_dev3
            system_grid.DataSource = table_dev3

            system_grid.Columns(0).Width = 1500

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        TabControl1.SelectedTab = tab_one
    End Sub


    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click
        TabControl1.SelectedTab = tab_three
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        TabControl1.SelectedTab = tab_four
    End Sub

    Private Sub rma_grid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles rma_grid.CellDoubleClick
        If (rma_grid.CurrentCell.ColumnIndex = 11) Then

            Dim z_number As String : z_number = "1z"
            If IsDBNull(rma_grid.CurrentCell.Value) = False Then
                z_number = rma_grid.CurrentCell.Value
                If String.IsNullOrEmpty(rma_grid.CurrentCell.Value) = False And String.Equals(rma_grid.CurrentCell.Value, "") = False And String.Equals(rma_grid.CurrentCell.Value, "") = False Then
                    Process.Start("https://wwwapps.ups.com/WebTracking/track?track=yes&trackNums=" & z_number)
                End If

            End If

        End If

        If (rma_grid.CurrentCell.ColumnIndex = 0) Then

            Dim rma_id As String : rma_id = ""
            If IsDBNull(rma_grid.CurrentCell.Value) = False Then
                rma_id = rma_grid.CurrentCell.Value
                If String.IsNullOrEmpty(rma_grid.CurrentCell.Value) = False And String.Equals(rma_grid.CurrentCell.Value, "") = False And String.Equals(rma_grid.CurrentCell.Value, "") = False Then
                    Call Show_sslaw(rma_id)  'show table with parts needed
                End If

            End If

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        rma2_grid.Rows.Clear()

        Dim ctrl As Control
        For Each ctrl In Me.GroupBox2.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next

        rma2_grid.Rows.Clear()

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@RMA_id", rma_id.Text)
            cmd.CommandText = "SELECT * from parts.SSLAW_RMA where RMA_id = @RMA_id"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    build_box.Text = reader(2).ToString
                    part_box.Text = reader(3).ToString
                    ups_rma_box.Text = reader(4).ToString
                    atronix_rma_box.Text = reader(5).ToString
                    sin_box.Text = reader(6).ToString
                    sout_box.Text = reader(7).ToString
                    date_r.Text = If(String.Equals(reader(8).ToString, ""), "1/1/1900 6:31 PM", reader(8).ToString)
                    date_s.Text = If(String.Equals(reader(9).ToString, ""), "1/1/1900 6:31 PM", reader(9).ToString)
                    repair_d.Text = If(String.Equals(reader(10).ToString, ""), "1/1/1900 6:31 PM", reader(10).ToString)
                    tracking.Text = reader(11).ToString
                    upspo.Text = reader(12).ToString
                    atc.Text = reader(13).ToString
                    repair_c.Text = reader(14).ToString
                    inv_n.Text = reader(15).ToString
                    inv_date.Text = If(String.Equals(reader(16).ToString, ""), "1/1/1900 6:31 PM", reader(16).ToString)
                    paid_d.Text = If(String.Equals(reader(17).ToString, ""), "1/1/1900 6:31 PM", reader(17).ToString)
                    notes_t.Text = reader(18).ToString
                    name_box.Text = reader(19).ToString
                    email_box.Text = reader(20).ToString
                    phone_box.Text = reader(21).ToString

                End While

            End If

            reader.Close()

            Dim cmd_po As New MySqlCommand
            cmd_po.Parameters.AddWithValue("@rma", rma_id.Text)
            cmd_po.CommandText = "SELECT part_name, qty, labor_cost, material_cost, total, final_t from RMA_comp where RMA_id = @rma"

            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    rma2_grid.Rows.Add()  'add a new row
                    rma2_grid.Rows(i).Cells(0).Value = readerv(0).ToString
                    rma2_grid.Rows(i).Cells(1).Value = readerv(1).ToString
                    rma2_grid.Rows(i).Cells(2).Value = readerv(2).ToString
                    rma2_grid.Rows(i).Cells(3).Value = readerv(3).ToString
                    rma2_grid.Rows(i).Cells(4).Value = readerv(4).ToString
                    rma2_grid.Rows(i).Cells(5).Value = readerv(4).ToString
                    rma2_grid.Rows(i).Cells(6).Value = readerv(5).ToString

                    i = i + 1
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try

            '-------- load feature code tables -----------
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT * from parts.SSLAW_RMA", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            rma_grid.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To rma_grid.ColumnCount - 1
                With rma_grid.Columns(i)
                    .Width = 190
                End With
            Next i

            rma_grid.Columns(18).Width = 700

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' ======== SEARCH PART NAME ===================

        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer
        Dim field As String : field = ""
        field = field_c.Text


        If Match_b.Checked = True Then


            For Each row As DataGridViewRow In rma_grid.Rows
                If String.Compare(row.Cells.Item(field).Value.ToString, TextBox14.Text) = 0 Then
                    rowindex = row.Index
                    rma_grid.CurrentCell = rma_grid.Rows(rowindex).Cells(0)
                    found_po = True
                    Exit For
                End If
            Next
            If found_po = False Then
                MsgBox("RMA not found!")
            End If

            Else

            '============== Partial match ============================
            Try
                Dim cmdstr As String : cmdstr = "SELECT * from parts.SSLAW_RMA where " & field & " LIKE  @search"
                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                cmd.Parameters.AddWithValue("@search", "%" & TextBox14.Text & "%")

                Dim table_s As New DataTable
                Dim adapter_s As New MySqlDataAdapter(cmd)
                adapter_s.Fill(table_s)
                rma_grid.DataSource = table_s

                For i = 0 To rma_grid.ColumnCount - 1
                    With rma_grid.Columns(i)
                        .Width = 190
                    End With
                Next i

                rma_grid.Columns(18).Width = 700

            Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If

    End Sub

    Private Sub DateTimePicker6_ValueChanged(sender As Object, e As EventArgs) Handles from_p.ValueChanged
        Try
            Dim cmdstr As String : cmdstr = "SELECT * from parts.SSLAW_RMA where " & date_field.Text & " >= @from and " & date_field.Text & " <= @to"
            Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
            cmd.Parameters.AddWithValue("@from", from_p.Value)
            cmd.Parameters.AddWithValue("@to", to_p.Value)

            Dim table_s As New DataTable
            Dim adapter_s As New MySqlDataAdapter(cmd)
            adapter_s.Fill(table_s)
            rma_grid.DataSource = table_s

            For i = 0 To rma_grid.ColumnCount - 1
                With rma_grid.Columns(i)
                    .Width = 190
                End With
            Next i

            rma_grid.Columns(18).Width = 700

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub DateTimePicker7_ValueChanged(sender As Object, e As EventArgs) Handles to_p.ValueChanged
        Try
            Dim cmdstr As String : cmdstr = "SELECT * from parts.SSLAW_RMA where " & date_field.Text & " >= @from and " & date_field.Text & " <= @to"
            Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
            cmd.Parameters.AddWithValue("@from", from_p.Value)
            cmd.Parameters.AddWithValue("@to", to_p.Value)

            Dim table_s As New DataTable
            Dim adapter_s As New MySqlDataAdapter(cmd)
            adapter_s.Fill(table_s)
            rma_grid.DataSource = table_s

            For i = 0 To rma_grid.ColumnCount - 1
                With rma_grid.Columns(i)
                    .Width = 190
                End With
            Next i

            rma_grid.Columns(18).Width = 700

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub rma_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles rma_grid.RowsAdded
        rma_f.Text = "RMA found: " & rma_grid.Rows.Count.ToString
    End Sub

    Private Sub rma_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles rma_grid.RowsRemoved
        rma_f.Text = "RMA found: " & rma_grid.Rows.Count.ToString
    End Sub

    Private Sub PR_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles rma2_grid.CellValueChanged

        '  Dim dlgR As DialogResult
        '  dlgR = MessageBox.Show("Do you want to overwrite this value?", "Attention!", MessageBoxButtons.YesNo)

        '  If dlgR = DialogResult.Yes Then
        Call Total_rows()
        '  End If

    End Sub

    Sub Total_rows()

        'Calculate total row number


        Dim total As Double : total = 0


        For i = 0 To rma2_grid.Rows.Count - 1
            If rma2_grid.Rows(i).IsNewRow Then Continue For

            If (IsNumeric(rma2_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(rma2_grid.Rows(i).Cells(5).Value())) Then

                rma2_grid.Rows(i).Cells(1).Value() = If(String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(1).Value()) = True Or IsNumeric(rma2_grid.Rows(i).Cells(1).Value()) = False, 0, rma2_grid.Rows(i).Cells(1).Value())
                rma2_grid.Rows(i).Cells(5).Value() = If(String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(5).Value()) = True Or IsNumeric(rma2_grid.Rows(i).Cells(5).Value()) = False, 0, rma2_grid.Rows(i).Cells(5).Value())

                rma2_grid.Rows(i).Cells(6).Value() = rma2_grid.Rows(i).Cells(1).Value() * rma2_grid.Rows(i).Cells(5).Value()
            End If

        Next

        For i = 0 To rma2_grid.Rows.Count - 1
            total = total + If(IsNumeric(rma2_grid.Rows(i).Cells(6).Value()), rma2_grid.Rows(i).Cells(6).Value(), 0)
        Next

        total_c.Text = "Total Cost: " & total
        repair_c.Text = total


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If (allParts.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(allParts.CurrentCell.Value.ToString, "") = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString)
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ add components to adv table ----------------------------------
                        Try
                            Dim cmd As New MySqlCommand
                            cmd.Parameters.AddWithValue("@part", allParts.CurrentCell.Value.ToString)
                            cmd.CommandText = "SELECT labor_cost, material_cost, total_cost from parts.SSLAW_parts where part_name = @part"
                            cmd.Connection = Login.Connection
                            Dim reader As MySqlDataReader
                            reader = cmd.ExecuteReader

                            If reader.HasRows Then
                                While reader.Read
                                    rma2_grid.Rows.Add(New String() {allParts.CurrentCell.Value.ToString, qty_change, reader(0).ToString, reader(1).ToString, reader(2), reader(2), (reader(2) * qty_change)})
                                End While
                            End If

                            reader.Close()
                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)

                        End Try

                        '-----------------------------------------------------
                    Else
                        MsgBox("Please input an integer.")
                    End If
                Catch
                    MsgBox("Please input a whole number.")
                End Try

            End If

        End If
    End Sub

    Private Sub rma2_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles rma2_grid.RowsAdded
        Call Total_rows()
    End Sub

    Function date_convert(date_po As Date) As Date
        date_convert = date_po
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '-------------------------- ADD RMA RECORD ---------------------
        If String.IsNullOrEmpty(part_box.Text) = True Or String.Equals(ups_rma_box.Text, "") = True Then
            MessageBox.Show("Please Fill part name and UPS RMA field")
        Else
            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@UPS_RMA", ups_rma_box.Text)
                cmd4.CommandText = "SELECT * from SSLAW_RMA where UPS_RMA = @UPS_RMA"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If Not reader4.HasRows Then

                    reader4.Close()

                    Dim myrma As Double

                    Dim Create_cmd As New MySqlCommand

                    Create_cmd.Parameters.AddWithValue("@building", build_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Part_Name", part_box.Text)
                    Create_cmd.Parameters.AddWithValue("@UPS_RMA", ups_rma_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Atronix_RMA", atronix_rma_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Serial_Number_IN", sin_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Serial_Number_Out", sout_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Date_Received", date_convert(date_r.Text))
                    Create_cmd.Parameters.AddWithValue("@Date_Shipped", date_convert(date_s.Text))
                    Create_cmd.Parameters.AddWithValue("@Repaired_date", date_convert(repair_d.Text))
                    Create_cmd.Parameters.AddWithValue("@Tracking", tracking.Text)
                    Create_cmd.Parameters.AddWithValue("@UPS_PO", upspo.Text)
                    Create_cmd.Parameters.AddWithValue("@ATC_Case", atc.Text)
                    Create_cmd.Parameters.AddWithValue("@Repair_Cost", If(IsNumeric(repair_c.Text) = True, repair_c.Text, 0))
                    Create_cmd.Parameters.AddWithValue("@Invoice_number", inv_n.Text)
                    Create_cmd.Parameters.AddWithValue("@Invoice_Date", date_convert(inv_date.Text))
                    Create_cmd.Parameters.AddWithValue("@Paid_Date", date_convert(paid_d.Text))
                    Create_cmd.Parameters.AddWithValue("@repair_notes", notes_t.Text)
                    Create_cmd.Parameters.AddWithValue("@Contact", name_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Email", email_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Phone", phone_box.Text)


                    Create_cmd.CommandText = "INSERT INTO SSLAW_RMA(building_name, Part_Name, UPS_RMA, Atronix_RMA, Serial_Number_IN, Serial_Number_Out, Date_Received, Date_Shipped, Repaired_date, Tracking,UPS_PO,ATC_Case,Repair_Cost,Invoice_number,Invoice_Date,Paid_Date,repair_notes, Contact, Email, Phone) VALUES (@building, @Part_Name, @UPS_RMA, @Atronix_RMA, @Serial_Number_IN, @Serial_Number_Out, @Date_Received, @Date_Shipped, @Repaired_date, @Tracking, @UPS_PO, @ATC_Case, @Repair_Cost, @Invoice_number, @Invoice_Date, @Paid_Date, @repair_notes, @Contact, @Email,@Phone)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)


                    Dim cmd As New MySqlCommand
                    cmd.CommandText = "select RMA_id from SSLAW_RMA order by RMA_id desc limit 1"
                    cmd.Connection = Login.Connection
                    cmd.ExecuteNonQuery()
                    Dim reader As MySqlDataReader
                    reader = cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            myrma = reader(0).ToString
                        End While
                    End If

                    reader.Close()

                    For i = 0 To rma2_grid.Rows.Count - 1

                        If String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(0).Value) = False And String.Equals(rma2_grid.Rows(i).Cells(0).Value, "") = False Then
                            If String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(1).Value) = False And String.Equals(rma2_grid.Rows(i).Cells(1).Value, "") = False Then

                                Dim cmd2 As New MySqlCommand
                                cmd2.Parameters.Clear()
                                cmd2.Parameters.AddWithValue("@rma", myrma)
                                cmd2.Parameters.AddWithValue("@part_name", rma2_grid.Rows(i).Cells(0).Value)
                                cmd2.Parameters.AddWithValue("@qty", If(IsNumeric(rma2_grid.Rows(i).Cells(1).Value) = True, rma2_grid.Rows(i).Cells(1).Value, 0))
                                cmd2.Parameters.AddWithValue("@labor", If(IsNumeric(rma2_grid.Rows(i).Cells(2).Value) = True, rma2_grid.Rows(i).Cells(2).Value, 0))
                                cmd2.Parameters.AddWithValue("@material", If(IsNumeric(rma2_grid.Rows(i).Cells(3).Value) = True, rma2_grid.Rows(i).Cells(3).Value, 0))
                                cmd2.Parameters.AddWithValue("@total", If(IsNumeric(rma2_grid.Rows(i).Cells(5).Value) = True, rma2_grid.Rows(i).Cells(5).Value, 0))
                                cmd2.Parameters.AddWithValue("@ftotal", If(IsNumeric(rma2_grid.Rows(i).Cells(6).Value) = True, rma2_grid.Rows(i).Cells(6).Value, 0))

                                cmd2.CommandText = "INSERT INTO RMA_comp(RMA_id, part_name,qty,labor_cost,material_cost,total, final_t) VALUES (@rma,@part_name, @qty,@labor,@material,@total,@ftotal)"
                                cmd2.Connection = Login.Connection
                                cmd2.ExecuteNonQuery()

                            End If
                        End If
                    Next


                    MessageBox.Show("RMA added succesfully!")

                    For Each ctrl In Me.GroupBox2.Controls
                        If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                            ctrl.Text = ""
                        End If
                    Next

                    rma2_grid.Rows.Clear()

                Else
                    reader4.Close()
                    MessageBox.Show("RMA already exists!")
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub


    Sub Refresh_tables(c As MySqlConnection)

        '---------- REFRESH DATA IN GRID ----------------------
        Try
            'REFRESH PARTS TABLE 
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT * from parts.SSLAW_RMA", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            rma_grid.DataSource = table

            'Setting Columns size for Parts Datagrid
            For i = 0 To rma_grid.ColumnCount - 1
                With rma_grid.Columns(i)
                    .Width = 190
                End With
            Next i

            rma_grid.Columns(18).Width = 700


            'refresh sslaw parts

            Dim table_dev2 As New DataTable
            Dim adapter_dev2 As New MySqlDataAdapter("SELECT * from  SSLAW_parts", Login.Connection)

            adapter_dev2.Fill(table_dev2)
            sslaw_g.DataSource = table_dev2
            sslaw_g.DataSource = table_dev2

            sslaw_g.Columns(0).Width = 1500
            sslaw_g.Columns(1).Width = 200
            sslaw_g.Columns(2).Width = 200


        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-----------DELETE RMA--------------------

        If String.Equals(rma_id.Text, "") = True Then
            MessageBox.Show("Please Fill RMA id box")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim Create_vd As New MySqlCommand
                Dim Create_acn As New MySqlCommand
                cmd.Parameters.AddWithValue("@RMA", rma_id.Text)
                cmd.CommandText = "SELECT * from SSLAW_RMA where RMA_id = @RMA"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    Dim dlgR As DialogResult
                    dlgR = MessageBox.Show("Are you sure you want to delete this RMA?", "Attention!", MessageBoxButtons.YesNo)
                    If dlgR = DialogResult.Yes Then

                        reader.Close()

                        Create_cmd.Parameters.AddWithValue("@RMA", rma_id.Text)
                        Create_vd.Parameters.AddWithValue("@RMA", rma_id.Text)

                        Create_cmd.CommandText = "DELETE FROM SSLAW_RMA where RMA_id = @RMA"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()
                        Create_vd.CommandText = "DELETE FROM RMA_comp where RMA_id = @RMA"
                        Create_vd.Connection = Login.Connection
                        Create_vd.ExecuteNonQuery()

                        MessageBox.Show("RMA deleted succesfully")
                        Call Refresh_tables(Login.Connection)

                        Dim ctrl As Control
                        For Each ctrl In Me.GroupBox2.Controls
                            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                                ctrl.Text = ""
                            End If
                        Next

                        rma2_grid.Rows.Clear()
                    Else
                        reader.Close()

                    End If
                Else
                    MessageBox.Show("We could not find the RMA specified to be removed")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '================================ UPDATE PART ================================
        If String.Equals(rma_id.Text, "") = True And String.IsNullOrEmpty(part_box.Text) = False Then
            MessageBox.Show("Please Fill RMA id box and Part name")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim Create_vd As New MySqlCommand

                cmd.Parameters.AddWithValue("@RMA", rma_id.Text)
                cmd.CommandText = "SELECT * from SSLAW_RMA where RMA_id = @RMA"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@RMA", rma_id.Text)
                    Create_cmd.Parameters.AddWithValue("@building", build_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Part_Name", If(String.IsNullOrEmpty(part_box.Text), "", part_box.Text))
                    Create_cmd.Parameters.AddWithValue("@UPS_RMA", ups_rma_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Atronix_RMA", atronix_rma_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Serial_Number_IN", sin_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Serial_Number_Out", sout_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Date_Received", date_convert(date_r.Text))
                    Create_cmd.Parameters.AddWithValue("@Date_Shipped", date_convert(date_s.Text))
                    Create_cmd.Parameters.AddWithValue("@Repaired_date", date_convert(repair_d.Text))
                    Create_cmd.Parameters.AddWithValue("@Tracking", tracking.Text)
                    Create_cmd.Parameters.AddWithValue("@UPS_PO", upspo.Text)
                    Create_cmd.Parameters.AddWithValue("@ATC_Case", atc.Text)
                    Create_cmd.Parameters.AddWithValue("@Repair_Cost", If(IsNumeric(repair_c.Text) = True, repair_c.Text, 0))
                    Create_cmd.Parameters.AddWithValue("@Invoice_number", inv_n.Text)
                    Create_cmd.Parameters.AddWithValue("@Invoice_Date", date_convert(inv_date.Text))
                    Create_cmd.Parameters.AddWithValue("@Paid_Date", date_convert(paid_d.Text))
                    Create_cmd.Parameters.AddWithValue("@repair_notes", notes_t.Text)
                    Create_cmd.Parameters.AddWithValue("@Contact", name_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Email", email_box.Text)
                    Create_cmd.Parameters.AddWithValue("@Phone", phone_box.Text)

                    Create_cmd.CommandText = "update SSLAW_RMA set building_name = @building, Part_Name = @Part_Name, UPS_RMA = @UPS_RMA, Atronix_RMA = @Atronix_RMA, Serial_Number_IN = @Serial_Number_IN, Serial_Number_Out = @Serial_Number_Out, Date_Received = @Date_Received, Date_Shipped = @Date_Shipped, Repaired_date = @Repaired_date, Tracking = @Tracking , UPS_PO = @UPS_PO, ATC_Case = @ATC_Case, Repair_Cost = @Repair_Cost, Invoice_number = @Invoice_number, Invoice_Date = @Invoice_Date, Paid_Date = @Paid_Date, repair_notes = @repair_notes, Contact = @Contact, Email = @Email, Phone = @Phone where RMA_id = @RMA"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    'remove previous data in rma_comp
                    Create_vd.Parameters.AddWithValue("@RMA", rma_id.Text)
                    Create_vd.CommandText = "DELETE FROM RMA_comp where RMA_id = @RMA"
                    Create_vd.Connection = Login.Connection
                    Create_vd.ExecuteNonQuery()


                    For i = 0 To rma2_grid.Rows.Count - 1

                        If String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(0).Value) = False And String.Equals(rma2_grid.Rows(i).Cells(0).Value, "") = False Then
                            If String.IsNullOrEmpty(rma2_grid.Rows(i).Cells(1).Value) = False And String.Equals(rma2_grid.Rows(i).Cells(1).Value, "") = False Then

                                Dim cmd2 As New MySqlCommand
                                cmd2.Parameters.Clear()
                                cmd2.Parameters.AddWithValue("@rma", rma_id.Text)
                                cmd2.Parameters.AddWithValue("@part_name", rma2_grid.Rows(i).Cells(0).Value)
                                cmd2.Parameters.AddWithValue("@qty", If(IsNumeric(rma2_grid.Rows(i).Cells(1).Value) = True, rma2_grid.Rows(i).Cells(1).Value, 0))
                                cmd2.Parameters.AddWithValue("@labor", If(IsNumeric(rma2_grid.Rows(i).Cells(2).Value) = True, rma2_grid.Rows(i).Cells(2).Value, 0))
                                cmd2.Parameters.AddWithValue("@material", If(IsNumeric(rma2_grid.Rows(i).Cells(3).Value) = True, rma2_grid.Rows(i).Cells(3).Value, 0))
                                cmd2.Parameters.AddWithValue("@total", If(IsNumeric(rma2_grid.Rows(i).Cells(5).Value) = True, rma2_grid.Rows(i).Cells(5).Value, 0))
                                cmd2.Parameters.AddWithValue("@ftotal", If(IsNumeric(rma2_grid.Rows(i).Cells(6).Value) = True, rma2_grid.Rows(i).Cells(6).Value, 0))

                                cmd2.CommandText = "INSERT INTO RMA_comp(RMA_id, part_name,qty,labor_cost,material_cost,total, final_t) VALUES (@rma,@part_name, @qty,@labor,@material,@total,@ftotal)"
                                cmd2.Connection = Login.Connection
                                cmd2.ExecuteNonQuery()

                            End If
                        End If
                    Next


                    MessageBox.Show("RMA updated succesfully")
                    Call Refresh_tables(Login.Connection)

                Else
                    MessageBox.Show("We could not find the RMA specified to be updated")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If

    End Sub

    Sub Show_sslaw(rma As Integer)

        SSLAW_parts.Visible = True

        Try

            Dim cmd_po As New MySqlCommand
            cmd_po.Parameters.AddWithValue("@rma", rma)
            cmd_po.CommandText = "SELECT part_name, qty, labor_cost, material_cost, total, final_t from RMA_comp where RMA_id = @rma"

            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    SSLAW_parts.need_grid.Rows.Add()  'add a new row
                    SSLAW_parts.need_grid.Rows(i).Cells(1).Value = readerv(0).ToString
                    SSLAW_parts.need_grid.Rows(i).Cells(2).Value = readerv(1).ToString
                    SSLAW_parts.need_grid.Rows(i).Cells(3).Value = readerv(2).ToString
                    SSLAW_parts.need_grid.Rows(i).Cells(4).Value = readerv(3).ToString
                    SSLAW_parts.need_grid.Rows(i).Cells(5).Value = readerv(4).ToString
                    SSLAW_parts.need_grid.Rows(i).Cells(6).Value = readerv(5).ToString

                    i = i + 1
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Dim f_cost As Double : f_cost = 0

        For j = 0 To SSLAW_parts.need_grid.Rows.Count - 1

            f_cost = f_cost + If(IsNumeric(SSLAW_parts.need_grid.Rows(j).Cells(6).Value) = True, SSLAW_parts.need_grid.Rows(j).Cells(6).Value, 0)

            Try
                If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & SSLAW_parts.need_grid.Rows(j).Cells(1).Value.ToString.Replace("/", "") & ".jpg") Is Nothing Then
                    Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & SSLAW_parts.need_grid.Rows(j).Cells(1).Value.ToString.Replace("/", "") & ".jpg")
                    SSLAW_parts.need_grid.Rows(j).Cells(0).Value = bmp1
                End If

            Catch ex As Exception
                '------------- default image -----------------
                Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\random_p.png")
                SSLAW_parts.need_grid.Rows(j).Cells(0).Value = bmp1
            End Try


        Next

        SSLAW_parts.total_c.Text = "Repair Cost: $" & f_cost

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim ctrl As Control
        For Each ctrl In Me.GroupBox2.Controls
            If (ctrl.GetType() Is GetType(TextBox) Or ctrl.GetType() Is GetType(ComboBox)) Then
                ctrl.Text = ""
            End If
        Next

        rma2_grid.Rows.Clear()
    End Sub

    Private Sub sslaw_g_DoubleClick(sender As Object, e As EventArgs) Handles sslaw_g.DoubleClick
        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW

        Dim part_name_t = sslaw_g.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_Name", part_name_t)
            cmd.CommandText = "SELECT * from SSLAW_parts where part_name = @Part_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    part_t.Text = reader(0).ToString
                    labor_t.Text = reader(1).ToString
                    mat_t.Text = reader(2).ToString
                    total_cost_box.Text = reader(3).ToString
                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try


    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Dim total_c As Double : total_c = 0

        '----------------------------------- ENTER A NEW PART FROM AN EXISTING PART------------------------------------------

        '---------MAKE SURE TEXTBOXES ARE FILL UP------------------------
        If String.Equals(part_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name field")
        Else

            'if total box is empty then total is sum of material and labor

            If String.Equals(total_cost_box.Text, "") Then
                total_c = If(IsNumeric(labor_t.Text), labor_t.Text, 0) + If(IsNumeric(mat_t.Text), mat_t.Text, 0)
            End If

            '----------- MAKE SURE THE PART DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                cmd.CommandText = "SELECT Part_Name from SSLAW_parts where part_Name = @Part_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If Not reader.HasRows Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                    Create_cmd.Parameters.AddWithValue("@labor", If(IsNumeric(labor_t.Text), labor_t.Text, 0))
                    Create_cmd.Parameters.AddWithValue("@material", If(IsNumeric(mat_t.Text), mat_t.Text, 0))
                    Create_cmd.Parameters.AddWithValue("@total", If(IsNumeric(total_c), total_c, 0))
                    Create_cmd.CommandText = "INSERT INTO SSLAW_parts(part_name, labor_cost, material_cost, total_cost) VALUES (@Part_Name, @labor, @material, @total)"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    Call Refresh_tables(Login.Connection)
                    MessageBox.Show("Part " & part_t.Text & " added succesfully!")

                Else
                    MessageBox.Show("Part " & part_t.Text & " already exist!")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If String.Equals(part_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name")
        Else
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                cmd.CommandText = "SELECT Part_Name from SSLAW_parts where part_Name = @Part_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    Dim dlgR As DialogResult
                    dlgR = MessageBox.Show("Are you sure you want to delete this Part?", "Attention!", MessageBoxButtons.YesNo)
                    If dlgR = DialogResult.Yes Then

                        reader.Close()

                        Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                        Create_cmd.CommandText = "DELETE FROM SSLAW_parts where part_Name = @Part_Name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()


                        MessageBox.Show("Part deleted succesfully")
                        Call Refresh_tables(Login.Connection)
                    Else
                        reader.Close()

                    End If
                Else
                    MessageBox.Show("We could not find the Part specified to be removed")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        '================================ UPDATE PART ================================
        Dim total_c As Double : total_c = 0
        Dim update_flag As Boolean : update_flag = True

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@part_name", part_t.Text)
            cmd.CommandText = "SELECT Part_name from SSLAW_parts where part_name = @part_name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader


            'MAKE SURE THE PART EXIST SO IT CAN BE UPDATED
            If Not reader.HasRows Then
                MessageBox.Show("Part does not exist! ... Update incomplete")
                update_flag = False
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '--------------------------------------------------------------------------------------------------------------------------------------

        If update_flag = True Then
            '-- Start updating the part ---------
            If String.Equals(total_cost_box.Text, "") Then
                total_c = If(IsNumeric(labor_t.Text), labor_t.Text, 0) + If(IsNumeric(mat_t.Text), mat_t.Text, 0)
            End If


            Try
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                Create_cmd.Parameters.AddWithValue("@labor", If(IsNumeric(labor_t.Text), labor_t.Text, 0))
                Create_cmd.Parameters.AddWithValue("@material", If(IsNumeric(mat_t.Text), mat_t.Text, 0))
                Create_cmd.Parameters.AddWithValue("@total", If(IsNumeric(total_c), total_c, 0))

                Create_cmd.CommandText = "UPDATE SSLAW_parts SET part_name = @Part_Name, labor_cost = @labor, material_cost = @material where part_Name = @Part_Name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                Call Refresh_tables(Login.Connection)
                MessageBox.Show("Part updated succesfully")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub system_grid_DoubleClick(sender As Object, e As EventArgs) Handles system_grid.DoubleClick
        Dim part_name_t = system_grid.CurrentCell.Value.ToString

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@System_Name", part_name_t)
            cmd.CommandText = "SELECT * from Systems_SSLAW where System_name = @System_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    TextBox4.Text = reader(0).ToString
                End While

            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try
    End Sub
End Class