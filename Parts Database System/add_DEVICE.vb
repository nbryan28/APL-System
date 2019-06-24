Imports MySql.Data.MySqlClient

Public Class add_DEVICE
    'Private Sub Button5_MouseEnter(sender As Object, e As EventArgs)
    '    Button5.BackColor = Color.Red
    'End Sub

    'Private Sub Button5_MouseLeave(sender As Object, e As EventArgs)
    '    Button5.BackColor = Color.Maroon
    'End Sub


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    'Private Sub add_DEVICE_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    '    '----------------------------------- LOAD COMPONENTS ----------------------------------------
    '    Try
    '        Dim tabl As New DataTable
    '        Dim ad As New MySqlDataAdapter("Select Part_name from parts_table", Login.Connection)
    '        ad.Fill(tabl)
    '        device_grid.DataSource = tabl

    '        'Setting Columns size 
    '        For i = 0 To device_grid.ColumnCount - 1
    '            With device_grid.Columns(i)
    '                .Width = 690
    '            End With
    '        Next i

    '    Catch ex As Exception
    '    End Try
    'End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    '    'Look for an specific part
    '    Dim found_po As Boolean
    '    Dim rowindex As Integer
    '    For Each row As DataGridViewRow In device_grid.Rows
    '        If String.Compare(row.Cells.Item("Part_name").Value.ToString, device_grid_box.Text) = 0 Then
    '            rowindex = row.Index
    '            device_grid.CurrentCell = device_grid.Rows(rowindex).Cells(0)
    '            found_po = True
    '            Exit For
    '        End If
    '    Next
    '    If found_po = False Then
    '        MsgBox(" Part not found!")
    '    End If
    'End Sub

    Private Sub add_DEVICE_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        '========= UNELEGANT WAY OF CLEANING CONTROLS ===========================

        Dim ctrl As Control
        For Each ctrl In Me.GroupBox1.Controls
            If (ctrl.GetType() Is GetType(TextBox)) Then
                ctrl.Text = ""
            End If
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '------------- ADD DEVICE -------------------

        '---------MAKE SURE TEXTBOXES ARE FILL UP ------------------------
        If String.Equals(advName.Text, "") = True Or String.Equals(legacy.Text, "") = True Then
            MessageBox.Show("Please Fill Assembly textbox")
        Else

            '----------- MAKE SURE THE DEVICE DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim device_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Device_Name", legacy.Text)
                cmd.Parameters.AddWithValue("@ADV_Number", advName.Text)
                cmd.CommandText = "SELECT Device_Name from devices where DEVICE_Name = @Device_Name and ADV_Number = @ADV_Number"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows And advName.Text Like "ADV-*" Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    device_cmd.Parameters.AddWithValue("@Device_Name", legacy.Text)
                    device_cmd.Parameters.AddWithValue("@ADV_number", advName.Text)
                    device_cmd.Parameters.AddWithValue("@Description", descr.Text)
                    device_cmd.Parameters.AddWithValue("@labor", If(IsNumeric(labor_c.Text) = False, 0, labor_c.Text))
                    device_cmd.Parameters.AddWithValue("@bulk", If(IsNumeric(bulk_c.Text) = False, 0, labor_c.Text))
                    device_cmd.Parameters.AddWithValue("@Legacy_ADA", legacy.Text)

                    Create_cmd.Parameters.AddWithValue("@Device_Name", legacy.Text)
                    Create_cmd.Parameters.AddWithValue("@Description", descr.Text)
                    Create_cmd.Parameters.AddWithValue("@Legacy_ADA", legacy.Text)


                    device_cmd.CommandText = "INSERT INTO devices(DEVICE_Name, ADV_Number, Description, Bulk_Cost, Labor_Cost, Legacy_ADA_Number) VALUES (@Device_Name, @ADV_Number, @Description, @bulk, @labor, @Legacy_ADA)"
                    Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Legacy_ADA_Number, Primary_Vendor, MFG_type) VALUES (@Device_Name,'Atronix',@Description, null,'Preferred', 'Assembly',null,1,@Legacy_ADA,'Atronix','Assembly' )"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    device_cmd.Connection = Login.Connection
                    device_cmd.ExecuteNonQuery()

                    '--- insert into inventory
                    Dim Create_cmdi As New MySqlCommand
                    Create_cmdi.Parameters.AddWithValue("@part_name", legacy.Text)
                    Create_cmdi.Parameters.AddWithValue("@description", descr.Text)
                    Create_cmdi.Parameters.AddWithValue("@manufacturer", "Atronix")
                    Create_cmdi.Parameters.AddWithValue("@MFG_type", "Assembly")


                    Create_cmdi.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, units, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, null,0,0,0,0)"
                    Create_cmdi.Connection = Login.Connection
                    Create_cmdi.ExecuteNonQuery()

                    '---------------------------


                    '--------------------------------------------------------------------------
                    MessageBox.Show("Device " & legacy.Text & " created succesfully")
                    ' device_sel.Text = "Selected Device: " & deviceName.Text
                    Call edit_KIT.Reload_visualizer()




                Else
                    MessageBox.Show("Device " & legacy.Text & " already exist!, Also Device Numbers should start with ADV-")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        ''------------------- ADD COMPONENTS TO THE DEVICE ---------------------
        'Dim component As String : component = ""
        'component = device_grid.CurrentCell.Value.ToString

        'If String.Equals(deviceName.Text, "") = False Or String.Equals(advName.Text, "") = False Then

        '    If String.Equals(component, "") = False And IsNumeric(qty_box.Text) = True Then

        '        If String.Equals(component, "") = False Then

        '            If CType(qty_box.Text, Double) > 0 Then

        '                Try

        '                    Dim Create_cmd As New MySqlCommand

        '                    Create_cmd.Parameters.AddWithValue("@Part_Name", component)
        '                    Create_cmd.Parameters.AddWithValue("@ADV_Number", advName.Text)
        '                    Create_cmd.CommandText = "SELECT Part_Name from adv  where Part_Name = @Part_Name and ADV_Number = @ADV_Number"
        '                    Create_cmd.Connection = Login.Connection
        '                    Dim reader As MySqlDataReader
        '                    reader = Create_cmd.ExecuteReader

        '                    If Not reader.HasRows Then
        '                        'IF IT DOES NOT EXIST THEN ADD IT
        '                        reader.Close()
        '                        Dim cmd As New MySqlCommand
        '                        cmd.Parameters.AddWithValue("@Part_Name", component)
        '                        cmd.Parameters.AddWithValue("@ADV_Number", advName.Text)
        '                        cmd.Parameters.AddWithValue("@LegacyADA", legacy.Text)
        '                        cmd.Parameters.AddWithValue("@qty", qty_box.Text)
        '                        cmd.CommandText = "INSERT INTO adv VALUES (@Part_Name, @ADV_Number, @qty, @LegacyADA);"
        '                        cmd.Connection = Login.Connection
        '                        cmd.ExecuteNonQuery()
        '                        MessageBox.Show("(" & qty_box.Text & ")" & " " & component & " have been added")
        '                    Else
        '                        MessageBox.Show(component & " was already added")

        '                    End If

        '                Catch ex As Exception
        '                    MessageBox.Show(ex.ToString)
        '                End Try
        '            Else
        '                MessageBox.Show("Qty must be greater than zero!")
        '            End If

        '        Else
        '            MessageBox.Show("Please enter Qty")
        '        End If

        '    Else
        '        MessageBox.Show("Invalid Wrong Quantity")
        '    End If
        'Else
        '    MessageBox.Show("Device not specified")
        'End If

    End Sub

    Private Sub legacy_TextChanged(sender As Object, e As EventArgs) Handles legacy.TextChanged

        advName.Text = "ADV-" & legacy.Text

    End Sub
End Class