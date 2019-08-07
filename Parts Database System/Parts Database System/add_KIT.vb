Imports MySql.Data.MySqlClient

Public Class add_KIT
    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If

    End Sub

    Private Sub add_KIT_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        '========= UNELEGANT WAY OF CLEANING CONTROLS ===========================

        Dim ctrl As Control
        For Each ctrl In Me.GroupBox1.Controls
            If (ctrl.GetType() Is GetType(TextBox)) Then
                ctrl.Text = ""
            End If
        Next
    End Sub

    Private Sub Button5_MouseEnter(sender As Object, e As EventArgs) Handles Button5.MouseEnter
        Button5.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Button5_MouseLeave(sender As Object, e As EventArgs) Handles Button5.MouseLeave
        Button5.BackColor = Color.DarkCyan
    End Sub


    Private Sub add_KIT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------------------------------- LOAD KIT ----------------------------------------
        Try
            Dim tabl As New DataTable
            Dim ad As New MySqlDataAdapter("Select Part_name from parts_table", Login.Connection)
            ad.Fill(tabl)
            kit_grid.DataSource = tabl

            'Setting Columns size 
            For i = 0 To kit_grid.ColumnCount - 1
                With kit_grid.Columns(i)
                    .Width = 690
                End With
            Next i

        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'look for an specific part
        Dim found_po As Boolean
        Dim rowindex As Integer
        For Each row As DataGridViewRow In kit_grid.Rows
            If String.Compare(row.Cells.Item("Part_name").Value.ToString, kit_part_grid.Text) = 0 Then
                rowindex = row.Index
                kit_grid.CurrentCell = kit_grid.Rows(rowindex).Cells(0)
                found_po = True
                Exit For
            End If
        Next
        If found_po = False Then
            MsgBox(" Part not found!")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '------------- ADD KIT -------------------

        '---------MAKE SURE TEXTBOXES ARE FILL UP ------------------------
        If String.Equals(kit_name.Text, "") = True Or String.Equals(akt_box.Text, "") = True Then
            MessageBox.Show("Please Fill Kit Name and AKN Number")
        Else

            '----------- MAKE SURE THE KIT DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                Dim kit_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Kit_Name", kit_name.Text)
                cmd.Parameters.AddWithValue("@AKN_Number", akt_box.Text)
                cmd.CommandText = "SELECT Kit_Name from kits where Kit_Name = @Kit_Name and AKN_Number = @AKN_Number"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader


                If Not reader.HasRows And akt_box.Text Like "AKT-*" Then

                    'IF IT DOES NOT EXIST THEN CREATE IT
                    reader.Close()

                    kit_cmd.Parameters.AddWithValue("@Kit_Name", kit_name.Text)
                    kit_cmd.Parameters.AddWithValue("@AKN_number", akt_box.Text)
                    kit_cmd.Parameters.AddWithValue("@Description", desc.Text)
                    kit_cmd.Parameters.AddWithValue("@Legacy_ADA", legacy.Text)

                    Create_cmd.Parameters.AddWithValue("@Kit_Name", kit_name.Text)
                    Create_cmd.Parameters.AddWithValue("@Description", desc.Text)
                    Create_cmd.Parameters.AddWithValue("@Legacy_ADA", legacy.Text)
                    Create_cmd.Parameters.AddWithValue("@Primary_v", p_vendor.Text)

                    kit_cmd.CommandText = "INSERT INTO kits(Kit_Name, AKN_Number, Description, Legacy_ADA_Number) VALUES (@Kit_Name, @AKN_Number, @Description, @Legacy_ADA)"
                    Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Legacy_ADA_Number, Primary_Vendor) VALUES (@Kit_Name,null,@Description, null,'Preferred', 'KIT',null,1,@Legacy_ADA,@Primary_v )"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()
                    kit_cmd.Connection = Login.Connection
                    kit_cmd.ExecuteNonQuery()

                    '------------------- Reload kit visualizer in form1 -----------------
                    'Dim table_v As New DataTable
                    'Dim adapter_v As New MySqlDataAdapter("(SELECT KIT_Name as DEVICE_KIT_Name,  Description from Kits) UNION ALL (SELECT DEVICE_Name as DEVICE_KIT_Name, Description from Devices)", Login.Connection)

                    'adapter_v.Fill(table_v)     'Visualizer 
                    'Form1.Device_Grid.DataSource = table_v
                    'Form1.Device_Grid.DataSource = table_v   'Visualizer table fill

                    ''Setting Columns size for Visualizer Datagrid
                    'For i = 0 To Form1.Device_Grid.ColumnCount - 1
                    '    With Form1.Device_Grid.Columns(i)
                    '        .Width = 370
                    '    End With
                    'Next i
                    '--------------------------------------------------------------------------
                    MessageBox.Show("Kit " & kit_name.Text & " created succesfully")
                    sel_kit.Text = "Selected Kit: " & kit_name.Text





                Else
                    MessageBox.Show("Kit " & kit_name.Text & " already exist!, Also Kit Numbers should start with AKT-")
                    'DO NOT FORGET TO CLOSE THE READER
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '------------------- ADD COMPONENTS TO THE KIT ---------------------
        Dim component As String : component = ""
        component = kit_grid.CurrentCell.Value.ToString

        If String.Equals(kit_name.Text, "") = False Or String.Equals(akt_box.Text, "") = False Then

            If String.Equals(component, "") = False And IsNumeric(qty_box.Text) = True Then

                If String.Equals(component, "") = False Then

                    If CType(qty_box.Text, Double) > 0 Then

                        Try

                            Dim Create_cmd As New MySqlCommand

                            Create_cmd.Parameters.AddWithValue("@Part_Name", component)
                            Create_cmd.Parameters.AddWithValue("@AKN_Number", akt_box.Text)
                            Create_cmd.CommandText = "SELECT Part_Name from akn  where Part_Name = @Part_Name and AKN_Number = @AKN_Number"
                            Create_cmd.Connection = Login.Connection
                            Dim reader As MySqlDataReader
                            reader = Create_cmd.ExecuteReader

                            If Not reader.HasRows Then
                                'IF IT DOES NOT EXIST THEN ADD IT
                                reader.Close()
                                Dim cmd As New MySqlCommand
                                cmd.Parameters.AddWithValue("@Part_Name", component)
                                cmd.Parameters.AddWithValue("@AKN_Number", akt_box.Text)
                                cmd.Parameters.AddWithValue("@LegacyADA", legacy.Text)
                                cmd.Parameters.AddWithValue("@qty", qty_box.Text)
                                cmd.CommandText = "INSERT INTO akn VALUES (@Part_Name, @AKN_Number, @qty, @LegacyADA);"
                                cmd.Connection = Login.Connection
                                cmd.ExecuteNonQuery()
                                MessageBox.Show("(" & qty_box.Text & ")" & " " & component & " have been added")

                            Else
                                MessageBox.Show(component & " was already added")

                            End If

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try
                    Else
                        MessageBox.Show("Qty must be greater than zero!")
                    End If

                Else
                    MessageBox.Show("Please enter Qty")
                End If

            Else
                MessageBox.Show("Invalid Wrong Quantity")
            End If
        Else
            MessageBox.Show("Kit not specified")
        End If

    End Sub
End Class