Imports MySql.Data.MySqlClient

Public Class edit_KIT
    Private Sub edit_KIT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '----------------------------------- LOAD KIT ----------------------------------------
        Try

            'Fill first KIT table where the user can select the kit to be modified
            Dim tabl As New DataTable
            Dim ad As New MySqlDataAdapter("Select KIT_Name from kits", Login.Connection)
            ad.Fill(tabl)
            kit_sel.DataSource = tabl

            'Setting Columns size 
            For i = 0 To kit_sel.ColumnCount - 1
                With kit_sel.Columns(i)
                    .Width = 540
                End With
            Next i

            'Fill the all parta table where the user can select the parts to be added to the kit member
            Dim tabl2 As New DataTable
            Dim ad2 As New MySqlDataAdapter("Select Part_Name, Part_Description, Part_Status, Part_Type from parts_table", Login.Connection)
            ad2.Fill(tabl2)
            allParts.DataSource = tabl2

            'Setting Columns size 
            For i = 0 To allParts.ColumnCount - 1
                With allParts.Columns(i)
                    .Width = 340
                End With
            Next i

        Catch ex As Exception
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

    Private Sub Button6_MouseEnter(sender As Object, e As EventArgs) Handles Button6.MouseEnter
        Button6.BackColor = Color.DarkCyan
    End Sub

    Private Sub Button6_MouseLeave(sender As Object, e As EventArgs) Handles Button6.MouseLeave
        Button6.BackColor = Color.CadetBlue
    End Sub

    Private Sub Button5_MouseEnter(sender As Object, e As EventArgs) Handles Button5.MouseEnter
        Button5.BackColor = Color.Firebrick
    End Sub

    Private Sub Button5_MouseLeave(sender As Object, e As EventArgs) Handles Button5.MouseLeave
        Button5.BackColor = Color.Maroon
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'look for an specific part
        Dim found_po As Boolean
        Dim rowindex As Integer
        For Each row As DataGridViewRow In kit_sel.Rows
            If String.Compare(row.Cells.Item("Kit_Name").Value.ToString, kit_select_box.Text) = 0 Then
                rowindex = row.Index
                kit_sel.CurrentCell = kit_sel.Rows(rowindex).Cells(0)
                found_po = True
                Exit For
            End If
        Next
        If found_po = False Then
            MsgBox(" Kit not found!")
        End If
    End Sub



    Private Sub kit_sel_CellContentDoubleClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles kit_sel.CellContentDoubleClick
        'SHOW DATA IN THE TEXTBOXES WHEN DOUBLE CLICKED IN THE DATAGRIDVIEW

        '----------------- I am too lazy to write a big query so im going to do it in three parts, kit, vendor and acn --------------------

        Dim name_kit = kit_sel.CurrentCell.Value.ToString


        Try
            '------------- get AKT values -----------------------
            Dim cmd As New MySqlCommand
            Dim cmd_p As New MySqlCommand
            Dim cmd_a As New MySqlCommand
            Dim reader As MySqlDataReader
            ' Dim reader_p As MySqlDataReader


            cmd.Parameters.AddWithValue("@Kit_Name", name_kit)
            cmd.CommandText = "SELECT * from kits where KIT_Name = @Kit_Name"

            cmd_p.CommandText = "SELECT Primary_Vendor from parts_table where Part_Name = @Kit_Name"
            cmd_p.Parameters.AddWithValue("@Kit_Name", name_kit)




            cmd.Connection = Login.Connection
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    kitname.Text = reader(0).ToString
                    aktname.Text = reader(1).ToString
                    desc.Text = reader(2).ToString
                    legacy.Text = reader(3).ToString
                End While
            End If

            reader.Close()

            '-------------------- vendor kit---------------- 
            'cmd_p.Connection = Login.Connection
            'reader_p = cmd_p.ExecuteReader
            'If reader_p.HasRows Then
            '    While reader_p.Read
            '        pvendor.Text = reader_p(0).ToString
            '    End While
            'End If
            'reader_p.Close()



            '--------------- Fill up kit members ---------------
            Call Populate_kit()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'delete Kit data in parts, akn, vendor, acn and Kits table

        If String.Equals(kitname.Text, "") = True Or String.Equals(aktname.Text, "") = True Then
            MessageBox.Show("Please Fill Kit Name and AKT Number")
        Else
            Try
                Dim cmd As New MySqlCommand 'use this mysql command to check if the kit exist
                Dim Create_cmd As New MySqlCommand 'parts command
                Dim Create_vd As New MySqlCommand  'vendor command             
                Dim Create_akn As New MySqlCommand  'akn command
                Dim Create_kits As New MySqlCommand  'kit command
                Dim Create_adv As New MySqlCommand  'dev command

                cmd.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                cmd.Parameters.AddWithValue("@akt", aktname.Text)
                cmd.CommandText = "SELECT Kit_Name from kits where Kit_Name = @Kit_Name and AKN_Number = @akt"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then

                    Dim dlgR As DialogResult
                    dlgR = MessageBox.Show("Are you sure you want to delete this Kit?", "Attention!", MessageBoxButtons.YesNo)
                    If dlgR = DialogResult.Yes Then

                        reader.Close()

                        Create_cmd.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                        Create_vd.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                        Create_adv.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                        Create_kits.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                        Create_akn.Parameters.AddWithValue("@akt", aktname.Text)

                        Create_cmd.CommandText = "DELETE FROM parts_table where Part_Name = @Kit_Name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()
                        Create_vd.CommandText = "DELETE FROM vendors_table where Part_Name = @Kit_Name"
                        Create_vd.Connection = Login.Connection
                        Create_vd.ExecuteNonQuery()
                        Create_adv.CommandText = "DELETE FROM adv where Part_Name = @Kit_Name"
                        Create_adv.Connection = Login.Connection
                        Create_adv.ExecuteNonQuery()
                        Create_kits.CommandText = "DELETE FROM kits where Kit_Name = @Kit_Name"
                        Create_kits.Connection = Login.Connection
                        Create_kits.ExecuteNonQuery()
                        Create_akn.CommandText = "DELETE FROM akn where AKN_Number = @akt"
                        Create_akn.Connection = Login.Connection
                        Create_akn.ExecuteNonQuery()

                        Call Reload_visualizer()
                        '-------- refresh the kit datagrid in first tab ----------------
                        Dim tabl As New DataTable
                        Dim ad As New MySqlDataAdapter("Select KIT_Name from kits", Login.Connection)
                        ad.Fill(tabl)
                        kit_sel.DataSource = tabl

                        'Setting Columns size 
                        For i = 0 To kit_sel.ColumnCount - 1
                            With kit_sel.Columns(i)
                                .Width = 540
                            End With
                        Next i
                        '------------------------------------------------------------------
                        MessageBox.Show("Kit deleted succesfully")

                    Else
                        reader.Close()

                    End If
                Else
                    MessageBox.Show("We could not find the Kit specified to be removed")
                    reader.Close()
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '------------------------ Remove a member of the kit -------------------------------------------

        If (current_p.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(current_p.CurrentCell.Value.ToString, "") = False Then

                Try
                    Dim cmd As New MySqlCommand
                    cmd.Parameters.AddWithValue("@Part_Name", current_p.CurrentCell.Value.ToString)
                    cmd.Parameters.AddWithValue("@ada", legacy.Text)
                    cmd.CommandText = "DELETE FROM akn where Part_Name = @Part_Name and Legacy_ADA = @ada"
                    cmd.Connection = Login.Connection
                    cmd.ExecuteNonQuery()

                    Dim cmdstr As String : cmdstr = "Select Part_Name, Qty from akn where AKN_Number = @akn"
                    Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
                    cmd_k.Parameters.AddWithValue("@akn", aktname.Text)

                    Dim table_s As New DataTable
                    Dim ad As New MySqlDataAdapter(cmd_k)
                    ad.Fill(table_s)
                    current_p.DataSource = table_s

                    'Setting Columns size 
                    For i = 0 To current_p.ColumnCount - 1
                        With current_p.Columns(i)
                            .Width = 280
                        End With
                    Next i

                    MessageBox.Show("Part removed")
                    Call Reload_visualizer()


                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If

        End If
    End Sub

    Sub Reload_visualizer()

        'Try
        '    '------------------- Reload kit visualizer in form1 -----------------
        '    Dim table_v As New DataTable
        '    Dim adapter_v As New MySqlDataAdapter("(SELECT Legacy_ADA_Number as Assembly_KIT,  Description from Kits) UNION ALL (SELECT Legacy_ADA_Number as Assembly_KIT, Description from Devices)", Login.Connection)

        '    adapter_v.Fill(table_v)     'Visualizer 
        '    Form1.Device_Grid.DataSource = table_v
        '    Form1.Device_Grid.DataSource = table_v   'Visualizer table fill

        '    'Setting Columns size for Visualizer Datagrid
        '    For i = 0 To Form1.Device_Grid.ColumnCount - 1
        '        With Form1.Device_Grid.Columns(i)
        '            .Width = 570
        '        End With
        '    Next i
        '    '-----------------------------------------------------------------------
        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '---------------------------------------- CHANGE QTY OF KIT MEMBER ----------------------

        If (current_p.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(current_p.CurrentCell.Value.ToString, "") = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please input the new quantity.")
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ update qty ----------------------------------
                        Try
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@Part_Name", current_p.CurrentCell.Value.ToString)
                            Create_cmd.Parameters.AddWithValue("@ada", legacy.Text)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.CommandText = "UPDATE akn  SET  Qty = @qty where Part_Name = @Part_Name and Legacy_ADA = @ada"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Call Reload_visualizer()
                            Call Populate_kit()

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

    Sub Populate_kit()

        'Populate the data grid of the kit members
        Try
            '--------------- populate kit members -----------------

            Dim cmdstr As String : cmdstr = "Select Part_Name, Qty from akn where AKN_Number = @akn"
            Dim cmd_k As New MySqlCommand(cmdstr, Login.Connection)
            cmd_k.Parameters.AddWithValue("@akn", aktname.Text)

            Dim table_s As New DataTable
            Dim ad As New MySqlDataAdapter(cmd_k)
            ad.Fill(table_s)
            current_p.DataSource = table_s

            'Setting Columns size 
            For i = 0 To current_p.ColumnCount - 1
                With current_p.Columns(i)
                    .Width = 280
                End With
            Next i

        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '--------------- Update KIT --------------------
        If String.Equals(kitname.Text, "") = False Or String.Equals(aktname.Text, "") = False Then

            Try
                Dim Create_cmd As New MySqlCommand
                Dim Create_acn As New MySqlCommand
                Dim Create_part As New MySqlCommand
                Dim Create_akn As New MySqlCommand
                ' -------------- kit table par ------------------------
                Create_cmd.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                Create_cmd.Parameters.AddWithValue("@AKN_Number", aktname.Text)
                Create_cmd.Parameters.AddWithValue("@Description", desc.Text)
                Create_cmd.Parameters.AddWithValue("@legacy", legacy.Text)

                '------------- akn table par --------------------------
                Create_akn.Parameters.AddWithValue("@legacy", legacy.Text)
                Create_akn.Parameters.AddWithValue("@AKN_Number", aktname.Text)
                '---------------- parts table par --------------------
                '      Create_part.Parameters.AddWithValue("@primary_vendor", pvendor.Text)
                Create_part.Parameters.AddWithValue("@Kit_Name", kitname.Text)
                Create_part.Parameters.AddWithValue("@Description", desc.Text)
                Create_part.Parameters.AddWithValue("@legacy", legacy.Text)


                '-------- update kits
                Create_cmd.CommandText = "UPDATE kits  SET  Description = @Description, Legacy_ADA_Number = @legacy  where KIT_Name = @Kit_Name and AKN_Number = @AKN_Number"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()
                '-----------------------------------------------

                '--------------------------------------------------
                '--------- update akn
                Create_akn.CommandText = "UPDATE akn  SET  Legacy_ADA = @legacy  where AKN_Number = @AKN_Number"
                Create_akn.Connection = Login.Connection
                Create_akn.ExecuteNonQuery()
                '---------------------------------------------------
                '------- part
                Create_part.CommandText = "UPDATE parts_table SET Part_Description = @Description, Primary_Vendor = @primary_vendor, Legacy_ADA_Number = @legacy where Part_Name = @Kit_Name"
                Create_part.Connection = Login.Connection
                Create_part.ExecuteNonQuery()


                Call Reload_visualizer()
                MessageBox.Show("Kit updated succesfully")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Else
            MessageBox.Show("Please Fill Kit Name and AKT Number")
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        '-------- Add components to the Device ----------------
        If (allParts.CurrentCell.ColumnIndex = 0) Then

            If String.Equals(allParts.CurrentCell.Value.ToString, "") = False And String.Equals(aktname.Text, "") = False Then

                Dim qty_change As Integer
                Dim qty_test As Integer

                Try
                    qty_change = InputBox("Please enter the number of " & allParts.CurrentCell.Value.ToString & " this kit need")
                    If Integer.TryParse(qty_change, qty_test) Then
                        '------------------------ add components to adv table ----------------------------------
                        Try
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@Part_Name", allParts.CurrentCell.Value.ToString)
                            Create_cmd.Parameters.AddWithValue("@qty", qty_change)
                            Create_cmd.Parameters.AddWithValue("@akn", aktname.Text)
                            Create_cmd.Parameters.AddWithValue("@legacy", legacy.Text)
                            Create_cmd.CommandText = "INSERT INTO akn VALUES (@Part_Name, @akn, @qty, @legacy)"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()
                            Call Populate_kit()
                            MessageBox.Show(allParts.CurrentCell.Value.ToString & " added succesfully")

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
End Class