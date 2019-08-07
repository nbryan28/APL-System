Imports MySql.Data.MySqlClient


Public Class Enter_Panel_info

    Public listn As List(Of String)
    Public listplc As List(Of String)
    Public listle As List(Of String)

    Public list_hp As List(Of String)
    Public list_type As List(Of String)

    Public nolist As List(Of String)

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If BOM_types.edit_panel = False Then

            If IsNothing(p_name1.SelectedItem) = False And IsNothing(p_name2.SelectedItem) = False And IsNothing(p_name3.SelectedItem) = False And String.IsNullOrEmpty(desc_box.Text) = False And String.IsNullOrEmpty(q_box.Text) = False Then

                If IsNumeric(q_box.Text) = True Then

                    '--- if vfd option has been selected
                    If options_v.Visible = True Then

                        Dim exist_c As Boolean : exist_c = False
                        Dim option_t As String : option_t = ""
                        Dim no_check As Boolean : no_check = True

                        For i = 0 To options_v.Items.Count - 1
                            If options_v.GetItemChecked(i) = True Then
                                no_check = False
                            End If
                        Next

                        If no_check = False Then
                            '--- determine the number of options ---
                            If options_v.GetItemChecked(0) = True Then
                                option_t = "0"

                            ElseIf options_v.GetItemChecked(1) Then
                                option_t = "1"

                            ElseIf options_v.GetItemChecked(2) Then
                                option_t = "2"

                            ElseIf options_v.GetItemChecked(3) Then
                                option_t = "3"

                            ElseIf options_v.GetItemChecked(4) Then
                                option_t = "4"

                            ElseIf options_v.GetItemChecked(5) Then
                                option_t = "5"

                            ElseIf options_v.GetItemChecked(6) Then
                                option_t = "6"

                            ElseIf options_v.GetItemChecked(7) Then
                                option_t = "7"

                            End If
                        Else
                            option_t = "0"
                        End If



                        '----------------------------------------

                        Dim check_cmd As New MySqlCommand
                            check_cmd.Parameters.AddWithValue("@mr_name", BOM_types.job_Selected & "_Panel_" & p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t)
                            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                            check_cmd.Connection = Login.Connection
                            check_cmd.ExecuteNonQuery()

                            Dim reader As MySqlDataReader
                            reader = check_cmd.ExecuteReader

                            If reader.HasRows Then
                                exist_c = True
                            End If

                            reader.Close()

                        If exist_c = False Then

                            BOM_types.Panel_grid.AllowUserToAddRows = True
                            BOM_types.Panel_grid.Rows.Clear()

                            BOM_types.temp_panel_name = p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t
                            BOM_types.temp_panel_desc = desc_box.Text
                            BOM_types.temp_panel_qty = q_box.Text

                            BOM_types.Panel_n_l.Text = "Name:  " & p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t
                            BOM_types.qty_l.Text = "Qty: " & q_box.Text

                            Me.Visible = False

                        Else
                            MessageBox.Show("There is already a Panel with that name!")

                        End If

                        '---- no vfd option selected
                    Else

                        Dim exist_c As Boolean : exist_c = False

                        Dim check_cmd As New MySqlCommand
                        check_cmd.Parameters.AddWithValue("@mr_name", BOM_types.job_Selected & "_Panel_" & p_name1.Text & "_" & p_name2.Text & p_name3.Text)
                        check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                        check_cmd.Connection = Login.Connection
                        check_cmd.ExecuteNonQuery()

                        Dim reader As MySqlDataReader
                        reader = check_cmd.ExecuteReader

                        If reader.HasRows Then
                            exist_c = True
                        End If

                        reader.Close()

                        '---------------------------
                        If exist_c = False Then

                            BOM_types.Panel_grid.AllowUserToAddRows = True
                            BOM_types.Panel_grid.Rows.Clear()

                            BOM_types.temp_panel_name = p_name1.Text & "_" & p_name2.Text & p_name3.Text
                            BOM_types.temp_panel_desc = desc_box.Text
                            BOM_types.temp_panel_qty = q_box.Text

                            BOM_types.Panel_n_l.Text = "Name:  " & p_name1.Text & "_" & p_name2.Text & p_name3.Text
                            BOM_types.qty_l.Text = "Qty: " & q_box.Text

                            Me.Visible = False

                        Else
                            MessageBox.Show("There is already a Panel with that name!")

                        End If

                    End If
                Else
                    MessageBox.Show("Incorrect Qty")
                End If

            Else
                MessageBox.Show("Please enter a Panel Name, description and Qty")
            End If

        Else

            If IsNothing(p_name1.SelectedItem) = False And IsNothing(p_name2.SelectedItem) = False And IsNothing(p_name3.SelectedItem) = False Then

                If IsNumeric(q_box.Text) = True Then

                    If options_v.Visible = True Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////
                        Try
                            Dim option_t As String : option_t = ""
                            Dim no_check As Boolean : no_check = True

                            For i = 0 To options_v.Items.Count - 1
                                If options_v.GetItemChecked(i) = True Then
                                    no_check = False
                                End If
                            Next

                            If no_check = False Then
                                '--- determine the number of options ---
                                '--- determine the number of options ---
                                If options_v.GetItemChecked(0) = True Then
                                    option_t = "0"

                                ElseIf options_v.GetItemChecked(1) Then
                                    option_t = "1"

                                ElseIf options_v.GetItemChecked(2) Then
                                    option_t = "2"

                                ElseIf options_v.GetItemChecked(3) Then
                                    option_t = "3"

                                ElseIf options_v.GetItemChecked(4) Then
                                    option_t = "4"

                                ElseIf options_v.GetItemChecked(5) Then
                                    option_t = "5"

                                ElseIf options_v.GetItemChecked(6) Then
                                    option_t = "6"

                                ElseIf options_v.GetItemChecked(7) Then
                                    option_t = "7"

                                End If
                            Else
                                option_t = "0"
                            End If

                            Dim exist_c As Boolean : exist_c = False
                            Dim last_mr_name As String = BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name
                            BOM_types.temp_panel_name = p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t
                            BOM_types.temp_panel_desc = desc_box.Text
                            BOM_types.temp_panel_qty = q_box.Text

                            BOM_types.Panel_n_l.Text = "Name:  " & p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t
                            BOM_types.qty_l.Text = "Qty: " & q_box.Text


                            '---------- check if new name is valid ----------------------

                            Dim check_cmd As New MySqlCommand
                            check_cmd.Parameters.AddWithValue("@mr_name", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                            check_cmd.Connection = Login.Connection
                            check_cmd.ExecuteNonQuery()

                            Dim reader As MySqlDataReader
                            reader = check_cmd.ExecuteReader

                            If reader.HasRows Then
                                exist_c = True
                            End If

                            reader.Close()

                            '---------------------------
                            If exist_c = False Then

                                Dim Create_cmd As New MySqlCommand
                                Create_cmd.Parameters.AddWithValue("@mr_name", last_mr_name)
                                Create_cmd.Parameters.AddWithValue("@mr_name_n", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                                Create_cmd.Parameters.AddWithValue("@Panel_name", BOM_types.temp_panel_name)
                                Create_cmd.Parameters.AddWithValue("@Panel_qty", BOM_types.temp_panel_qty)
                                Create_cmd.Parameters.AddWithValue("@Panel_desc", BOM_types.temp_panel_desc)

                                Create_cmd.CommandText = "UPDATE Material_Request.mr SET mr_name = @mr_name_n , Panel_name = @Panel_name, Panel_qty = @Panel_qty, Panel_desc = @Panel_desc  where mr_name = @mr_name"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()

                                '-- update mr_data
                                Dim Create_cmd2 As New MySqlCommand
                                Create_cmd2.Parameters.AddWithValue("@mr_name", last_mr_name)
                                Create_cmd2.Parameters.AddWithValue("@mr_name_n", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                                Create_cmd2.CommandText = "UPDATE Material_Request.mr_data SET mr_name = @mr_name_n where mr_name = @mr_name"
                                Create_cmd2.Connection = Login.Connection
                                Create_cmd2.ExecuteNonQuery()

                                Me.Visible = False

                            Else
                                MessageBox.Show("There is already a Panel with that name!")
                            End If

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try
                        '/////////////////////////////////////////////////////////////////////////////////
                    Else

                        '----/////////////////////// update panel bom //////////////////////////////////
                        Try

                            Dim exist_c As Boolean : exist_c = False

                            Dim last_mr_name As String = BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name
                            BOM_types.temp_panel_name = p_name1.Text & "_" & p_name2.Text & p_name3.Text
                            BOM_types.temp_panel_desc = desc_box.Text
                            BOM_types.temp_panel_qty = q_box.Text

                            BOM_types.Panel_n_l.Text = "Name:  " & p_name1.Text & "_" & p_name2.Text & p_name3.Text
                            BOM_types.qty_l.Text = "Qty: " & q_box.Text

                            '---------- check if new name is valid ----------------------

                            Dim check_cmd As New MySqlCommand
                            check_cmd.Parameters.AddWithValue("@mr_name", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                            check_cmd.Connection = Login.Connection
                            check_cmd.ExecuteNonQuery()

                            Dim reader As MySqlDataReader
                            reader = check_cmd.ExecuteReader

                            If reader.HasRows Then
                                exist_c = True
                            End If

                            reader.Close()

                            '---------------------------
                            If exist_c = False Then

                                Dim Create_cmd As New MySqlCommand
                                Create_cmd.Parameters.AddWithValue("@mr_name", last_mr_name)
                                Create_cmd.Parameters.AddWithValue("@mr_name_n", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                                Create_cmd.Parameters.AddWithValue("@Panel_name", BOM_types.temp_panel_name)
                                Create_cmd.Parameters.AddWithValue("@Panel_qty", BOM_types.temp_panel_qty)
                                Create_cmd.Parameters.AddWithValue("@Panel_desc", BOM_types.temp_panel_desc)

                                Create_cmd.CommandText = "UPDATE Material_Request.mr SET mr_name = @mr_name_n , Panel_name = @Panel_name, Panel_qty = @Panel_qty, Panel_desc = @Panel_desc  where mr_name = @mr_name"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()

                                '-- update mr_data
                                Dim Create_cmd2 As New MySqlCommand
                                Create_cmd2.Parameters.AddWithValue("@mr_name", last_mr_name)
                                Create_cmd2.Parameters.AddWithValue("@mr_name_n", BOM_types.job_Selected & "_Panel_" & BOM_types.temp_panel_name)
                                Create_cmd2.CommandText = "UPDATE Material_Request.mr_data SET mr_name = @mr_name_n where mr_name = @mr_name"
                                Create_cmd2.Connection = Login.Connection
                                Create_cmd2.ExecuteNonQuery()

                                Me.Visible = False

                            Else
                                MessageBox.Show("There is already a Panel with that name!")
                            End If

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try
                        '/////////////////////////////////////////////////////////////////////////

                    End If

                Else
                        MessageBox.Show("Incorrect Qty")
                End If

            Else
                MessageBox.Show("Please enter a Panel Name and Qty")

            End If
        End If
    End Sub

    Private Sub Enter_Panel_info_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub Enter_Panel_info_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        listn = New List(Of String)  'this contains numbers
        For i = 1 To 100
            listn.Add(i)
        Next

        listplc = New List(Of String)  'number and plc

        listplc.Add("PLC")
        For i = 1 To 100
            listplc.Add(i)
        Next

        listle = New List(Of String)  'add letters and numbers
        For Each c In "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
            listle.Add(c)
        Next


        list_hp = New List(Of String)  'hp values
        list_hp.Add("01")
        list_hp.Add("02")
        list_hp.Add("03")
        list_hp.Add("05")
        list_hp.Add("07")
        list_hp.Add("10")
        list_hp.Add("15")
        list_hp.Add("20")

        list_type = New List(Of String)  'type values current is just 1 
        list_type.Add("1")


        For i = 1 To 100
            listle.Add(i)
        Next

        nolist = New List(Of String)  '-use for cp
        nolist.Add("-")

        '=== set comboboxes =======
        p_name1.Text = "ADA"
        Call add_combo(p_name2, listplc)
        Call add_combo(p_name3, listle)
        '-------------------------



        If BOM_types.edit_panel = True Then
            desc_box.Text = BOM_types.temp_panel_desc
            q_box.Text = BOM_types.temp_panel_qty
            Me.Text = BOM_types.temp_panel_name
            Label1.Text = "Change Name To"
        Else

            desc_box.Text = ""
            q_box.Text = 1
            Me.Text = "Panel Info"
            Label1.Text = "Panel Name:"

        End If



        p_name1.ResetText()
        p_name2.ResetText()
        p_name3.ResetText()
    End Sub

    Sub add_combo(combo As ComboBox, mylist As List(Of String))

        combo.Items.Clear()
        '--fill a combobox with particularl ist
        For i = 0 To mylist.Count - 1
            combo.Items.Add(mylist.Item(i).ToString)
        Next

    End Sub

    Private Sub p_name1_SelectedValueChanged(sender As Object, e As EventArgs) Handles p_name1.SelectedValueChanged

        Dim var_p As String : var_p = "ADA"

        If Not p_name1.SelectedItem Is Nothing Then
            var_p = p_name1.SelectedItem.ToString
        Else
            var_p = "ADA"
        End If

        '----------- display vfd --------
        If String.Equals(var_p, "VFD") = True Then
            ' p_name4.Visible = True
            p_name3.Visible = True
            Call add_combo(p_name2, list_hp) '  Call add_combo(p_name2, listn)
            Call add_combo(p_name3, listn) '  Call add_combo(p_name3, listn)
            'hp_l.Visible = True
            options_v.Visible = True

        ElseIf String.Equals(var_p, "ADA") = True Then
            Call add_combo(p_name2, listplc)
            Call add_combo(p_name3, listle)
            '  p_name4.Visible = False
            p_name3.Visible = True
            '  hp_l.Visible = False
            options_v.Visible = False

        ElseIf String.Equals(var_p, "CP") = True Then
            Call add_combo(p_name2, listn)
            Call add_combo(p_name3, nolist)
            p_name3.Text = "-"
            p_name3.Visible = False
            ' p_name4.Visible = False
            '  hp_l.Visible = False
            options_v.Visible = False


        End If
        '-----------------------------------




    End Sub

    Private Sub options_v_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles options_v.ItemCheck
        For i = 0 To options_v.Items.Count - 1
            If (i <> e.Index) = True Then
                options_v.SetItemChecked(i, False)
            End If
        Next
    End Sub
End Class