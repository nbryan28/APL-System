Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Edit_BOM_Panels
    Public my_job As String
    Public mbom As String
    Public field_bom As String
    Public need_by_date As String
    Public shipping_ad As String

    Public Smtp_Server As New SmtpClient

    Public listn As List(Of String)
    Public listplc As List(Of String)
    Public listle As List(Of String)
    Public nolist As List(Of String)

    Public list_hp As List(Of String)
    Public list_type As List(Of String)

    Public inp_mr_name As String
    Public inp_rev_name As String


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("Panel")
        Part_Picker1.specify(Panel_grid)
    End Sub

    Private Sub Edit_BOM_Panels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call EnableDoubleBuffered(Panel_grid)
        Call EnableDoubleBuffered(progress_grid)

        TabControl1.TabPages.Remove(TabPage2) 'remove cost tab

        '--- load comboboxes content ----
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

        For i = 1 To 100
            listle.Add(i)
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

        nolist = New List(Of String)  '-use for cp
        nolist.Add("-")

        '=== set comboboxes =======
        p_name1.Text = "ADA"
        Call add_combo(p_name2, listplc)
        Call add_combo(p_name3, listle)


        '--------------------------------


        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"


        Panel_grid.Rows.Clear()
        progress_grid.Rows.Clear()

        my_job = My_Material_r.open_job
        mbom = ""
        field_bom = ""


        Try
            ComboBox1.Items.Clear()

            Dim cmd_j As New MySqlCommand
            cmd_j.Parameters.AddWithValue("@job", my_job)
            cmd_j.CommandText = "SELECT distinct Panel_name from Material_Request.mr where BOM_type = 'Panel' and job = @job"
            cmd_j.Connection = Login.Connection

            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    ComboBox1.Items.Add(readerj(0))
                End While
            End If

            readerj.Close()


            '---- get master bom ---
            Dim cmd_j2 As New MySqlCommand
            cmd_j2.Parameters.AddWithValue("@job", my_job)
            cmd_j2.CommandText = "SELECT mr_name, need_date, shipping_ad from Material_Request.mr where BOM_type = 'MB' and job = @job order by release_date desc limit 1"
            cmd_j2.Connection = Login.Connection

            Dim readerj2 As MySqlDataReader
            readerj2 = cmd_j2.ExecuteReader

            If readerj2.HasRows Then
                While readerj2.Read
                    mbom = readerj2(0).ToString
                    need_by_date = readerj2(1).ToString
                    shipping_ad = readerj2(2).ToString
                End While
            End If

            readerj2.Close()

            '-----------field bom to use Merge_fuction ----------
            Dim cmd_j3 As New MySqlCommand
            cmd_j3.Parameters.AddWithValue("@job", my_job)
            cmd_j3.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = 'Field' and job = @job order by release_date desc limit 1"
            cmd_j3.Connection = Login.Connection

            Dim readerj3 As MySqlDataReader
            readerj3 = cmd_j3.ExecuteReader

            If readerj3.HasRows Then
                While readerj3.Read
                    field_bom = readerj3(0).ToString
                End While
            End If

            readerj3.Close()
            '------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub add_combo(combo As ComboBox, mylist As List(Of String))

        combo.Items.Clear()
        '--fill a combobox with particularl ist
        For i = 0 To mylist.Count - 1
            combo.Items.Add(mylist.Item(i).ToString)
        Next

    End Sub

    Private Sub Edit_BOM_Panels_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        '-part selector assem
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("Assembly")
        Part_Picker1.specify(progress_grid)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        '-- ADD PANEL
        If IsNothing(p_name1.SelectedItem) = False And IsNothing(p_name2.SelectedItem) = False And IsNothing(p_name3.SelectedItem) = False And String.IsNullOrEmpty(panel_desc.Text) = False And String.IsNullOrEmpty(qty_b.Text) = False Then

            '   Dim result1 As DialogResult = MessageBox.Show("Would you like to save your progress and keep working on other revisions before releasing it?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'ask if want to keep working

            'If (result1 = DialogResult.Yes) Then
            '    '--- store revision
            '    SAVE_revision_az.mode_t = "add"
            '    SAVE_revision_az.ShowDialog()
            'Else

            '-------- release revision immediately (add panel)

            Dim new_panel As String : new_panel = p_name1.Text & "_" & p_name2.Text & p_name3.Text

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
                Else
                    option_t = "0"
                End If

                new_panel = p_name1.Text & "-" & p_name2.Text & "." & p_name3.Text & "." & option_t
            End If

            If String.Equals(new_panel, "") = False Then

                Dim exist_c As Boolean : exist_c = False

                Try
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.AddWithValue("@mr_name", my_job & "_Panel_" & new_panel)
                    check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name"
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        exist_c = True
                    End If

                    reader.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                ' save BOM package
                If exist_c = False Then

                    If Panel_grid.Rows.Count > 1 Then  'add parts to panel

                        Dim result1 As DialogResult = MessageBox.Show("Would you like to save your progress and keep working on other revisions before releasing it?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'ask if want to keep working

                        If (result1 = DialogResult.Yes) Then
                            '--- store revision
                            SAVE_revision_az.mode_t = "add"
                            SAVE_revision_az.mr_name = my_job & "_Panel_" & new_panel
                            SAVE_revision_az.panel_n = new_panel
                            SAVE_revision_az.ShowDialog()
                        Else

                            Dim result As DialogResult = MessageBox.Show("Do you really want to add this Panel?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                            If (result = DialogResult.Yes) Then

                                Dim mr_panel As String : mr_panel = my_job & "_Panel_" & new_panel
                                Call BOM_types.Save_bom(Panel_grid, mr_panel, my_job, "Panel", False, new_panel, panel_desc.Text, 1, mbom)


                                Try

                                    Dim n_bom As Double : n_bom = 0
                                    Dim check_cmd8 As New MySqlCommand
                                    check_cmd8.Parameters.AddWithValue("@job", my_job)
                                    check_cmd8.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

                                    check_cmd8.Connection = Login.Connection
                                    check_cmd8.ExecuteNonQuery()

                                    Dim reader8 As MySqlDataReader
                                    reader8 = check_cmd8.ExecuteReader

                                    If reader8.HasRows Then
                                        While reader8.Read
                                            n_bom = reader8(0)
                                        End While
                                    End If

                                    n_bom = n_bom + 1

                                    reader8.Close()
                                    '---------------------------------------------------------

                                    Dim Create_cmd As New MySqlCommand
                                    Create_cmd.Parameters.AddWithValue("@mr_name", mr_panel)
                                    Create_cmd.Parameters.AddWithValue("@released", "Y")
                                    Create_cmd.Parameters.AddWithValue("@released_by", current_user)
                                    Create_cmd.Parameters.AddWithValue("@id_bom", n_bom)
                                    Create_cmd.Parameters.AddWithValue("@Panel_qty", If(IsNumeric(qty_b.Text) = True, qty_b.Text, 1))

                                    Create_cmd.CommandText = "UPDATE Material_Request.mr SET  released = @released,  released_by = @released_by, release_date = now(), id_bom = @id_bom, Panel_qty = @Panel_qty where mr_name = @mr_name"
                                    Create_cmd.Connection = Login.Connection
                                    Create_cmd.ExecuteNonQuery()

                                Catch ex As Exception
                                    MessageBox.Show(ex.ToString)
                                End Try

                                Call BOM_types.Create_build_request(my_job)  'create build_request
                                Call BOM_types.Create_MPL(my_job)  'create MPL

                                Call My_Material_r.Merge_and_release_MB(field_bom)  'create a revision of the Master BOM

                                MessageBox.Show("Panel added succesfully!")
                                Me.Visible = False
                                My_Material_r.Visible = False
                                BOM_init.Visible = True

                            End If
                        End If
                    Else
                        MessageBox.Show("Please Add parts to this Panel")
                    End If
                Else
                    MessageBox.Show("There is already a Panel with that name!")
                End If


            Else
                MessageBox.Show("Please enter a Panel name")
            End If
            '  End If
        Else
            MessageBox.Show("Please enter a Panel Name, description and Qty")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

        Dim result As DialogResult = MessageBox.Show("Do you really want to add this Assemblies BOM?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            Try

                '---calculate how many boms we got
                Dim n_bom As Double : n_bom = 0
                Dim check_cmd8 As New MySqlCommand
                check_cmd8.Parameters.AddWithValue("@job", my_job)
                check_cmd8.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

                check_cmd8.Connection = Login.Connection
                check_cmd8.ExecuteNonQuery()

                Dim reader8 As MySqlDataReader
                reader8 = check_cmd8.ExecuteReader

                If reader8.HasRows Then
                    While reader8.Read
                        n_bom = reader8(0)
                    End While
                End If

                n_bom = n_bom + 1

                reader8.Close()
                '---------------------------------

                For i = 0 To progress_grid.Rows.Count - 1

                    If progress_grid.Rows(i).IsNewRow Then Continue For

                    If String.IsNullOrEmpty(progress_grid.Rows(i).Cells(0).Value) = False Then

                        Dim exist_c As Boolean = False
                        Dim check_cmd As New MySqlCommand
                        check_cmd.Parameters.Clear()
                        check_cmd.Parameters.AddWithValue("@mr_name", my_job & "_Assembly_" & progress_grid.Rows(i).Cells(0).Value)
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

                            '--- enter data to mr -------
                            Dim main_cmd As New MySqlCommand
                            main_cmd.Parameters.AddWithValue("@mr_name", my_job & "_Assembly_" & progress_grid.Rows(i).Cells(0).Value)
                            main_cmd.Parameters.AddWithValue("@created_by", current_user)
                            main_cmd.Parameters.AddWithValue("@job", my_job)
                            main_cmd.Parameters.AddWithValue("@BOM_type", "Assembly")
                            main_cmd.Parameters.AddWithValue("@Panel_name", progress_grid.Rows(i).Cells(0).Value)
                            main_cmd.Parameters.AddWithValue("@Panel_desc", progress_grid.Rows(i).Cells(1).Value)
                            main_cmd.Parameters.AddWithValue("@Panel_qty", progress_grid.Rows(i).Cells(5).Value)
                            main_cmd.Parameters.AddWithValue("@MBOM", mbom)
                            main_cmd.Parameters.AddWithValue("@need_by_date", need_by_date)
                            main_cmd.Parameters.AddWithValue("@shipping_ad", shipping_ad)
                            main_cmd.Parameters.AddWithValue("@released_by", current_user)
                            main_cmd.Parameters.AddWithValue("@id_bom", n_bom)

                            main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, job, BOM_type, Panel_name, Panel_desc, Panel_qty, MBOM, need_date, shipping_ad, released, released_by, release_date, id_bom) VALUES (@mr_name, now(), @created_by, @job, @BOM_type, @Panel_name, @Panel_desc, @Panel_qty, @MBOM, @need_date, @shipping_ad, 'Y', @released_by, now(), @id_bom)"
                            main_cmd.Connection = Login.Connection
                            main_cmd.ExecuteNonQuery()


                            '-------- enter data to mr_data


                            If IsNumeric(progress_grid.Rows(i).Cells(5).Value) = True And progress_grid.Rows(i).Cells(0).Value Is Nothing = False Then

                                Dim Create_cmd6 As New MySqlCommand
                                Create_cmd6.Parameters.Clear()
                                Create_cmd6.Parameters.AddWithValue("@mr_name", my_job & "_Assembly_" & progress_grid.Rows(i).Cells(0).Value)
                                Create_cmd6.Parameters.AddWithValue("@Part_No", If(progress_grid.Rows(i).Cells(0).Value Is Nothing, "unknown part", progress_grid.Rows(i).Cells(0).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Description", If(progress_grid.Rows(i).Cells(1).Value Is Nothing, "", progress_grid.Rows(i).Cells(1).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(progress_grid.Rows(i).Cells(2).Value Is Nothing, "", progress_grid.Rows(i).Cells(2).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Vendor", If(progress_grid.Rows(i).Cells(3).Value Is Nothing, "", progress_grid.Rows(i).Cells(3).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Price", If(progress_grid.Rows(i).Cells(4).Value Is Nothing, "", progress_grid.Rows(i).Cells(4).Value.ToString.Replace("$", "")))
                                Create_cmd6.Parameters.AddWithValue("@Qty", If(progress_grid.Rows(i).Cells(5).Value Is Nothing, "", progress_grid.Rows(i).Cells(5).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@subtotal", If(progress_grid.Rows(i).Cells(6).Value Is Nothing, "", progress_grid.Rows(i).Cells(6).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@mfg_type", If(progress_grid.Rows(i).Cells(7).Value Is Nothing, "Panel", progress_grid.Rows(i).Cells(7).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Part_status", If(progress_grid.Rows(i).Cells(8).Value Is Nothing, "Special Order", progress_grid.Rows(i).Cells(8).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Part_type", If(progress_grid.Rows(i).Cells(9).Value Is Nothing, "Other", progress_grid.Rows(i).Cells(9).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@Notes", If(progress_grid.Rows(i).Cells(10).Value Is Nothing, "", progress_grid.Rows(i).Cells(10).Value.ToString))
                                Create_cmd6.Parameters.AddWithValue("@need_by_date", If(progress_grid.Rows(i).Cells(11).Value Is Nothing, "", progress_grid.Rows(i).Cells(11).Value.ToString))


                                Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes,  need_by_date) VALUES (@mr_name, @Part_No, @Description,  @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type,  @Part_status, @Part_type, @Notes,  @need_by_date)"
                                Create_cmd6.Connection = Login.Connection
                                Create_cmd6.ExecuteNonQuery()

                            End If

                            n_bom = n_bom + 1
                        End If
                    End If
                Next

                Call BOM_types.Create_build_request(my_job)  'create build_request
                Call BOM_types.Create_build_request(my_job)  'create MPL
                Call My_Material_r.Merge_and_release_MB(field_bom)  'create a revision of the Master BOM

                MessageBox.Show("Assemblies added")
                Me.Visible = False
                My_Material_r.Visible = False
                BOM_init.Visible = True

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '--- create a revision of the panel bom selected with zero qty in the mr_data and zero qty in mr

        If Not ComboBox1.SelectedItem Is Nothing Then

            Dim result As DialogResult = MessageBox.Show("Do you really want to zero out this Panel?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                Try
                    'get panel_bom 

                    Dim result1 As DialogResult = MessageBox.Show("Would you like to save your progress and keep working on other revisions before applying this change?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'ask if want to keep working

                    If (result1 = DialogResult.Yes) Then
                        '--- store revision
                        SAVE_revision_az.mode_t = "zero"
                        SAVE_revision_az.mr_name = my_job & "_Panel_" & ComboBox1.Text
                        SAVE_revision_az.panel_n = ComboBox1.Text
                        SAVE_revision_az.ShowDialog()

                    Else

                        Dim mr_name As String : mr_name = ""
                        Dim cmd_j3 As New MySqlCommand
                        cmd_j3.Parameters.AddWithValue("@job", my_job)
                        cmd_j3.Parameters.AddWithValue("@panel", ComboBox1.Text)
                        cmd_j3.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = 'Panel' and job = @job and Panel_name = @panel order by release_date desc limit 1"
                        cmd_j3.Connection = Login.Connection

                        Dim readerj3 As MySqlDataReader
                        readerj3 = cmd_j3.ExecuteReader

                        If readerj3.HasRows Then
                            While readerj3.Read
                                mr_name = readerj3(0).ToString
                            End While
                        End If

                        readerj3.Close()

                        Dim dimen_table = New DataTable
                        dimen_table.Columns.Add("Part_No", GetType(String))
                        dimen_table.Columns.Add("Description", GetType(String))
                        dimen_table.Columns.Add("Manufacturer", GetType(String))
                        dimen_table.Columns.Add("Vendor", GetType(String))
                        dimen_table.Columns.Add("Price", GetType(String))
                        dimen_table.Columns.Add("Qty", GetType(String))
                        dimen_table.Columns.Add("Subtotal", GetType(String))
                        dimen_table.Columns.Add("mfg_type", GetType(String))
                        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                        dimen_table.Columns.Add("qty_needed", GetType(String))
                        dimen_table.Columns.Add("Assembly_name", GetType(String))
                        dimen_table.Columns.Add("Part_status", GetType(String))
                        dimen_table.Columns.Add("Part_Type", GetType(String))
                        dimen_table.Columns.Add("Notes", GetType(String))
                        dimen_table.Columns.Add("Need_by_date", GetType(String))


                        Dim cmd_a As New MySqlCommand
                        cmd_a.Parameters.AddWithValue("@mr_name", mr_name)
                        cmd_a.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, Subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, Need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                        cmd_a.Connection = Login.Connection

                        Dim readera As MySqlDataReader
                        readera = cmd_a.ExecuteReader

                        If readera.HasRows Then
                            While readera.Read
                                dimen_table.Rows.Add(readera(0).ToString, readera(1).ToString, readera(2).ToString, readera(3).ToString, readera(4).ToString, readera(5).ToString, readera(6).ToString, readera(7).ToString, If(readera(8) Is DBNull.Value, "", readera(8)), If(readera(9) Is DBNull.Value, "", readera(9)), If(readera(10) Is DBNull.Value, "", readera(10)), readera(11).ToString, readera(12).ToString, readera(13).ToString, readera(14).ToString)
                            End While
                        End If

                        readera.Close()

                        '--------- start creating a revision ---------
                        Dim name_rev As String = "_rev"
                        Dim i As Integer : i = 0 'counter

                        Dim shipping As String : shipping = ""
                        Dim Date_Created As Date
                        Dim created_by As String : created_by = "unknown"
                        Dim id_bom As String : id_bom = ""

                        Dim need_date As String : need_date = ""
                        Dim Panel_name As String : Panel_name = ""
                        Dim Panel_qty As String : Panel_qty = 0
                        Dim Panel_desc As String : Panel_desc = ""
                        Dim BOM_type As String : BOM_type = ""
                        Dim MBOM As String : MBOM = ""

                        '---- get id_bom

                        Dim id_bom_r As String : id_bom_r = ""
                        Dim cmd411 As New MySqlCommand
                        cmd411.Parameters.AddWithValue("@mr_name", mr_name)
                        cmd411.CommandText = "SELECT id_bom, job from Material_Request.mr where mr_name = @mr_name"
                        cmd411.Connection = Login.Connection
                        Dim reader411 As MySqlDataReader
                        reader411 = cmd411.ExecuteReader

                        If reader411.HasRows Then
                            While reader411.Read
                                id_bom_r = reader411(0).ToString
                            End While
                        End If

                        reader411.Close()

                        '-------------------------------------
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@job", my_job)
                        cmd4.Parameters.AddWithValue("@id_bom", id_bom_r)
                        cmd4.CommandText = "Select shipping_ad, Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_type, MBOM  from Material_Request.mr where released = 'Y' and job = @job and id_bom = @id_bom order by release_date"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then

                            While reader4.Read
                                shipping = If(reader4(0) Is Nothing, "", reader4(0).ToString)
                                Date_Created = CType(reader4(1), Date)
                                created_by = If(reader4(2) Is Nothing, "unknown", reader4(2).ToString)
                                id_bom = If(reader4(3) Is Nothing, "", reader4(3).ToString)

                                need_date = If(reader4(4) Is Nothing, "", reader4(4).ToString)
                                Panel_name = If(reader4(5) Is Nothing, "", reader4(5).ToString)
                                Panel_desc = If(reader4(7) Is Nothing, "", reader4(7).ToString)
                                BOM_type = If(reader4(8) Is Nothing, "", reader4(8).ToString)
                                MBOM = If(reader4(9) Is Nothing, "", reader4(9).ToString)

                                i = i + 1
                            End While

                        End If

                        reader4.Close()

                        '----------------------------------------------
                        name_rev = name_rev & i  'last part of name of file ex: filename_revx
                        Dim indexof_s = mr_name.IndexOf("_rev")

                        If indexof_s < 0 Then
                            indexof_s = mr_name.Count
                        End If

                        '/////// start inserting revision to table /////////////

                        Dim main_cmd As New MySqlCommand
                        main_cmd.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                        main_cmd.Parameters.AddWithValue("@released_by", current_user)
                        main_cmd.Parameters.AddWithValue("@job", my_job)
                        main_cmd.Parameters.AddWithValue("@shipping", shipping)
                        main_cmd.Parameters.AddWithValue("@Date_Created", Date_Created)
                        main_cmd.Parameters.AddWithValue("@created_by", created_by)
                        main_cmd.Parameters.AddWithValue("@id_bom", id_bom)

                        main_cmd.Parameters.AddWithValue("@need_date", need_date)
                        main_cmd.Parameters.AddWithValue("@Panel_name", Panel_name)
                        main_cmd.Parameters.AddWithValue("@Panel_qty", Panel_qty)
                        main_cmd.Parameters.AddWithValue("@Panel_desc", Panel_desc)
                        main_cmd.Parameters.AddWithValue("@BOM_type", BOM_type)
                        main_cmd.Parameters.AddWithValue("@MBOM", MBOM)

                        main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, released, released_by, release_date, job, shipping_ad, Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_Type, MBOM) VALUES (@mr_name,'Y', @released_by, now(), @job, @shipping, @Date_Created, @created_by, @id_bom, @need_date, @Panel_name, @Panel_qty, @Panel_desc, @BOM_Type, @MBOM)"
                        main_cmd.Connection = Login.Connection
                        main_cmd.ExecuteNonQuery()

                        '-------- enter data to mr_data
                        For i = 0 To dimen_table.Rows.Count - 1

                            Dim Create_cmd6 As New MySqlCommand
                            Create_cmd6.Parameters.Clear()
                            Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table.Rows(i).Item(0) Is Nothing, "", dimen_table.Rows(i).Item(0)))
                            Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table.Rows(i).Item(1) Is Nothing, "", dimen_table.Rows(i).Item(1)))
                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(dimen_table.Rows(i).Item(2) Is Nothing, "", dimen_table.Rows(i).Item(2)))
                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(dimen_table.Rows(i).Item(3) Is Nothing, "", dimen_table.Rows(i).Item(3)))
                            Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(4) Is Nothing, "", dimen_table.Rows(i).Item(4).ToString.Replace("$", "")))
                            Create_cmd6.Parameters.AddWithValue("@Qty", 0)
                            Create_cmd6.Parameters.AddWithValue("@subtotal", 0)
                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(dimen_table.Rows(i).Item(7) Is Nothing, "", dimen_table.Rows(i).Item(7)))
                            Create_cmd6.Parameters.AddWithValue("@qty_fullfilled", If(dimen_table.Rows(i).Item(8) Is Nothing, "", dimen_table.Rows(i).Item(8)))
                            Create_cmd6.Parameters.AddWithValue("@qty_needed", If(dimen_table.Rows(i).Item(9) Is Nothing, "", dimen_table.Rows(i).Item(9)))
                            Create_cmd6.Parameters.AddWithValue("@Assembly_name", If(dimen_table.Rows(i).Item(10) Is Nothing, "", dimen_table.Rows(i).Item(10)))
                            Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table.Rows(i).Item(11) Is Nothing, "", dimen_table.Rows(i).Item(11)))
                            Create_cmd6.Parameters.AddWithValue("@Part_Type", If(dimen_table.Rows(i).Item(12) Is Nothing, "", dimen_table.Rows(i).Item(12)))
                            Create_cmd6.Parameters.AddWithValue("@Notes", If(dimen_table.Rows(i).Item(13) Is Nothing, "", dimen_table.Rows(i).Item(13)))
                            Create_cmd6.Parameters.AddWithValue("@Need_by_date", If(dimen_table.Rows(i).Item(14) Is Nothing, "", dimen_table.Rows(i).Item(14)))


                            Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, need_by_date ) VALUES (@mr_name, @Part_No, @Description, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fullfilled, @qty_needed, @Assembly_name, @Part_status, @Part_Type, @Notes,  @Need_by_date  )"
                            Create_cmd6.Connection = Login.Connection
                            Create_cmd6.ExecuteNonQuery()

                        Next

                        Call BOM_types.Create_build_request(my_job)    '-- create build_request revision
                        Call BOM_types.Create_MPL(my_job)   '-------- create MPL

                        Call My_Material_r.Merge_and_release_MB(field_bom)



                        '////////////////////// ---------------------  notify -----------------------
                        If enable_mess = True Then

                            Dim mail_n As String : mail_n = "Material Request Revision for Project " & my_job & "  has been released" & vbCrLf & vbCrLf _
                     & "Material Request Revised: " & mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev & vbCrLf _
                     & "Material Request Revised: " & My_Material_r.real_mr & vbCrLf _
                     & "Revised by: " & current_user



                            Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("General Management", "Material Request Revision has been Released for Project " & my_job, mail_n)


                            '--- sent email-------
                            'add email addresses
                            Dim emails_addr As New List(Of String)()

                            'procurement
                            emails_addr.Add("ecoy@atronixengineering.com")
                            emails_addr.Add("fvargas@atronixengineering.com")
                            emails_addr.Add("mmorris@atronixengineering.com")
                            emails_addr.Add("sowens@atronixengineering.com")

                            'mfg
                            emails_addr.Add("shenley@atronixengineering.com")
                            emails_addr.Add("mowens@atronixengineering.com")

                            'inventory
                            emails_addr.Add("dnix@atronixengineering.com")
                            emails_addr.Add("dmoore@atronixengineering.com")


                            ' For i = 0 To emails_addr.Count - 1

                            Try
                                Dim e_mail As New MailMessage()
                                e_mail = New MailMessage()
                                e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                                For j = 0 To emails_addr.Count - 1
                                    e_mail.To.Add(emails_addr.Item(j))
                                Next

                                e_mail.Subject = "Material Request Revision for Project " & my_job & "  has been Released"
                                e_mail.IsBodyHtml = False
                                e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                                Smtp_Server.Send(e_mail)

                            Catch error_t As Exception
                                MsgBox(error_t.ToString)
                            End Try
                            '   Next

                        End If
                        '---------------------
                        Call Inventory_manage.General_inv_cal()   'recalculate inventory values
                        MessageBox.Show("Panel Zero Out")
                        Me.Visible = False
                        My_Material_r.Visible = False
                        BOM_init.Visible = True

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Panel_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In Panel_grid.SelectedRows
                Try
                    Panel_grid.Rows.Remove(r)
                Catch ex As Exception
                    MessageBox.Show("This row cannot be deleted")
                End Try
            Next
        Else
            MessageBox.Show("Select the row you want to delete.")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        If progress_grid.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In progress_grid.SelectedRows
                Try
                    progress_grid.Rows.Remove(r)
                Catch ex As Exception
                    MessageBox.Show("This row cannot be deleted")
                End Try
            Next
        Else
            MessageBox.Show("Select the row you want to delete.")
        End If
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
            Call add_combo(p_name2, list_hp) 'Call add_combo(p_name2, listn)
            Call add_combo(p_name3, listn) 'Call add_combo(p_name3, listn)
            ' hp_l.Visible = True
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
            '   p_name4.Visible = False
            '  hp_l.Visible = False
            options_v.Visible = False


        End If
        '-----------------------------------



    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        '--------- display Panel BOM  --------

        Panel_grid.Rows.Clear()
        p_opened.Text = "Panel Opened: " & ComboBox1.Text

        Dim my_p As String : my_p = "xx"

        Try

            '-- fill date from mr
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@MBOM", Me.Text)
            cmd4.Parameters.AddWithValue("@Panel_name", ComboBox1.Text)
            cmd4.CommandText = "SELECT Panel_name, Panel_desc, Panel_qty, mr_name from Material_Request.mr where MBOM = @MBOM and Panel_name = @Panel_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    ' temp_panel_name = reader4(0).ToString
                    panel_desc.Text = reader4(1).ToString
                    qty_b.Text = reader4(2).ToString
                    my_p = reader4(3).ToString


                End While
            End If

            reader4.Close()

            '-- fill data from mr_data

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", my_p)
            cmd41.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type,  Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                Dim i As Integer : i = 0
                While reader41.Read
                    Panel_grid.Rows.Add(New String() {})
                    Panel_grid.Rows(i).Cells(0).Value = reader41(0).ToString
                    Panel_grid.Rows(i).Cells(1).Value = reader41(1).ToString
                    Panel_grid.Rows(i).Cells(2).Value = reader41(2).ToString
                    Panel_grid.Rows(i).Cells(3).Value = reader41(3).ToString
                    Panel_grid.Rows(i).Cells(4).Value = reader41(4).ToString
                    Panel_grid.Rows(i).Cells(5).Value = reader41(5).ToString
                    Panel_grid.Rows(i).Cells(6).Value = reader41(6).ToString
                    Panel_grid.Rows(i).Cells(7).Value = reader41(7).ToString
                    Panel_grid.Rows(i).Cells(8).Value = reader41(8).ToString
                    Panel_grid.Rows(i).Cells(9).Value = reader41(9).ToString
                    Panel_grid.Rows(i).Cells(10).Value = reader41(10).ToString
                    Panel_grid.Rows(i).Cells(11).Value = reader41(11).ToString

                    i = i + 1
                End While
            End If

            reader41.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:B").ColumnWidth = 40
                xlWorkSheet.Range("C:C").ColumnWidth = 30
                xlWorkSheet.Range("D:E").ColumnWidth = 20
                xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To Panel_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = Panel_grid.Columns(i).HeaderText
                    For j As Integer = 0 To Panel_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = Panel_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Dim valid_mr_name As String = If(String.IsNullOrEmpty(ComboBox1.Text) = False, ComboBox1.Text, "None")

                For Each c In Path.GetInvalidFileNameChars()
                    If valid_mr_name.Contains(c) Then
                        valid_mr_name = valid_mr_name.Replace(c, "")
                    End If
                Next

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Panel_" & valid_mr_name & "_data.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Table exported successfully!")
            End If
        End If
    End Sub

    Private Sub ImportBOMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportBOMToolStripMenuItem.Click
        '----------- Import Material Request from excel file -------------

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*ods;"

            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)

                Dim i As Integer : i = 2

                While (wb.ActiveSheet.Cells(i, 1).Value IsNot Nothing)

                    Call BOM_types.Get_part_data(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 2).value, wb.Worksheets(1).Cells(i, 3).value, wb.Worksheets(1).Cells(i, 4).value, wb.Worksheets(1).Cells(i, 5).value, wb.Worksheets(1).Cells(i, 6).value, wb.Worksheets(1).Cells(i, 7).value, "Panel", wb.Worksheets(1).Cells(i, 9).value, wb.Worksheets(1).Cells(i, 10).value, wb.Worksheets(1).Cells(i, 11).value, Panel_grid)
                    i = i + 1

                End While
                '---------------------------------------------
                'get latest cost
                For i = 0 To Panel_grid.Rows.Count - 1
                    If Panel_grid.Rows(i).DefaultCellStyle.BackColor = Color.CadetBlue Then
                        Panel_grid.Rows(i).Cells(4).Value = "$" & Panel_grid.Rows(i).Cells(4).Value
                    Else
                        Panel_grid.Rows(i).Cells(4).Value = "$" & Form1.Get_Latest_Cost(Login.Connection, Panel_grid.Rows(i).Cells(0).Value, Panel_grid.Rows(i).Cells(3).Value)
                    End If
                Next
                '-------------------------------------------


                wb.Close(False)
                MessageBox.Show("Done")

            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        '------ create a Panel rev with a different qty panel number 

        If Not ComboBox1.SelectedItem Is Nothing Then

            Try

                Dim qty_change As Integer
                Dim qty_test As Integer

                qty_change = InputBox("Please enter the new number of " & ComboBox1.Text)

                If Integer.TryParse(qty_change, qty_test) Then
                    'get panel_bom 

                    Dim result1 As DialogResult = MessageBox.Show("Would you like to save your progress and keep working on other revisions before applying this change?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'ask if want to keep working

                    If (result1 = DialogResult.Yes) Then
                        '--- store revision
                        SAVE_revision_az.mode_t = "change"
                        SAVE_revision_az.p_qty = qty_change
                        SAVE_revision_az.mr_name = my_job & "_Panel_" & ComboBox1.Text
                        SAVE_revision_az.panel_n = ComboBox1.Text
                        SAVE_revision_az.ShowDialog()

                    Else

                        Dim mr_name As String : mr_name = ""
                        Dim cmd_j3 As New MySqlCommand
                        cmd_j3.Parameters.AddWithValue("@job", my_job)
                        cmd_j3.Parameters.AddWithValue("@panel", ComboBox1.Text)
                        cmd_j3.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = 'Panel' and job = @job and Panel_name = @panel order by release_date desc limit 1"
                        cmd_j3.Connection = Login.Connection

                        Dim readerj3 As MySqlDataReader
                        readerj3 = cmd_j3.ExecuteReader

                        If readerj3.HasRows Then
                            While readerj3.Read
                                mr_name = readerj3(0).ToString
                            End While
                        End If

                        readerj3.Close()

                        Dim dimen_table = New DataTable
                        dimen_table.Columns.Add("Part_No", GetType(String))
                        dimen_table.Columns.Add("Description", GetType(String))
                        dimen_table.Columns.Add("Manufacturer", GetType(String))
                        dimen_table.Columns.Add("Vendor", GetType(String))
                        dimen_table.Columns.Add("Price", GetType(String))
                        dimen_table.Columns.Add("Qty", GetType(String))
                        dimen_table.Columns.Add("Subtotal", GetType(String))
                        dimen_table.Columns.Add("mfg_type", GetType(String))
                        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                        dimen_table.Columns.Add("qty_needed", GetType(String))
                        dimen_table.Columns.Add("Assembly_name", GetType(String))
                        dimen_table.Columns.Add("Part_status", GetType(String))
                        dimen_table.Columns.Add("Part_Type", GetType(String))
                        dimen_table.Columns.Add("Notes", GetType(String))
                        dimen_table.Columns.Add("Need_by_date", GetType(String))


                        Dim cmd_a As New MySqlCommand
                        cmd_a.Parameters.AddWithValue("@mr_name", mr_name)
                        cmd_a.CommandText = "SELECT Part_No, Description, Manufacturer, Vendor, Price, Qty, Subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, Need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                        cmd_a.Connection = Login.Connection

                        Dim readera As MySqlDataReader
                        readera = cmd_a.ExecuteReader

                        If readera.HasRows Then
                            While readera.Read
                                dimen_table.Rows.Add(readera(0).ToString, readera(1).ToString, readera(2).ToString, readera(3).ToString, readera(4).ToString, readera(5).ToString, readera(6).ToString, readera(7).ToString, If(readera.IsDBNull(8) = True, "0", readera(8)), If(readera.IsDBNull(9) = True, "0", readera(9)), If(readera.IsDBNull(10) = True, "", readera(10)), readera(11).ToString, readera(12).ToString, readera(13).ToString, readera(14).ToString)
                            End While
                        End If

                        readera.Close()

                        '--------- start creating a revision ---------
                        Dim name_rev As String = "_rev"
                        Dim i As Integer : i = 0 'counter

                        Dim shipping As String : shipping = ""
                        Dim Date_Created As Date
                        Dim created_by As String : created_by = "unknown"
                        Dim id_bom As String : id_bom = ""

                        Dim need_date As String : need_date = ""
                        Dim Panel_name As String : Panel_name = ""
                        Dim Panel_qty As String : Panel_qty = qty_change
                        Dim Panel_desc As String : Panel_desc = ""
                        Dim BOM_type As String : BOM_type = ""
                        Dim MBOM As String : MBOM = ""

                        '---- get id_bom

                        Dim id_bom_r As String : id_bom_r = ""
                        Dim cmd411 As New MySqlCommand
                        cmd411.Parameters.AddWithValue("@mr_name", mr_name)
                        cmd411.CommandText = "SELECT id_bom, job from Material_Request.mr where mr_name = @mr_name"
                        cmd411.Connection = Login.Connection
                        Dim reader411 As MySqlDataReader
                        reader411 = cmd411.ExecuteReader

                        If reader411.HasRows Then
                            While reader411.Read
                                id_bom_r = reader411(0).ToString
                            End While
                        End If

                        reader411.Close()

                        '-------------------------------------
                        Dim cmd4 As New MySqlCommand
                        cmd4.Parameters.AddWithValue("@job", my_job)
                        cmd4.Parameters.AddWithValue("@id_bom", id_bom_r)
                        cmd4.CommandText = "Select shipping_ad, Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_type, MBOM  from Material_Request.mr where released = 'Y' and job = @job and id_bom = @id_bom order by release_date"
                        cmd4.Connection = Login.Connection
                        Dim reader4 As MySqlDataReader
                        reader4 = cmd4.ExecuteReader

                        If reader4.HasRows Then

                            While reader4.Read
                                shipping = If(reader4(0) Is Nothing, "", reader4(0).ToString)
                                Date_Created = CType(reader4(1), Date)
                                created_by = If(reader4(2) Is Nothing, "unknown", reader4(2).ToString)
                                id_bom = If(reader4(3) Is Nothing, "", reader4(3).ToString)

                                need_date = If(reader4(4) Is Nothing, "", reader4(4).ToString)
                                Panel_name = If(reader4(5) Is Nothing, "", reader4(5).ToString)
                                Panel_desc = If(reader4(7) Is Nothing, "", reader4(7).ToString)
                                BOM_type = If(reader4(8) Is Nothing, "", reader4(8).ToString)
                                MBOM = If(reader4(9) Is Nothing, "", reader4(9).ToString)

                                i = i + 1
                            End While

                        End If

                        reader4.Close()

                        '----------------------------------------------
                        name_rev = name_rev & i  'last part of name of file ex: filename_revx
                        Dim indexof_s = mr_name.IndexOf("_rev")

                        If indexof_s < 0 Then
                            indexof_s = mr_name.Count
                        End If

                        '/////// start inserting revision to table /////////////

                        Dim main_cmd As New MySqlCommand
                        main_cmd.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                        main_cmd.Parameters.AddWithValue("@released_by", current_user)
                        main_cmd.Parameters.AddWithValue("@job", my_job)
                        main_cmd.Parameters.AddWithValue("@shipping", shipping)
                        main_cmd.Parameters.AddWithValue("@Date_Created", Date_Created)
                        main_cmd.Parameters.AddWithValue("@created_by", created_by)
                        main_cmd.Parameters.AddWithValue("@id_bom", id_bom)

                        main_cmd.Parameters.AddWithValue("@need_date", need_date)
                        main_cmd.Parameters.AddWithValue("@Panel_name", Panel_name)
                        main_cmd.Parameters.AddWithValue("@Panel_qty", Panel_qty)
                        main_cmd.Parameters.AddWithValue("@Panel_desc", Panel_desc)
                        main_cmd.Parameters.AddWithValue("@BOM_type", BOM_type)
                        main_cmd.Parameters.AddWithValue("@MBOM", MBOM)

                        main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, released, released_by, release_date, job, shipping_ad, Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_Type, MBOM) VALUES (@mr_name,'Y', @released_by, now(), @job, @shipping, @Date_Created, @created_by, @id_bom, @need_date, @Panel_name, @Panel_qty, @Panel_desc, @BOM_Type, @MBOM)"
                        main_cmd.Connection = Login.Connection
                        main_cmd.ExecuteNonQuery()

                        '-------- enter data to mr_data
                        For i = 0 To dimen_table.Rows.Count - 1

                            Dim Create_cmd6 As New MySqlCommand
                            Create_cmd6.Parameters.Clear()
                            Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                            Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table.Rows(i).Item(0) Is Nothing, "", dimen_table.Rows(i).Item(0)))
                            Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table.Rows(i).Item(1) Is Nothing, "", dimen_table.Rows(i).Item(1)))
                            Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(dimen_table.Rows(i).Item(2) Is Nothing, "", dimen_table.Rows(i).Item(2)))
                            Create_cmd6.Parameters.AddWithValue("@Vendor", If(dimen_table.Rows(i).Item(3) Is Nothing, "", dimen_table.Rows(i).Item(3)))
                            Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(4) Is Nothing, "", dimen_table.Rows(i).Item(4).ToString.Replace("$", "")))
                            Create_cmd6.Parameters.AddWithValue("@Qty", If(dimen_table.Rows(i).Item(5) Is Nothing, "0", dimen_table.Rows(i).Item(5)))
                            Create_cmd6.Parameters.AddWithValue("@subtotal", 0)
                            Create_cmd6.Parameters.AddWithValue("@mfg_type", If(dimen_table.Rows(i).Item(7) Is Nothing, "", dimen_table.Rows(i).Item(7)))
                            Create_cmd6.Parameters.AddWithValue("@qty_fullfilled", If(dimen_table.Rows(i).Item(8) Is Nothing, "", dimen_table.Rows(i).Item(8)))
                            Create_cmd6.Parameters.AddWithValue("@qty_needed", If(dimen_table.Rows(i).Item(9) Is Nothing, "", dimen_table.Rows(i).Item(9)))
                            Create_cmd6.Parameters.AddWithValue("@Assembly_name", If(dimen_table.Rows(i).Item(10) Is Nothing, "", dimen_table.Rows(i).Item(10)))
                            Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table.Rows(i).Item(11) Is Nothing, "", dimen_table.Rows(i).Item(11)))
                            Create_cmd6.Parameters.AddWithValue("@Part_Type", If(dimen_table.Rows(i).Item(12) Is Nothing, "", dimen_table.Rows(i).Item(12)))
                            Create_cmd6.Parameters.AddWithValue("@Notes", If(dimen_table.Rows(i).Item(13) Is Nothing, "", dimen_table.Rows(i).Item(13)))
                            Create_cmd6.Parameters.AddWithValue("@Need_by_date", If(dimen_table.Rows(i).Item(14) Is Nothing, "", dimen_table.Rows(i).Item(14)))


                            Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, need_by_date ) VALUES (@mr_name, @Part_No, @Description, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fullfilled, @qty_needed, @Assembly_name, @Part_status, @Part_Type, @Notes,  @Need_by_date  )"
                            Create_cmd6.Connection = Login.Connection
                            Create_cmd6.ExecuteNonQuery()

                        Next

                        Call BOM_types.Create_build_request(my_job)    '-- create build_request revision
                        Call BOM_types.Create_MPL(my_job)   '-------- create MPL

                        Call My_Material_r.Merge_and_release_MB(field_bom)



                        '////////////////////// ---------------------  notify -----------------------
                        If enable_mess = True Then

                            Dim mail_n As String : mail_n = "Material Request Revision for Project " & my_job & "  has been released" & vbCrLf & vbCrLf _
             & "Material Request Revised: " & mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev & vbCrLf _
             & "Material Request Revised: " & My_Material_r.real_mr & vbCrLf _
             & "Revised by: " & current_user



                            Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request Revision has been Released for Project " & my_job, mail_n)
                            Call Sent_mail.Sent_multiple_emails("General Management", "Material Request Revision has been Released for Project " & my_job, mail_n)


                            '--- sent email-------
                            'add email addresses
                            Dim emails_addr As New List(Of String)()

                            'procurement
                            emails_addr.Add("ecoy@atronixengineering.com")
                            emails_addr.Add("fvargas@atronixengineering.com")
                            emails_addr.Add("mmorris@atronixengineering.com")
                            emails_addr.Add("sowens@atronixengineering.com")

                            'mfg
                            emails_addr.Add("shenley@atronixengineering.com")
                            emails_addr.Add("mowens@atronixengineering.com")

                            'inventory
                            emails_addr.Add("dnix@atronixengineering.com")
                            emails_addr.Add("dmoore@atronixengineering.com")


                            ' For i = 0 To emails_addr.Count - 1

                            Try
                                Dim e_mail As New MailMessage()
                                e_mail = New MailMessage()
                                e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                                For j = 0 To emails_addr.Count - 1
                                    e_mail.To.Add(emails_addr.Item(j))
                                Next

                                e_mail.Subject = "Material Request Revision for Project " & my_job & "  has been Released"
                                e_mail.IsBodyHtml = False
                                e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                                Smtp_Server.Send(e_mail)

                            Catch error_t As Exception
                                MsgBox(error_t.ToString)
                            End Try
                            '   Next

                        End If
                        '---------------------
                        Call Inventory_manage.General_inv_cal()   'recalculate inventory values
                        MessageBox.Show("Panel Qty Changed")
                        Me.Visible = False
                        My_Material_r.Visible = False
                        BOM_init.Visible = True

                    End If
                End If

            Catch ex As Exception
                MessageBox.Show("Enter an integer, please")
            End Try

        Else
            MessageBox.Show("Please, select a panel")
        End If


    End Sub

    Private Sub options_v_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles options_v.ItemCheck
        For i = 0 To options_v.Items.Count - 1
            If (i <> e.Index) = True Then
                options_v.SetItemChecked(i, False)
            End If
        Next
    End Sub

    Private Sub SeeRevisionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeeRevisionsToolStripMenuItem.Click
        TabControl1.TabPages.Insert(1, TabPage2)

        open_grid.Rows.Clear()

        Try
            '-------- get MBOM ------------
            Dim MB As String : MB = ""

            Dim cmdm As New MySqlCommand
            cmdm.Parameters.AddWithValue("@job", My_Material_r.open_job)
            cmdm.Parameters.AddWithValue("@BOM_type", "MB")
            cmdm.CommandText = "SELECT mr_name from Material_Request.mr where BOM_type = @BOM_type  and job = @job order by release_date desc limit 1"
            cmdm.Connection = Login.Connection
            Dim readerm As MySqlDataReader
            readerm = cmdm.ExecuteReader

            If readerm.HasRows Then
                While readerm.Read
                    MB = readerm(0).ToString
                End While
            End If

            readerm.Close()
            '----------------------------

            Dim cmd48 As New MySqlCommand
            cmd48.Parameters.AddWithValue("@MB", MB)
            cmd48.CommandText = "SELECT mr_name, rev_name, created_date, created_by from Revisions.mr_rev where MB = @MB and new_panel = 'Add'"
            cmd48.Connection = Login.Connection
            Dim reader48 As MySqlDataReader
            reader48 = cmd48.ExecuteReader

            If reader48.HasRows Then
                Dim i As Integer : i = 0
                While reader48.Read
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(i).Cells(0).Value = reader48(0).ToString
                    open_grid.Rows(i).Cells(1).Value = reader48(1).ToString
                    open_grid.Rows(i).Cells(2).Value = reader48(2).ToString
                    open_grid.Rows(i).Cells(3).Value = reader48(3).ToString

                    i = i + 1
                End While
            End If

            reader48.Close()

            TabControl1.SelectedTab = TabPage2

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        progress_grid.Rows.Clear()
        TabControl1.TabPages.Remove(TabPage2)
    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        '--- open panel revision -----

        progress_grid.Rows.Clear()

        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@mr_name", open_grid.Rows(index_k).Cells(0).Value)
            check_cmd1.Parameters.AddWithValue("@rev_name", open_grid.Rows(index_k).Cells(1).Value)
            check_cmd1.CommandText = "select Part_No, description_t, Manufacturer, Vendor, Price, new_qty, mfg_type, part_status, part_type, Notes, need_date  from Revisions.mr_rev_data where mr_name = @mr_name and rev_name = @rev_name"

            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then

                While reader1.Read
                    progress_grid.Rows.Add(New String() {reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString, reader1(5).ToString, 0, reader1(6).ToString, reader1(7).ToString, reader1(8).ToString, reader1(9).ToString, reader1(10).ToString})
                End While
            End If

            reader1.Close()

            inp_mr_name = open_grid.Rows(index_k).Cells(0).Value
            inp_rev_name = open_grid.Rows(index_k).Cells(1).Value

        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click

        '----------- open panel in progress -----------
        Try

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@mr_name", inp_mr_name)
            check_cmd1.Parameters.AddWithValue("@rev_name", inp_rev_name)
            check_cmd1.CommandText = "select panel_name, desc_name, qty_name  from Revisions.mr_rev where mr_name = @mr_name and rev_name = @rev_name"

            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then
                While reader1.Read
                    p_opened.Text = "Panel In Progress Opened:  " & If(IsDBNull(reader1(0)), "", reader1(0).ToString)
                    panel_desc.Text = If(IsDBNull(reader1(1)), "", reader1(1).ToString)
                    qty_b.Text = If(IsDBNull(reader1(2)), "", reader1(2).ToString)

                End While
            End If

            reader1.Close()

            Panel_grid.Rows.Clear()

            For i = 0 To progress_grid.Rows.Count - 1
                Panel_grid.Rows.Add(New String() {progress_grid.Rows(i).Cells(0).Value, progress_grid.Rows(i).Cells(1).Value, progress_grid.Rows(i).Cells(2).Value, progress_grid.Rows(i).Cells(3).Value, progress_grid.Rows(i).Cells(4).Value, progress_grid.Rows(i).Cells(5).Value, progress_grid.Rows(i).Cells(6).Value, progress_grid.Rows(i).Cells(7).Value, progress_grid.Rows(i).Cells(8).Value, progress_grid.Rows(i).Cells(9).Value, progress_grid.Rows(i).Cells(10).Value, progress_grid.Rows(i).Cells(11).Value})
            Next


            TabControl1.SelectedTab = TabPage1
            progress_grid.Rows.Clear()
            TabControl1.TabPages.Remove(TabPage2)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub open_grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles open_grid.CellClick
        '--- open panel revision -----

        progress_grid.Rows.Clear()

        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@mr_name", open_grid.Rows(index_k).Cells(0).Value)
            check_cmd1.Parameters.AddWithValue("@rev_name", open_grid.Rows(index_k).Cells(1).Value)
            check_cmd1.CommandText = "select Part_No, description_t, Manufacturer, Vendor, Price, new_qty, mfg_type, part_status, part_type, Notes, need_date  from Revisions.mr_rev_data where mr_name = @mr_name and rev_name = @rev_name"

            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then

                While reader1.Read
                    progress_grid.Rows.Add(New String() {reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString, reader1(5).ToString, 0, reader1(6).ToString, reader1(7).ToString, reader1(8).ToString, reader1(9).ToString, reader1(10).ToString})
                End While
            End If

            reader1.Close()

            inp_mr_name = open_grid.Rows(index_k).Cells(0).Value
            inp_rev_name = open_grid.Rows(index_k).Cells(1).Value

        End If
    End Sub
End Class