Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO

Public Class My_Material_r

    Public mode As String ' this control the open dialog box
    Public isitreleased As Boolean
    Public open_job As String
    Public index As Integer 'control excel import iteration
    Public Smtp_Server As New SmtpClient
    Public rev_mode As Boolean

    Public real_mr As String



    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'If isitreleased = False Then
        '    Inventory_manage.part_sel = ""
        '    Part_selector.ShowDialog()
        'Else
        '    If rev_mode = True Then
        '        Inventory_manage.part_sel = "revision"
        '        Part_selector.ShowDialog()
        '    End If
        'End If

        '--- part spare order
        Part_Picker1.load_data()
        Part_Picker1.Width = 1195
        Part_Picker1.Height = 709
        Part_Picker1.mfg_type_Set("x")
        Part_Picker1.specify(rev_grid)
        Part_Picker1.qty_box.Text = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TabControl1.SelectedTab Is TabPage1 Then


            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer

            For Each row As DataGridViewRow In PR_grid.Rows
                If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False Then
                    If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text, True) = 0 Then
                        rowindex = row.Index
                        PR_grid.CurrentCell = PR_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If

                End If
            Next

            If found_po = False Then
                MsgBox("Part not found!")
            End If

        ElseIf TabControl1.SelectedTab Is TabPage3 Then

            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer

            For Each row As DataGridViewRow In single_grid.Rows
                If String.IsNullOrEmpty(row.Cells.Item("DataGridViewTextBoxColumn3").Value) = False Then
                    If String.Compare(row.Cells.Item("DataGridViewTextBoxColumn3").Value.ToString, TextBox1.Text, True) = 0 Then
                        rowindex = row.Index
                        single_grid.CurrentCell = single_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If

                End If
            Next

            If found_po = False Then
                MsgBox("Part not found!")
            End If

        ElseIf TabControl1.SelectedTab Is TabPage4 Then


            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer

            For Each row As DataGridViewRow In compare_grid.Rows
                If String.IsNullOrEmpty(row.Cells.Item("DataGridViewTextBoxColumn17").Value) = False Then
                    If String.Compare(row.Cells.Item("DataGridViewTextBoxColumn17").Value.ToString, TextBox1.Text, True) = 0 Then
                        rowindex = row.Index
                        compare_grid.CurrentCell = compare_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If

                End If
            Next

            If found_po = False Then
                MsgBox("Part not found!")
            End If
        Else

            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer

            For Each row As DataGridViewRow In rev_grid.Rows
                If String.IsNullOrEmpty(row.Cells.Item("DataGridViewTextBoxColumn1").Value) = False Then
                    If String.Compare(row.Cells.Item("DataGridViewTextBoxColumn1").Value.ToString, TextBox1.Text, True) = 0 Then
                        rowindex = row.Index
                        rev_grid.CurrentCell = rev_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If

                End If
            Next

            If found_po = False Then
                MsgBox("Part not found!")
            End If

        End If
    End Sub

    'Private Sub CreateNewMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewMaterialRequestToolStripMenuItem.Click
    '    PR_grid.Rows.Clear()
    '    compare_grid.Rows.Clear()
    '    single_grid.Rows.Clear()
    '    ComboBox2.Items.Clear()
    '    mrbox1.Items.Clear()
    '    mrbox2.Items.Clear()

    '    Me.Text = "My Material Requests"
    '    isitreleased = False
    '    job_label.Text = "Open Project:"
    '    TabControl1.TabPages.Remove(TabPage2)
    '    Inventory_manage.part_sel = ""
    '    rev_mode = False
    'End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    'Private Sub ActiveMaterialRequestsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ActiveMaterialRequestsToolStripMenuItem.Click
    '    mode = "mr_inpro"
    '    Open_file.ShowDialog()
    'End Sub

    'Private Sub ReleasedMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReleasedMaterialRequestToolStripMenuItem.Click
    '    mode = "mr_rel"
    '    Open_file.ShowDialog()

    'End Sub

    'Private Sub DeleteMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteMaterialRequestToolStripMenuItem.Click
    '    '=== delete mr only if is in progress and you are general or engineer_manager


    '    Dim go_delete As Boolean : go_delete = True

    '    If String.Equals(Me.Text, "My Material Requests") = False And isitreleased = False Then

    '        Try
    '            Dim check_cmd As New MySqlCommand
    '            check_cmd.Parameters.AddWithValue("@mr_name", Me.Text)
    '            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name and released = 'Y'"
    '            check_cmd.Connection = Login.Connection
    '            check_cmd.ExecuteNonQuery()

    '            Dim reader As MySqlDataReader
    '            reader = check_cmd.ExecuteReader

    '            If reader.HasRows Then
    '                go_delete = False
    '            End If

    '            reader.Close()

    '            If go_delete = True Then
    '                If Engineer_management = True Or General_management = True Then

    '                    Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this Material Request?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

    '                    If (result = DialogResult.Yes) Then

    '                        Dim check_cmd2 As New MySqlCommand
    '                        check_cmd2.Parameters.AddWithValue("@mr_name", Me.Text)
    '                        check_cmd2.CommandText = "delete from Material_Request.mr where mr_name = @mr_name"
    '                        check_cmd2.Connection = Login.Connection
    '                        check_cmd2.ExecuteNonQuery()

    '                        Dim check_cmd3 As New MySqlCommand
    '                        check_cmd3.Parameters.AddWithValue("@mr_name", Me.Text)
    '                        check_cmd3.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
    '                        check_cmd3.Connection = Login.Connection
    '                        check_cmd3.ExecuteNonQuery()

    '                        PR_grid.Rows.Clear()
    '                        single_grid.Rows.Clear()
    '                        compare_grid.Rows.Clear()

    '                        Call Form1.Command_h(current_user, Me.Text & " was deleted", "None")
    '                        MessageBox.Show("Material Request deleted succesfully!")
    '                        Me.Text = "My Material Requests"

    '                    End If

    '                Else
    '                    MessageBox.Show("You are not authorized to delete a Material Request. Contact your supervisor")

    '                End If

    '            End If


    '        Catch ex As Exception
    '            MessageBox.Show(ex.ToString)
    '        End Try

    '    Else
    '        MessageBox.Show("Material Request not found")
    '    End If



    'End Sub

    'Private Sub SaveMRToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveMRToolStripMenuItem.Click
    '    '---- Save changes, will update the mr table, delete the mr_data and insert new data with the former name


    '    If String.Equals(Me.Text, "My Material Requests") = False And isitreleased = False Then

    '        If Special_order_check() = True Then

    '            Try
    '                Dim Create_cmd As New MySqlCommand
    '                Create_cmd.Parameters.AddWithValue("@mr_name", Me.Text)
    '                Create_cmd.CommandText = "UPDATE Material_Request.mr SET last_modified = now() where mr_name = @mr_name"
    '                Create_cmd.Connection = Login.Connection
    '                Create_cmd.ExecuteNonQuery()

    '                '--------- delete data from mr_data
    '                Dim Create_cmd2 As New MySqlCommand
    '                Create_cmd2.Parameters.AddWithValue("@mr_name", Me.Text)

    '                Create_cmd2.CommandText = "delete from Material_Request.mr_data where mr_name = @mr_name"
    '                Create_cmd2.Connection = Login.Connection
    '                Create_cmd2.ExecuteNonQuery()

    '                '--------- insert new data -------
    '                '-------- enter data to mr_data
    '                For i = 0 To PR_grid.Rows.Count - 1

    '                    If IsNumeric(PR_grid.Rows(i).Cells(6).Value) = True And PR_grid.Rows(i).Cells(0).Value Is Nothing = False Then

    '                        Dim Create_cmd6 As New MySqlCommand
    '                        Create_cmd6.Parameters.Clear()
    '                        Create_cmd6.Parameters.AddWithValue("@mr_name", Me.Text)
    '                        Create_cmd6.Parameters.AddWithValue("@Part_No", If(PR_grid.Rows(i).Cells(0).Value Is Nothing, "", PR_grid.Rows(i).Cells(0).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Description", If(PR_grid.Rows(i).Cells(1).Value Is Nothing, "", PR_grid.Rows(i).Cells(1).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(PR_grid.Rows(i).Cells(2).Value Is Nothing, "", PR_grid.Rows(i).Cells(2).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(PR_grid.Rows(i).Cells(3).Value Is Nothing, "", PR_grid.Rows(i).Cells(3).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Vendor", If(PR_grid.Rows(i).Cells(4).Value Is Nothing, "", PR_grid.Rows(i).Cells(4).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Price", If(PR_grid.Rows(i).Cells(5).Value Is Nothing, "", PR_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")))
    '                        Create_cmd6.Parameters.AddWithValue("@Qty", If(PR_grid.Rows(i).Cells(6).Value Is Nothing, "", PR_grid.Rows(i).Cells(6).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@subtotal", If(PR_grid.Rows(i).Cells(7).Value Is Nothing, "", PR_grid.Rows(i).Cells(7).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@mfg_type", If(PR_grid.Rows(i).Cells(8).Value Is Nothing, "", PR_grid.Rows(i).Cells(8).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@assembly", If(PR_grid.Rows(i).Cells(9).Value Is Nothing, "", PR_grid.Rows(i).Cells(9).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Notes", If(PR_grid.Rows(i).Cells(10).Value Is Nothing, "", PR_grid.Rows(i).Cells(10).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Part_status", If(PR_grid.Rows(i).Cells(11).Value Is Nothing, "", PR_grid.Rows(i).Cells(11).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@Part_Type", If(PR_grid.Rows(i).Cells(12).Value Is Nothing, "", PR_grid.Rows(i).Cells(12).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@full_panel", If(PR_grid.Rows(i).Cells(13).Value Is Nothing, "", PR_grid.Rows(i).Cells(13).Value.ToString))
    '                        Create_cmd6.Parameters.AddWithValue("@need_by_date", If(PR_grid.Rows(i).Cells(14).Value Is Nothing, "", PR_grid.Rows(i).Cells(14).Value.ToString))

    '                        Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @assembly, @Notes, @Part_status, @Part_type, @full_panel, @need_by_date)"
    '                        Create_cmd6.Connection = Login.Connection
    '                        Create_cmd6.ExecuteNonQuery()

    '                    End If

    '                Next

    '                MessageBox.Show("Changes Saved")

    '            Catch ex As Exception
    '                MessageBox.Show(ex.ToString)
    '            End Try

    '        Else

    '            MessageBox.Show("Invalid Special Order Part found")
    '        End If

    '    Else
    '        MessageBox.Show("Material Request not found")
    '    End If
    'End Sub

    'Private Sub SaveMRAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveMRAsToolStripMenuItem.Click
    '    Purchase_Request.whatipress = 3
    '    Save_MR.ShowDialog()

    'End Sub

    Private Sub PR_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles PR_grid.CellValueChanged
        For Each row As DataGridViewRow In PR_grid.Rows
            If row.IsNewRow Then Continue For
            If (IsNumeric(row.Cells(5).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = row.Cells(6).Value * row.Cells(5).Value
            End If
        Next
    End Sub

    Private Sub My_Material_r_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'hi
        ComboBox3.Text = "All BOMs"
        real_mr = ""

        '--check comparison boxes
        mrbox1.Items.Clear()
        mrbox2.Items.Clear()

        'Dim check_cmd As New MySqlCommand

        'check_cmd.CommandText = "select distinct mr_name from Material_Request.mr"
        'check_cmd.Connection = Login.Connection
        'check_cmd.ExecuteNonQuery()

        'Dim reader As MySqlDataReader
        'reader = check_cmd.ExecuteReader

        'If reader.HasRows Then
        '    While reader.Read
        '        mrbox1.Items.Add(reader(0))
        '        mrbox2.Items.Add(reader(0))
        '    End While
        'End If

        'reader.Close()
        '----------------------------------

        PR_grid.Columns(15).Visible = False
        PR_grid.Columns(13).Visible = False  'full panels

        If String.Equals(current_user, "mowens") = True Or String.Equals(current_user, "shenley") = True Then
            PR_grid.Columns(2).Visible = False
            PR_grid.Columns(5).Visible = False
            PR_grid.Columns(7).Visible = False
            PR_grid.Columns(8).Visible = False
            PR_grid.Columns(9).Visible = False
            PR_grid.Columns(10).Visible = False
            PR_grid.Columns(11).Visible = False
            PR_grid.Columns(12).Visible = False
            PR_grid.Columns(13).Visible = False
            PR_grid.Columns(15).Visible = True
        End If


        isitreleased = False
        ComboBox1.Text = "Delta"
        open_job = ""
        rev_mode = False

        TabControl1.TabPages.Remove(TabPage2)
        EnableDoubleBuffered(PR_grid)
        EnableDoubleBuffered(rev_grid)
        EnableDoubleBuffered(single_grid)
        EnableDoubleBuffered(compare_grid)

        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"

    End Sub

    'Private Sub RelaseMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RelaseMaterialRequestToolStripMenuItem.Click

    '    Assign_box.ShowDialog()

    'End Sub

    Private Sub CreateMRRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateMRRevisionToolStripMenuItem.Click

        If isitreleased = True Then

            If isit_MB(Me.Text) = False Then

                If rev_mode = False Then

                    Inventory_manage.part_sel = "revision"
                    rev_mode = True
                    rev_grid.Rows.Clear()

                    For i = 0 To PR_grid.Rows.Count - 1

                        If PR_grid.Rows(i).IsNewRow Then Continue For

                        rev_grid.Rows.Add()
                        rev_grid.Rows(i).Cells(0).Value = PR_grid.Rows(i).Cells(0).Value : rev_grid.Rows(i).Cells(0).ReadOnly = True 'part name
                        rev_grid.Rows(i).Cells(1).Value = PR_grid.Rows(i).Cells(1).Value : rev_grid.Rows(i).Cells(1).ReadOnly = True 'description
                        rev_grid.Rows(i).Cells(2).Value = PR_grid.Rows(i).Cells(2).Value : rev_grid.Rows(i).Cells(2).ReadOnly = True  'ada
                        rev_grid.Rows(i).Cells(3).Value = PR_grid.Rows(i).Cells(3).Value : rev_grid.Rows(i).Cells(3).ReadOnly = True  'manu
                        rev_grid.Rows(i).Cells(4).Value = PR_grid.Rows(i).Cells(4).Value : rev_grid.Rows(i).Cells(4).ReadOnly = True  'vendor
                        rev_grid.Rows(i).Cells(5).Value = PR_grid.Rows(i).Cells(5).Value : rev_grid.Rows(i).Cells(5).ReadOnly = True  'price
                        rev_grid.Rows(i).Cells(6).Value = PR_grid.Rows(i).Cells(8).Value : rev_grid.Rows(i).Cells(6).ReadOnly = False  'mfg type
                        rev_grid.Rows(i).Cells(7).Value = PR_grid.Rows(i).Cells(6).Value : rev_grid.Rows(i).Cells(7).ReadOnly = True  'current qty
                        rev_grid.Rows(i).Cells(8).Value = PR_grid.Rows(i).Cells(6).Value  'new qty
                        rev_grid.Rows(i).Cells(10).Value = PR_grid.Rows(i).Cells(9).Value  'assembly
                        rev_grid.Rows(i).Cells(11).Value = PR_grid.Rows(i).Cells(10).Value : rev_grid.Rows(i).Cells(11).ReadOnly = False  'notes
                        rev_grid.Rows(i).Cells(12).Value = PR_grid.Rows(i).Cells(11).Value : rev_grid.Rows(i).Cells(12).ReadOnly = True  'status
                        rev_grid.Rows(i).Cells(13).Value = PR_grid.Rows(i).Cells(12).Value : rev_grid.Rows(i).Cells(13).ReadOnly = True  'type
                        rev_grid.Rows(i).Cells(14).Value = PR_grid.Rows(i).Cells(13).Value  'rev_grid.Rows(i).Cells(14).ReadOnly = True  'status
                        rev_grid.Rows(i).Cells(15).Value = PR_grid.Rows(i).Cells(14).Value  'rev_grid.Rows(i).Cells(15).ReadOnly = True  'type
                    Next

                    TabControl1.TabPages.Insert(2, TabPage2)
                    TabControl1.SelectedTab = TabPage2
                Else
                    MessageBox.Show("You are already on revision mode")
                End If

            Else
                MessageBox.Show("Please make a Revision of the specific Panel, Field, Assembly or Spare Parts BOM instead")
            End If
        Else
                MessageBox.Show("This Material Request has not been released")
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Cursor.Current = Cursors.WaitCursor

        'Create revision of released mr
        Inventory_manage.part_sel = ""
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


        Try

            '---- get id_bom

            Dim id_bom_r As String : id_bom_r = ""
            Dim cmd411 As New MySqlCommand
            cmd411.Parameters.AddWithValue("@mr_name", Me.Text)
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



            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", open_job)
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
                    Panel_qty = If(IsNumeric(reader4(6)) = False, 0, reader4(6).ToString)
                    Panel_desc = If(reader4(7) Is Nothing, "", reader4(7).ToString)
                    BOM_type = If(reader4(8) Is Nothing, "", reader4(8).ToString)
                    MBOM = If(reader4(9) Is Nothing, "", reader4(9).ToString)

                    i = i + 1
                End While
            End If

            reader4.Close()

            name_rev = name_rev & i  'last part of name of file ex: filename_revx
            Dim indexof_s = Me.Text.IndexOf("_rev")

            If indexof_s < 0 Then
                indexof_s = Me.Text.Count
            End If

            '---------- put fulfilled_qty in revision mr -------

            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Part_No", GetType(String))
            dimen_table.Columns.Add("qty_fulfilled", GetType(String))

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", Me.Text)
            cmd41.CommandText = "Select Part_No, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    dimen_table.Rows.Add(reader41(0).ToString, If(reader41(1) Is DBNull.Value, 0, reader41(1)))
                End While
            End If

            reader41.Close()


            '/////// start inserting revision to table /////////////

            Dim main_cmd As New MySqlCommand
            main_cmd.Parameters.AddWithValue("@mr_name", Me.Text.Remove(indexof_s, Me.Text.Count - indexof_s) & name_rev)
            main_cmd.Parameters.AddWithValue("@released_by", current_user)
            main_cmd.Parameters.AddWithValue("@job", open_job)
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
            For i = 0 To rev_grid.Rows.Count - 1

                If IsNumeric(rev_grid.Rows(i).Cells(8).Value) = True Then

                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@mr_name", Me.Text.Remove(indexof_s, Me.Text.Count - indexof_s) & name_rev)
                    Create_cmd6.Parameters.AddWithValue("@Part_No", If(rev_grid.Rows(i).Cells(0).Value Is Nothing, "", rev_grid.Rows(i).Cells(0).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Description", If(rev_grid.Rows(i).Cells(1).Value Is Nothing, "", rev_grid.Rows(i).Cells(1).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(rev_grid.Rows(i).Cells(2).Value Is Nothing, "", rev_grid.Rows(i).Cells(2).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(rev_grid.Rows(i).Cells(3).Value Is Nothing, "", rev_grid.Rows(i).Cells(3).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Vendor", If(rev_grid.Rows(i).Cells(4).Value Is Nothing, "", rev_grid.Rows(i).Cells(4).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Price", If(rev_grid.Rows(i).Cells(5).Value Is Nothing, "", rev_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")))
                    Create_cmd6.Parameters.AddWithValue("@Qty", If(rev_grid.Rows(i).Cells(8).Value Is Nothing, "", rev_grid.Rows(i).Cells(8).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@subtotal", If(IsNumeric(rev_grid.Rows(i).Cells(5).Value) = True, CType(rev_grid.Rows(i).Cells(5).Value, Double), 0) * If(IsNumeric(rev_grid.Rows(i).Cells(8).Value) = True, CType(rev_grid.Rows(i).Cells(8).Value, Double), 0))
                    Create_cmd6.Parameters.AddWithValue("@mfg_type", If(rev_grid.Rows(i).Cells(6).Value Is Nothing, "", rev_grid.Rows(i).Cells(6).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Assembly_name", If(rev_grid.Rows(i).Cells(10).Value Is Nothing, "", rev_grid.Rows(i).Cells(10).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Notes", If(rev_grid.Rows(i).Cells(11).Value Is Nothing, "", rev_grid.Rows(i).Cells(11).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Part_status", If(rev_grid.Rows(i).Cells(12).Value Is Nothing, "", rev_grid.Rows(i).Cells(12).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@Part_type", If(rev_grid.Rows(i).Cells(13).Value Is Nothing, "", rev_grid.Rows(i).Cells(13).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@full_panel", If(rev_grid.Rows(i).Cells(14).Value Is Nothing, "", rev_grid.Rows(i).Cells(14).Value.ToString))
                    Create_cmd6.Parameters.AddWithValue("@need_by_date", If(rev_grid.Rows(i).Cells(15).Value Is Nothing, "", rev_grid.Rows(i).Cells(15).Value.ToString))

                    For j = 0 To dimen_table.Rows.Count - 1
                        If String.Equals(dimen_table.Rows(j).Item(0), rev_grid.Rows(i).Cells(0).Value) = True Then
                            Create_cmd6.Parameters.AddWithValue("@qty_fulfilled", dimen_table.Rows(j).Item(1))
                            Exit For
                        End If
                    Next

                    If String.Equals(BOM_type, "old_BOM") = True Then
                        Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released, latest_r ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y', 'x' )"
                        Call clean_old(open_job, Me.Text.Remove(indexof_s, Me.Text.Count - indexof_s) & name_rev)
                    Else
                        Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y' )"
                    End If

                    '   Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y' )"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()
                End If

            Next

            ' If String.Equals(BOM_type, "Panel") = True Or String.Equals(BOM_type, "Assembly") = True Then
            Call BOM_types.Create_build_request(open_job)    '-- create build_request revision
            Call BOM_types.Create_MPL(open_job)   '--- create MPL
            '  End If


            If isit_member_of_BOM_p(Me.Text) = True Then
                Call Merge_and_release_MB(Me.Text)
            End If


            '////////////////////// ---------------------  notify -----------------------
            If enable_mess = True Then

                Dim mail_n As String : mail_n = "Material Request Revision for Project " & job_label.Text & "  has been released" & vbCrLf & vbCrLf _
             & "Material Request Revised: " & Me.Text.Remove(indexof_s, Me.Text.Count - indexof_s) & name_rev & vbCrLf _
             & "Material Request Revised: " & real_mr & vbCrLf _
             & "Revised by: " & current_user



                Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request Revision has been Released for Project " & job_label.Text, mail_n)
                Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request Revision has been Released for Project " & job_label.Text, mail_n)
                Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request Revision has been Released for Project " & job_label.Text, mail_n)
                Call Sent_mail.Sent_multiple_emails("General Management", "Material Request Revision has been Released for Project " & job_label.Text, mail_n)


                '--- sent email-------
                'add email addresses
                Dim emails_addr As New List(Of String)()

                'procurement
                emails_addr.Add("ecoy@atronixengineering.com")
                emails_addr.Add("fvargas@atronixengineering.com")
                emails_addr.Add("mmorris@atronixengineering.com")
                emails_addr.Add("sowens@atronixengineering.com")

                ''mfg
                emails_addr.Add("shenley@atronixengineering.com")
                emails_addr.Add("mowens@atronixengineering.com")

                ''inventory
                emails_addr.Add("dnix@atronixengineering.com")
                emails_addr.Add("dmoore@atronixengineering.com")




                '  For i = 0 To emails_addr.Count - 1

                Try
                    Dim e_mail As New MailMessage()
                    e_mail = New MailMessage()
                    e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                    For j = 0 To emails_addr.Count - 1
                        e_mail.To.Add(emails_addr.Item(j))
                    Next

                    e_mail.Subject = "Material Request Revision for Project " & job_label.Text & "  has been Released"
                    e_mail.IsBodyHtml = False
                    e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                    Smtp_Server.Send(e_mail)

                Catch error_t As Exception
                    MsgBox(error_t.ToString)
                End Try
                '   Next


                '-------- confirmation message -------------'
                Call send_confirmation_r(current_user, Me.Text, indexof_s, name_rev)
                '------------------------------------------

            End If
            '---------------------

            MessageBox.Show("Revision created succesfully!")

            job_label.Text = "Open Project:"
            PR_grid.Rows.Clear()
            Me.Text = "My Material Requests"
            isitreleased = False
            rev_mode = False
            Call Inventory_manage.General_inv_cal()   'recalculate inventory values
            Cursor.Current = Cursors.Default

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        TabControl1.TabPages.Remove(TabPage2)


    End Sub

    Private Sub rev_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles rev_grid.CellValueChanged

        If String.Equals(ComboBox1.Text, "Delta") = True Then

            For Each row As DataGridViewRow In rev_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(9).Value) = True And IsNumeric(row.Cells(7).Value)) Then
                    row.Cells(8).Value = CType(row.Cells(7).Value, Double) + CType(row.Cells(9).Value, Double)
                End If
            Next

        Else

            For Each row As DataGridViewRow In rev_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(8).Value) = True And IsNumeric(row.Cells(7).Value)) Then
                    row.Cells(9).Value = CType(row.Cells(8).Value, Double) - CType(row.Cells(7).Value, Double)
                End If
            Next
        End If

    End Sub

    Sub recal_rev()

        If String.Equals(ComboBox1.Text, "Delta") = True Then

            For Each row As DataGridViewRow In rev_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(9).Value) = True And IsNumeric(row.Cells(7).Value)) Then
                    row.Cells(8).Value = CType(row.Cells(7).Value, Double) + CType(row.Cells(9).Value, Double)
                End If
            Next

        Else

            For Each row As DataGridViewRow In rev_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(8).Value) = True And IsNumeric(row.Cells(7).Value)) Then
                    row.Cells(9).Value = CType(row.Cells(8).Value, Double) - CType(row.Cells(7).Value, Double)
                End If
            Next

        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged

        Dim mr_name As String = ""

        Dim date_created As String : date_created = ""
        Dim created_by As String : created_by = ""
        Dim released_by As String : released_by = ""
        Dim release_date As String : release_date = ""
        Dim last_modified As String : last_modified = ""

        date_c_l.Text = "Date Created:"
        created_by_l.Text = "Created By:"
        release_by_l.Text = "Released By:"
        release_d_l.Text = "Released Date:"
        last_m_l.Text = "Last Modified:"

        If Not ComboBox2.SelectedItem Is Nothing Then
            mr_name = ComboBox2.SelectedItem.ToString
        Else
            mr_name = ""
        End If

        '-------------------
        Try
            single_grid.Rows.Clear()
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name, Notes, Part_status, Part_type, need_by_date  from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    single_grid.Rows.Add(New String() {})
                    single_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    single_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    single_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    single_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                    single_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                    single_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                    single_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                    single_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                    single_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                    single_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                    single_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                    single_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                    single_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                    single_grid.Rows(i).Cells(13).Value = reader4(13).ToString

                    i = i + 1
                End While
            End If

            reader4.Close()

            '-------------------- dates ----------

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT Date_Created, created_by, released_by, release_date, last_modified from Material_Request.mr where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then

                While reader41.Read
                    date_c_l.Text = "Date Created: " & reader41(0).ToString
                    created_by_l.Text = "Created By: " & reader41(1).ToString
                    release_by_l.Text = "Released By: " & reader41(2).ToString
                    release_d_l.Text = "Released Date: " & reader41(3).ToString
                    last_modified = "Last Modified:" & reader41(4).ToString
                End While

            End If
            reader41.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TabControl1.TabPages.Remove(TabPage2)
        Inventory_manage.part_sel = ""
        rev_mode = False
    End Sub

    'Private Sub ImportMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportMaterialRequestToolStripMenuItem.Click

    '    '----------- Import Material Request from excel file -------------

    '    If isitreleased = False Then

    '        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
    '        '  Dim my_assemblies = New List(Of String)()
    '        index = 0

    '        If xlApp Is Nothing Then
    '            MessageBox.Show("Excel is not properly installed!!")
    '        Else

    '            'Try
    '            '    '--------------  add to device

    '            '    Dim cmd2 As New MySqlCommand
    '            '    cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
    '            '    cmd2.Connection = Login.Connection
    '            '    Dim reader2 As MySqlDataReader
    '            '    reader2 = cmd2.ExecuteReader

    '            '    If reader2.HasRows Then
    '            '        While reader2.Read
    '            '            my_assemblies.Add(reader2(0).ToString)
    '            '        End While
    '            '    End If

    '            '    reader2.Close()
    '            'Catch ex As Exception
    '            '    MessageBox.Show(ex.ToString)
    '            'End Try



    '            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*ods;"

    '            If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

    '                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
    '                PR_grid.Rows.Clear()


    '                Dim i As Integer : i = 2

    '                While (wb.ActiveSheet.Cells(i, 1).Value IsNot Nothing)

    '                    ' If my_assemblies.Contains(Trim(wb.Worksheets(1).Cells(i, 1).value.ToString)) = True Then
    '                    'Call break_part(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 7).value)

    '                    ' Else

    '                    Call Get_part_data(Trim(wb.Worksheets(1).Cells(i, 1).value), wb.Worksheets(1).Cells(i, 2).value, wb.Worksheets(1).Cells(i, 3).value, wb.Worksheets(1).Cells(i, 4).value, wb.Worksheets(1).Cells(i, 5).value, wb.Worksheets(1).Cells(i, 6).value, wb.Worksheets(1).Cells(i, 7).value, wb.Worksheets(1).Cells(i, 8).value, wb.Worksheets(1).Cells(i, 9).value, wb.Worksheets(1).Cells(i, 10).value, wb.Worksheets(1).Cells(i, 11).value, wb.Worksheets(1).Cells(i, 12).value, wb.Worksheets(1).Cells(i, 13).value)

    '                    '  End If
    '                    i = i + 1

    '                End While
    '                '---------------------------------------------
    '                'get latest cost
    '                For i = 0 To PR_grid.Rows.Count - 1
    '                    If PR_grid.Rows(i).DefaultCellStyle.BackColor = Color.CadetBlue Then
    '                        PR_grid.Rows(i).Cells(5).Value = "$" & PR_grid.Rows(i).Cells(5).Value
    '                    Else
    '                        PR_grid.Rows(i).Cells(5).Value = "$" & Form1.Get_Latest_Cost(Login.Connection, PR_grid.Rows(i).Cells(0).Value, PR_grid.Rows(i).Cells(4).Value)
    '                    End If
    '                Next
    '                '-------------------------------------------


    '                wb.Close(False)


    '                MessageBox.Show("Done")

    '            End If
    '        End If
    '    Else
    '        MessageBox.Show("No changes allowed in released version")
    '    End If
    'End Sub

    'Private Sub RemovePartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemovePartToolStripMenuItem.Click
    '    If PR_grid.SelectedRows.Count > 0 Then
    '        For Each r As DataGridViewRow In PR_grid.SelectedRows
    '            Try
    '                PR_grid.Rows.Remove(r)
    '            Catch ex As Exception
    '                MessageBox.Show("This row cannot be deleted")
    '            End Try
    '        Next
    '    Else
    '        MessageBox.Show("Select the row you want to delete. If you are in revision mode, please zero out the current qty, instead of removing the row")
    '    End If
    'End Sub

    '------- Similar to the one in form1. break apart the assembly
    Sub break_part(component As String, qt As Double)

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

                    PR_grid.Rows.Add(New String() {})  'add a new row
                    PR_grid.Rows(index).DefaultCellStyle.BackColor = Color.Tan
                    PR_grid.Rows(index).ReadOnly = True
                    PR_grid.Rows(index).Cells(0).Value = readerv(0).ToString
                    PR_grid.Rows(index).Cells(1).Value = readerv(1).ToString
                    PR_grid.Rows(index).Cells(3).Value = readerv(2).ToString
                    PR_grid.Rows(index).Cells(4).Value = readerv(3).ToString
                    PR_grid.Rows(index).Cells(6).Value = readerv(4) * qt
                    PR_grid.Rows(index).Cells(8).Value = readerv(5).ToString
                    PR_grid.Rows(index).Cells(9).Value = component
                    PR_grid.Rows(index).Cells(11).Value = readerv(6).ToString
                    PR_grid.Rows(index).Cells(12).Value = readerv(7).ToString

                    index = index + 1
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Get_part_data(part As String, desc As String, ada As String, mfg As String, vendor As String, cost As String, qty As Double, sub_t As String, mfg_t As String, assem As String, notes_t As String, status As String, type_t As String)

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@part", part)
            cmd4.CommandText = "Select Part_Description, Legacy_ADA_Number, Manufacturer, Primary_Vendor, MFG_type, Part_Status, Part_Type  from parts_table where Part_Name = @part"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            'if part is found copy the data from parts table else just copy data from excel file
            If reader4.HasRows Then
                While reader4.Read

                    PR_grid.Rows.Add(New String() {})
                    PR_grid.Rows(index).Cells(0).Value = part
                    PR_grid.Rows(index).Cells(1).Value = reader4(0)
                    PR_grid.Rows(index).Cells(2).Value = reader4(1)
                    PR_grid.Rows(index).Cells(3).Value = reader4(2)
                    PR_grid.Rows(index).Cells(4).Value = reader4(3)
                    PR_grid.Rows(index).Cells(6).Value = qty
                    PR_grid.Rows(index).Cells(8).Value = reader4(4)
                    PR_grid.Rows(index).Cells(11).Value = reader4(5)
                    PR_grid.Rows(index).Cells(12).Value = reader4(6)

                    index = index + 1

                End While
            Else
                PR_grid.Rows.Add(New String() {})
                PR_grid.Rows(index).DefaultCellStyle.BackColor = Color.CadetBlue
                PR_grid.Rows(index).Cells(0).Value = part
                PR_grid.Rows(index).Cells(1).Value = desc
                PR_grid.Rows(index).Cells(2).Value = ada
                PR_grid.Rows(index).Cells(3).Value = mfg
                PR_grid.Rows(index).Cells(4).Value = vendor
                PR_grid.Rows(index).Cells(5).Value = cost
                PR_grid.Rows(index).Cells(6).Value = qty
                PR_grid.Rows(index).Cells(7).Value = sub_t
                PR_grid.Rows(index).Cells(8).Value = mfg_t
                PR_grid.Rows(index).Cells(10).Value = notes_t
                PR_grid.Rows(index).Cells(11).Value = "Special Order"
                PR_grid.Rows(index).Cells(12).Value = type_t

                index = index + 1
            End If

            reader4.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Generate_MPL(mpl_name As String, isitnew As Boolean)

        'Create MPL from MR
        Dim my_assem As New Dictionary(Of String, String)
        Dim dimen_table As DataTable

        '------ create temp table -----
        dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Supplier", GetType(String))
        dimen_table.Columns.Add("Qty_needed", GetType(String))


        '------ collect all assemblies and qtys ------
        For i = 0 To PR_grid.Rows.Count - 1
            If PR_grid.Rows(i).Cells(9).Value Is Nothing = False Then
                If my_assem.ContainsKey(PR_grid.Rows(i).Cells(9).Value) = False Then
                    my_assem.Add(PR_grid.Rows(i).Cells(9).Value, PR_grid.Rows(i).Cells(9).Value)
                End If
            End If
        Next

        Try

            '--------------- get assemblies number ---------------
            Dim i As Integer : i = 0

            For Each kvp As KeyValuePair(Of String, String) In my_assem

                Dim n_asse As Integer : n_asse = 1
                Dim desc As String : desc = ""
                Dim Atronix_n As String : Atronix_n = "xxxx"
                Dim first_part As String : first_part = "xxxx"

                Dim cmd_a As New MySqlCommand
                cmd_a.Parameters.AddWithValue("@assembly", my_assem(kvp.Key))
                cmd_a.CommandText = "Select ADV_Number, Description from devices where Legacy_ADA_Number = @assembly"
                cmd_a.Connection = Login.Connection

                Dim reader_k As MySqlDataReader
                reader_k = cmd_a.ExecuteReader


                If reader_k.HasRows Then
                    While reader_k.Read
                        Atronix_n = reader_k(0)
                        desc = reader_k(1)
                    End While
                End If

                reader_k.Close()

                '--------------------------------- get parts ----------------------------------------------
                Dim cmd_pd As New MySqlCommand
                cmd_pd.Parameters.AddWithValue("@adv_n", Atronix_n)
                cmd_pd.CommandText = "SELECT Qty , Part_Name from adv where ADV_Number = @adv_n limit 1"
                cmd_pd.Connection = Login.Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_pd.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        n_asse = readerv(0)
                        first_part = readerv(1)
                    End While
                End If

                readerv.Close()

                '---------------------------- get real qty (qty/part qty) --------------------------------
                Dim z As Integer : z = 0

                For Each row As DataGridViewRow In PR_grid.Rows

                    If String.IsNullOrEmpty(row.Cells.Item("Column10").Value) = False And PR_grid.Rows(z).Cells(9).Value Is Nothing = False Then

                        If String.Compare(row.Cells.Item("Column10").Value.ToString, first_part) = 0 And String.Equals(PR_grid.Rows(z).Cells(9).Value, my_assem(kvp.Key)) = True Then
                            n_asse = If(IsNumeric(PR_grid.Rows(z).Cells(6).Value), PR_grid.Rows(z).Cells(6).Value, 0) / n_asse
                            Exit For
                        End If

                    End If

                    z = z + 1
                Next
                '-------------------------------------------
                '------  add to temp table

                dimen_table.Rows.Add(my_assem(kvp.Key), desc, "Atronix", n_asse)
                i = i + 1
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



        '--------------///////// generate MPL ////////----------------
        Try
            '----------------------- enter data to mpl --------------------------
            If isitnew = True Then
                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.AddWithValue("@mpl_name", mpl_name)
                main_cmd.Parameters.AddWithValue("@created_by", current_user)
                main_cmd.CommandText = "INSERT INTO Master_Packing_List.mpl(mpl_name, Date_Created , created_by) VALUES (@mpl_name, now(), @created_by)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()
            End If

            '----------------------- enter data (field) to mpl_data ---------------------
            For i = 0 To PR_grid.Rows.Count - 1

                If IsNumeric(PR_grid.Rows(i).Cells(6).Value) = True And PR_grid.Rows(i).Cells(0).Value Is Nothing = False And PR_grid.Rows(i).Cells(8).Value Is Nothing = False Then

                    If String.Equals(PR_grid.Rows(i).Cells(8).Value, "Field") = True And String.IsNullOrEmpty(PR_grid.Rows(i).Cells(9).Value) = True Then
                        Dim Create_cmd6 As New MySqlCommand
                        Create_cmd6.Parameters.Clear()
                        Create_cmd6.Parameters.AddWithValue("@mpl_name", mpl_name)
                        Create_cmd6.Parameters.AddWithValue("@Part_No", If(PR_grid.Rows(i).Cells(0).Value Is Nothing, "", PR_grid.Rows(i).Cells(0).Value))
                        Create_cmd6.Parameters.AddWithValue("@Description", If(PR_grid.Rows(i).Cells(1).Value Is Nothing, "", PR_grid.Rows(i).Cells(1).Value))
                        Create_cmd6.Parameters.AddWithValue("@Supplier", If(PR_grid.Rows(i).Cells(4).Value Is Nothing, "", PR_grid.Rows(i).Cells(4).Value))
                        Create_cmd6.Parameters.AddWithValue("@Qty_needed", If(PR_grid.Rows(i).Cells(6).Value Is Nothing, "", PR_grid.Rows(i).Cells(6).Value))

                        Create_cmd6.CommandText = "INSERT INTO Master_Packing_List.mpl_data(mpl_name, Part_No, Description, Supplier, Qty_needed) VALUES (@mpl_name, @Part_No, @Description, @Supplier, @Qty_needed)"
                        Create_cmd6.Connection = Login.Connection
                        Create_cmd6.ExecuteNonQuery()
                    End If

                End If

            Next

            '---------- enter Assemblies to mpl table ----------

            For i = 0 To dimen_table.Rows.Count - 1

                If String.Equals(dimen_table.Rows(i).Item(0).ToString, "") = False Then

                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@mpl_name", mpl_name)
                    Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table.Rows(i).Item(0).ToString Is Nothing, "", dimen_table.Rows(i).Item(0).ToString))
                    Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table.Rows(i).Item(1).ToString Is Nothing, "", dimen_table.Rows(i).Item(1).ToString))
                    Create_cmd6.Parameters.AddWithValue("@Supplier", If(dimen_table.Rows(i).Item(2).ToString Is Nothing, "", dimen_table.Rows(i).Item(2).ToString))
                    Create_cmd6.Parameters.AddWithValue("@Qty_needed", If(dimen_table.Rows(i).Item(3).ToString Is Nothing, "", dimen_table.Rows(i).Item(3).ToString))

                    Create_cmd6.CommandText = "INSERT INTO Master_Packing_List.mpl_data(mpl_name, Part_No, Description, Supplier, Qty_needed) VALUES (@mpl_name, @Part_No, @Description, @Supplier, @Qty_needed)"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()

                End If

            Next


            '----------------------------------------------------
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '///////////////////////////////////////////////////////////

    End Sub

    Private Sub ExportMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMaterialRequestToolStripMenuItem.Click
        'export  MR to excel file

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
                For i As Integer = 0 To PR_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = PR_grid.Columns(i).HeaderText
                    For j As Integer = 0 To PR_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = PR_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Dim valid_mr_name As String = Me.Text

                For Each c In Path.GetInvalidFileNameChars()
                    If valid_mr_name.Contains(c) Then
                        valid_mr_name = valid_mr_name.Replace(c, "")
                    End If
                Next

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Material_request_" & valid_mr_name & ".xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Request exported successfully!")
            End If
        End If
    End Sub

    Private Sub My_Material_r_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        TabControl1.TabPages.Remove(TabPage2)
        rev_mode = False
    End Sub

    '------ this subroutine check for special orders and add them to the main database if no error is found
    Function Special_order_check() As Boolean

        Special_order_check = False

        Try
            For i = 0 To PR_grid.Rows.Count - 1

                If PR_grid.Rows(i).IsNewRow Then Continue For

                If String.IsNullOrEmpty(PR_grid.Rows(i).Cells(13).Value) = True Or String.Equals(PR_grid.Rows(i).Cells(0).Value, "") = True Then

                    Dim exist_c As Boolean : exist_c = False
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part", PR_grid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "Select * from parts_table where Part_Name = @part"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    'if part is found update boolean
                    If reader4.HasRows Then
                        While reader4.Read
                            exist_c = True
                        End While
                    End If

                    reader4.Close()

                    Dim first_pass As Boolean : first_pass = False
                    Dim last_pass As Boolean : last_pass = False

                    If exist_c = False Then

                        '-- check if its a valid part
                        If String.IsNullOrEmpty(PR_grid.Rows(i).Cells(0).Value) = False And String.Equals(PR_grid.Rows(i).Cells(0).Value, "") = False Then 'check for valid part name 
                            If String.IsNullOrEmpty(PR_grid.Rows(i).Cells(3).Value) = False And String.Equals(PR_grid.Rows(i).Cells(3).Value, "") = False Then 'check for valid manuf
                                first_pass = True
                            End If
                        End If

                        ' check if MFG Type is valid
                        If first_pass = True Then

                            If String.Equals(PR_grid.Rows(i).Cells(8).Value, "Panel") = True Or String.Equals(PR_grid.Rows(i).Cells(8).Value, "Field") = True Or String.Equals(PR_grid.Rows(i).Cells(8).Value, "Assembly") = True Or String.Equals(PR_grid.Rows(i).Cells(8).Value, "Bulk") = True Then
                                last_pass = True

                                'start inserting data
                                Dim Create_cmd As New MySqlCommand
                                Create_cmd.Parameters.AddWithValue("@Part_Name", PR_grid.Rows(i).Cells(0).Value)
                                Create_cmd.Parameters.AddWithValue("@Manufacturer", PR_grid.Rows(i).Cells(3).Value)
                                Create_cmd.Parameters.AddWithValue("@Part_Description", PR_grid.Rows(i).Cells(1).Value)
                                Create_cmd.Parameters.AddWithValue("@Part_Status", "Special Order")
                                Create_cmd.Parameters.AddWithValue("@Part_Type", "Other")
                                Create_cmd.Parameters.AddWithValue("@Primary_Vendor", PR_grid.Rows(i).Cells(4).Value)
                                Create_cmd.Parameters.AddWithValue("@MFG_type", PR_grid.Rows(i).Cells(8).Value)


                                Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description,  Part_Status, Part_Type, Primary_Vendor, MFG_type) VALUES (@Part_Name, @Manufacturer, @Part_Description, @Part_Status, @Part_Type, @Primary_Vendor, @MFG_type)"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()

                                '--- insert into inventory
                                Dim Create_cmdi As New MySqlCommand
                                Create_cmdi.Parameters.AddWithValue("@part_name", PR_grid.Rows(i).Cells(0).Value)
                                Create_cmdi.Parameters.AddWithValue("@description", PR_grid.Rows(i).Cells(1).Value)
                                Create_cmdi.Parameters.AddWithValue("@manufacturer", PR_grid.Rows(i).Cells(3).Value)
                                Create_cmdi.Parameters.AddWithValue("@MFG_type", PR_grid.Rows(i).Cells(8).Value)

                                Create_cmdi.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, 0,0,0,0)"
                                Create_cmdi.Connection = Login.Connection
                                Create_cmdi.ExecuteNonQuery()

                                '----------- insert vendor and cost -----
                                If IsNumeric(PR_grid.Rows(i).Cells(5).Value) = True And String.IsNullOrEmpty(PR_grid.Rows(i).Cells(4).Value) = False And String.Equals(PR_grid.Rows(i).Cells(4).Value, "") = False Then
                                    Dim main_cmd2 As New MySqlCommand
                                    main_cmd2.Parameters.AddWithValue("@Part_Name", PR_grid.Rows(i).Cells(0).Value)
                                    main_cmd2.Parameters.AddWithValue("@Vendor_Name", PR_grid.Rows(i).Cells(4).Value)
                                    main_cmd2.Parameters.AddWithValue("@Cost", If(IsNumeric(PR_grid.Rows(i).Cells(5).Value.ToString.Replace("$", "")) = True, PR_grid.Rows(i).Cells(5).Value.ToString.Replace("$", ""), 0))
                                    main_cmd2.CommandText = "INSERT INTO vendors_table(Part_Name, Vendor_Name, Cost, Purchase_Date) VALUES (@Part_Name, @Vendor_Name, @Cost, now())"
                                    main_cmd2.Connection = Login.Connection
                                    main_cmd2.ExecuteNonQuery()
                                End If

                            End If
                        End If

                        If last_pass = False Then
                            PR_grid.Rows(i).DefaultCellStyle.BackColor = Color.Firebrick
                            Special_order_check = False
                        Else
                            If i Mod 2 = 0 Then
                                PR_grid.Rows(i).DefaultCellStyle.BackColor = Color.WhiteSmoke
                            Else
                                PR_grid.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
                            End If

                            Special_order_check = True
                        End If

                    Else  'if part exist
                        Special_order_check = True

                    End If

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    'Private Sub ViewProjectPackingListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewProjectPackingListToolStripMenuItem.Click
    '    '-- GENERATE A Packing List -----
    '    If String.Equals("My Material Requests", Me.Text) = False Then
    '        Bom_packing.Visible = True
    '        Call Packing_List(Me.Text, Bom_packing.packing_grid)
    '        Bom_packing.bom_label.Text = "Packing List for BOM: " & Me.Text
    '        Bom_packing.job_label.Text = job_label.Text
    '    End If
    'End Sub

    'Private Sub ViewBuildRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewBuildRequestToolStripMenuItem.Click
    '    '---- GENERATE BUILD REQUEST ------
    '    If String.Equals("My Material Requests", Me.Text) = False Then
    '        BOM_build.Visible = True
    '        Call Build_Request(Me.Text, BOM_build.build_grid)
    '        BOM_build.bom_label.Text = "Build Request for BOM:  " & Me.Text
    '        BOM_build.job_label.Text = job_label.Text
    '    End If
    'End Sub

    '--------- This sub generates a Packing List and displayed it in a table
    Sub Packing_List(mr_name As String, my_grid As DataGridView)

        my_grid.Rows.Clear()

        Dim my_assem = New List(Of String)()
        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("description", GetType(String))
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("need_by_date", GetType(String))
        datatable.Columns.Add("isitpanel", GetType(String))
        datatable.Columns.Add("mfg_type", GetType(String))


        Try
            '--------------  add to device
            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assem.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()

            '---------------------------------------
            '------  enter data in datatable
            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@mr_name", mr_name)
            cmd3.CommandText = "SELECT Part_No, Description, Qty, need_by_date, full_panel, mfg_type from Material_Request.mr_data where mr_name = @mr_name"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    datatable.Rows.Add(reader3(0).ToString, reader3(1).ToString, reader3(2).ToString, reader3(3).ToString, reader3(4).ToString, reader3(5).ToString)
                End While
            End If

            reader3.Close()
            '--------------------------------------------

            For i = 0 To datatable.Rows.Count - 1
                If my_assem.Contains(datatable.Rows(i).Item(0).ToString) = True Or String.Equals("Yes", datatable.Rows(i).Item(4).ToString) = True Or String.Equals("Field", datatable.Rows(i).Item(5).ToString) = True Then
                    my_grid.Rows.Add(New String() {datatable.Rows(i).Item(2), datatable.Rows(i).Item(0), datatable.Rows(i).Item(1), datatable.Rows(i).Item(3)})
                End If
            Next


            '----------- add panels in full_panels table --------------------
            Dim cmd13 As New MySqlCommand
            cmd13.Parameters.AddWithValue("@mr_name", mr_name)
            cmd13.CommandText = "SELECT panel_name, panel_desc, qty, need_by_date from Material_Request.full_panels where mr_name = @mr_name"
            cmd13.Connection = Login.Connection
            Dim reader13 As MySqlDataReader
            reader13 = cmd13.ExecuteReader

            If reader13.HasRows Then
                While reader13.Read
                    my_grid.Rows.Add(reader13(2).ToString, reader13(0).ToString, reader13(1).ToString, reader13(3).ToString)
                End While
            End If

            reader13.Close()

            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    '--------- This sub generates a Build Request and displayed it in a table
    Sub Build_Request(mr_name As String, my_grid As DataGridView)

        my_grid.Rows.Clear()

        Dim my_assem = New List(Of String)()
        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("description", GetType(String))
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("need_by_date", GetType(String))
        datatable.Columns.Add("isitpanel", GetType(String))


        Try
            '--------------  add to device
            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assem.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()

            '------------------------------------------
            '------  enter data in datatable ----------
            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@mr_name", mr_name)
            cmd3.CommandText = "SELECT Part_No, Description, Qty, need_by_date, full_panel from Material_Request.mr_data where mr_name = @mr_name"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    datatable.Rows.Add(reader3(0).ToString, reader3(1).ToString, reader3(2).ToString, reader3(3).ToString, reader3(4).ToString)
                End While
            End If

            reader3.Close()
            '--------------------------------------------

            For i = 0 To datatable.Rows.Count - 1

                If my_assem.Contains(datatable.Rows(i).Item(0).ToString) = True Or String.Equals("Yes", datatable.Rows(i).Item(4).ToString) = True Then
                    my_grid.Rows.Add(New String() {datatable.Rows(i).Item(0), datatable.Rows(i).Item(1), datatable.Rows(i).Item(2), datatable.Rows(i).Item(3)})
                End If
            Next

            '-------------- add panels in full_panels table -----------------
            Dim cmd13 As New MySqlCommand
            cmd13.Parameters.AddWithValue("@mr_name", mr_name)
            cmd13.CommandText = "SELECT panel_name, panel_desc, qty, need_by_date from Material_Request.full_panels where mr_name = @mr_name"
            cmd13.Connection = Login.Connection
            Dim reader13 As MySqlDataReader
            reader13 = cmd13.ExecuteReader

            If reader13.HasRows Then
                While reader13.Read
                    my_grid.Rows.Add(reader13(0).ToString, reader13(1).ToString, reader13(2).ToString, reader13(3).ToString)
                End While
            End If

            reader13.Close()

            '-------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub ChangeShipppingAddressToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeShipppingAddressToolStripMenuItem.Click
        Shipping_Edit.ShowDialog()
    End Sub

    'Private Sub general_date_TextChanged(sender As Object, e As EventArgs) Handles general_date.TextChanged

    '    'enter a general need by date in your BOM
    '    If isitreleased = False Then
    '        If PR_grid.Rows.Count > 0 Then
    '            For i = 0 To PR_grid.Rows.Count - 1
    '                If PR_grid.Rows(i).IsNewRow Then Continue For
    '                PR_grid.Rows(i).Cells(14).Value = general_date.Text
    '            Next
    '        End If
    '    End If
    'End Sub

    'Private Sub MFGViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MFGViewToolStripMenuItem.Click
    '    '-- hide some columns fpr MFG people

    '    PR_grid.Columns(2).Visible = False
    '    PR_grid.Columns(5).Visible = False
    '    PR_grid.Columns(7).Visible = False
    '    PR_grid.Columns(8).Visible = False
    '    PR_grid.Columns(9).Visible = False
    '    PR_grid.Columns(10).Visible = False
    '    PR_grid.Columns(11).Visible = False
    '    PR_grid.Columns(12).Visible = False
    '    PR_grid.Columns(13).Visible = False
    '    PR_grid.Columns(15).Visible = True
    'End Sub

    'Private Sub SeeAllColumnsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeeAllColumnsToolStripMenuItem.Click
    '    '-- show all columns ----

    '    PR_grid.Columns(2).Visible = True
    '    PR_grid.Columns(5).Visible = True
    '    PR_grid.Columns(7).Visible = True
    '    PR_grid.Columns(8).Visible = True
    '    PR_grid.Columns(9).Visible = True
    '    PR_grid.Columns(10).Visible = True
    '    PR_grid.Columns(11).Visible = True
    '    PR_grid.Columns(12).Visible = True
    '    ' PR_grid.Columns(13).Visible = True
    '    PR_grid.Columns(15).Visible = False

    'End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '-- compare select mr
        If Not mrbox1.SelectedItem Is Nothing And Not mrbox2.SelectedItem Is Nothing Then

            compare_grid.Rows.Clear()

            Dim total_mr As New Dictionary(Of String, String)
            '-- store mr1 parts
            Dim dimen_table1 = New DataTable
            dimen_table1.Columns.Add("Part_No", GetType(String))
            dimen_table1.Columns.Add("Description", GetType(String))

            '-- store mr2 parts
            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("Description", GetType(String))

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                cmd4.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        dimen_table1.Rows.Add(reader4(0).ToString, reader4(1).ToString)
                    End While
                End If

                reader4.Close()

                '--- mr2
                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                cmd41.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        dimen_table2.Rows.Add(reader41(0).ToString, reader41(1).ToString)
                    End While
                End If

                reader41.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            For i = 0 To dimen_table1.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table1.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table1.Rows(i).Item(0), dimen_table1.Rows(i).Item(1))
                End If
            Next

            For i = 0 To dimen_table2.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table2.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table2.Rows(i).Item(0), dimen_table2.Rows(i).Item(1))
                End If
            Next

            '--- add to comp_Grid
            For Each kvp As KeyValuePair(Of String, String) In total_mr.ToArray
                compare_grid.Rows.Add(New String() {kvp.Key, total_mr(kvp.Key)})
            Next

            For i = 0 To compare_grid.Rows.Count - 1
                Try
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.Clear()
                    cmd41.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                    cmd41.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd41.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            compare_grid.Rows(i).Cells(2).Value = reader41(0).ToString
                            compare_grid.Rows(i).Cells(4).Value = reader41(1).ToString
                        End While
                    End If

                    reader41.Close()

                    Dim cmd42 As New MySqlCommand
                    cmd42.Parameters.Clear()
                    cmd42.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                    cmd42.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd42.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd42.Connection = Login.Connection
                    Dim reader42 As MySqlDataReader
                    reader42 = cmd42.ExecuteReader

                    If reader42.HasRows Then
                        While reader42.Read
                            compare_grid.Rows(i).Cells(3).Value = reader42(0).ToString
                            compare_grid.Rows(i).Cells(5).Value = reader42(1).ToString
                        End While
                    End If

                    reader42.Close()

                    compare_grid.Rows(i).Cells(6).Value = If(IsNumeric(compare_grid.Rows(i).Cells(5).Value), compare_grid.Rows(i).Cells(5).Value, 0) - If(IsNumeric(compare_grid.Rows(i).Cells(4).Value), compare_grid.Rows(i).Cells(4).Value, 0)

                    If compare_grid.Rows(i).Cells(6).Value <> 0 Then
                        compare_grid.Rows(i).Cells(6).Style.BackColor = Color.CadetBlue
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Next



        End If
    End Sub

    ' Private Sub UploadDrawingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UploadDrawingToolStripMenuItem.Click
    '---- upload pdf 

    'If String.Equals(Me.Text, "My Material Requests") = False Then

    '    If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

    '        Dim DestPath As String = "O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\"

    '        Dim draw_desc As String = InputBox("Enter Description of Drawing", "Drawing Description", "Drawing Description")

    '        CopyDirectory(DestPath, OpenFileDialog1.FileName, open_job, draw_desc)

    '    End If
    'End If
    ' End Sub

    Sub CopyDirectory(destPath As String, filename_s As String, job As String, draw_desc As String)

        '-- copy file to APL drawing folder


        Dim path_s As String : path_s = Path.GetDirectoryName(filename_s)
        Dim only_f As String : only_f = Path.GetFileName(filename_s)

        ' MessageBox.Show(filename_s)
        ' MessageBox.Show(destPath & "Drawing_" & Me.Text.Replace("/", ""))

        Dim desc_path As String : desc_path = destPath & "Drawing_" & only_f.Replace("/", "")

        If File.Exists(desc_path) = True Then
            MessageBox.Show("A Drawing with the same name already exist")
        Else
            File.Copy(filename_s, desc_path)

            Try
                Dim Create_cmd6 As New MySqlCommand
                Create_cmd6.Parameters.Clear()
                Create_cmd6.Parameters.AddWithValue("@job", job)
                Create_cmd6.Parameters.AddWithValue("@draw_description", draw_desc)
                Create_cmd6.Parameters.AddWithValue("@file_name", "Drawing_" & only_f.Replace("/", ""))
                Create_cmd6.Parameters.AddWithValue("@uploaded_by", current_user)

                Create_cmd6.CommandText = "INSERT INTO management.draw_pdf(job, draw_description, file_name , date_upload, uploaded_by) VALUES (@job, @draw_description, @file_name, now(), @uploaded_By)"
                Create_cmd6.Connection = Login.Connection
                Create_cmd6.ExecuteNonQuery()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            MessageBox.Show("Drawing upload successfully")
        End If


    End Sub

    Private Sub AddPanelsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddPanelsToolStripMenuItem.Click
        If String.Equals(Me.Text, "My Material Requests") = False Then

            Dim has_map As Boolean : has_map = False
            Dim mr_name As String : mr_name = ""

            Try
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@mr_name", Me.Text)
                check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name and BOM_type not like 'old_BOM' and BOM_type not like 'ASM'"

                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    has_map = True
                End If

                reader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            If has_map = True Then

                Edit_BOM_Panels.Text = Me.Text
                Edit_BOM_Panels.ShowDialog()

            Else
                MessageBox.Show("Sorry, this is an old type of BOM!")

            End If

        End If
    End Sub

    '---- check if BOM has assemblies
    Function does_it_has_assem(mr_name As String) As Boolean

        does_it_has_assem = False

        Dim my_assem = New List(Of String)()

        Try
            '---- assemblies list ----
            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assem.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()
            '-------------------------

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT Part_No from Material_Request.mr_data where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    If my_assem.Contains(reader41(0).ToString) = True Then
                        does_it_has_assem = True
                        Exit While
                    End If
                End While
            End If

            reader41.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function


    Function job_with_build(job As String) As Boolean

        job_with_build = False

        Dim bom_list = New List(Of String)()

        '------- get BOM ------
        '------ get all BOM of the job 
        Dim n_bom As Double : n_bom = 0
        Dim check_cmd As New MySqlCommand
        check_cmd.Parameters.AddWithValue("@job", job)
        check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

        check_cmd.Connection = Login.Connection
        check_cmd.ExecuteNonQuery()

        Dim reader As MySqlDataReader
        reader = check_cmd.ExecuteReader

        If reader.HasRows Then
            While reader.Read
                n_bom = reader(0)
            End While
        End If

        reader.Close()

        For i = 1 To n_bom

            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.Clear()
            check_cmd1.Parameters.AddWithValue("@job", job)
            check_cmd1.Parameters.AddWithValue("@id_bom", i)
            check_cmd1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then

                While reader1.Read
                    If bom_list.Contains(reader1(0).ToString) = False Then
                        bom_list.Add(reader1(0).ToString)
                    End If
                End While
            End If

            reader1.Close()
        Next
        '---------------------------------------------------------

        For i = 0 To bom_list.Count - 1

            If does_it_has_assem(bom_list.Item(i).ToString) = True Then
                job_with_build = True
                Exit For
            End If

            If Assign_box.does_it_has_panels(bom_list.Item(i).ToString) = True Then
                job_with_build = True
                Exit For
            End If
        Next

    End Function

    Private Sub OpenSavedRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenSavedRevisionToolStripMenuItem.Click
        If isitreleased = True Then
            If rev_mode = False Then
                Open_Revision.ShowDialog()
            Else
                MessageBox.Show("Close Revision Mode first")
            End If

        Else
            MessageBox.Show("Not a Released BOM!")
        End If
    End Sub

    Private Sub SaveRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveRevisionToolStripMenuItem.Click

        If rev_mode = True Then
            SAVE_revision.ShowDialog()
        Else
            MessageBox.Show("You are not in Revision Mode!")
        End If
    End Sub

    Private Sub BOMManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BOMManagerToolStripMenuItem.Click

        '------ open BOM manager

        Dim has_map As Boolean : has_map = False
        Dim mr_name As String : mr_name = ""

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@mr_name", Me.Text)
            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name and BOM_type not like 'old_BOM' and BOM_type not like 'ASM'"

            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                has_map = True
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        If has_map = True Then

            Try
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", open_job)
                check_cmd.CommandText = "select mr_name from Material_Request.mr where job = @job and BOM_type = 'MB' order by release_date desc limit 1"
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        mr_name = reader(0).ToString
                    End While
                End If

                reader.Close()

                If String.Equals(mr_name, "") = False Then

                    Package_BOM_tree.Load_BOM_map(mr_name, open_job)
                    Package_BOM_tree.Visible = True
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Else
            MessageBox.Show("No BOM Package associated with this Master BOM")
        End If


    End Sub

    Function isit_MB(mr_name As String) As Boolean
        '-- return true if its a MB new file

        isit_MB = False

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name and BOM_type = 'MB'"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                isit_MB = True
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function

    Function isit_member_of_BOM_p(mr_name As String)

        '---- returns true if this bom is member of a BOM Package. 

        isit_member_of_BOM_p = False

        Try
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd.CommandText = "select * from Material_Request.mr where mr_name = @mr_name and MBOM is not null"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                isit_member_of_BOM_p = True
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function

    Sub Merge_and_release_MB(mr_name As String)
        '--create a new MB with all the latest revisions of its components and name it MB_Revx

        Dim MB As String : MB = ""
        Dim new_MB As String : new_MB = ""

        Dim unmerge_table = New DataTable
        unmerge_table.Columns.Add("Part_No", GetType(String))
        unmerge_table.Columns.Add("Description", GetType(String))
        unmerge_table.Columns.Add("Manufacturer", GetType(String))
        unmerge_table.Columns.Add("Vendor", GetType(String))
        unmerge_table.Columns.Add("Price", GetType(String))
        unmerge_table.Columns.Add("Qty", GetType(String))
        unmerge_table.Columns.Add("Subtotal", GetType(String))
        unmerge_table.Columns.Add("mfg_type", GetType(String))
        unmerge_table.Columns.Add("Part_status", GetType(String))
        unmerge_table.Columns.Add("Part_Type", GetType(String))
        unmerge_table.Columns.Add("Need_by_date", GetType(String))
        unmerge_table.Columns.Add("Notes", GetType(String))


        '-- panel bom add
        Dim all_latest = New List(Of String)()

        Try

            '-------- get MBOM ------------
            Dim cmdm As New MySqlCommand
            cmdm.Parameters.AddWithValue("@job", open_job)
            cmdm.Parameters.AddWithValue("@mr_name", mr_name)
            cmdm.CommandText = "SELECT MBOM from Material_Request.mr where mr_name = @mr_name  and job = @job"
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

            Dim n_bom As Double : n_bom = 0
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", open_job)
            check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    n_bom = reader(0)
                End While
            End If

            reader.Close()
            '---------------------------------------------------


            For i = 2 To n_bom

                Dim check_cmd1 As New MySqlCommand
                check_cmd1.Parameters.Clear()
                check_cmd1.Parameters.AddWithValue("@job", open_job)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    While reader1.Read
                        all_latest.Add(reader1(0).ToString)
                    End While
                End If

                reader1.Close()
            Next

            '------enter data----
            For i = 0 To all_latest.Count - 1

                Dim qty_pa As Integer : qty_pa = 1

                '--- if its a panel bom then get qty if is null make value 1 ---
                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", all_latest.Item(i).ToString)
                cmd419.CommandText = "SELECT Panel_qty , BOM_type from Material_Request.mr where mr_name = @mr_name"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read
                        If String.Equals("Panel", reader419(1)) = True Then

                            qty_pa = If(reader419(0) Is DBNull.Value, 1, reader419(0))

                        End If
                    End While
                End If

                reader419.Close()
                '----------------------------------------------------------



                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.Clear()
                cmd41.Parameters.AddWithValue("@mr_name", all_latest.Item(i).ToString)
                cmd41.CommandText = "SELECT  Part_No, Description, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Part_status, Part_type, Notes, need_by_date from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        unmerge_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString, reader41(3).ToString, reader41(4).ToString, reader41(5).ToString * qty_pa, reader41(6).ToString, reader41(7).ToString, reader41(8).ToString, reader41(9).ToString, reader41(10).ToString, reader41(11).ToString)
                    End While
                End If

                reader41.Close()
            Next



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '------------- merge the parts of same name ---------
        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Subtotal", GetType(String))
        dimen_table.Columns.Add("mfg_type", GetType(String))
        dimen_table.Columns.Add("Part_status", GetType(String))
        dimen_table.Columns.Add("Part_Type", GetType(String))
        dimen_table.Columns.Add("Need_by_date", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))



        For i = 0 To unmerge_table.Rows.Count - 1
            '-- combine

            Dim found_t As Boolean : found_t = False

            For j = 0 To dimen_table.Rows.Count - 1

                If String.IsNullOrEmpty(unmerge_table.Rows(i).Item(0)) = False Then
                    If String.Equals(unmerge_table.Rows(i).Item(0).ToString, dimen_table.Rows(j).Item(0)) = True Then
                        dimen_table.Rows(j).Item(5) = If(IsNumeric(dimen_table.Rows(j).Item(5)), dimen_table.Rows(j).Item(5), 0) + CType(unmerge_table.Rows(i).Item(5).ToString, Double)
                        found_t = True
                        Exit For
                    End If
                End If

            Next

            If found_t = False Then
                dimen_table.Rows.Add(unmerge_table.Rows(i).Item(0).ToString, unmerge_table.Rows(i).Item(1).ToString, unmerge_table.Rows(i).Item(2).ToString, unmerge_table.Rows(i).Item(3).ToString, unmerge_table.Rows(i).Item(4).ToString, unmerge_table.Rows(i).Item(5).ToString, unmerge_table.Rows(i).Item(6).ToString, unmerge_table.Rows(i).Item(7).ToString, unmerge_table.Rows(i).Item(8).ToString, unmerge_table.Rows(i).Item(9).ToString, unmerge_table.Rows(i).Item(10).ToString, unmerge_table.Rows(i).Item(11).ToString)
            End If
        Next

        '-- all the data of the new MB revision is in dimen_table, now create a revision and insert the data
        ' ///////////////////////////////////////////////////////////////////////////////////////////////////////'

        Try
            Dim name_rev As String = "_rev"
            Dim iz As Integer : iz = 0 'counter

            Dim shipping As String : shipping = ""
            Dim Date_Created As Date
            Dim created_by As String : created_by = "unknown"
            Dim id_bom As String : id_bom = ""

            ' Dim need_date As String : need_d
            Dim Panel_name As String : Panel_name = ""
            Dim Panel_qty As String : Panel_qty = 1
            Dim Panel_desc As String : Panel_desc = ""
            Dim need_date As String : need_date = ""
            Dim BOM_type As String : BOM_type = ""
            Dim MBOM As String : MBOM = ""


            '---- get id_bom

            Dim id_bom_r As String : id_bom_r = ""
            Dim cmd411 As New MySqlCommand
            cmd411.Parameters.AddWithValue("@mr_name", MB)
            cmd411.CommandText = "SELECT id_bom from Material_Request.mr where mr_name = @mr_name"
            cmd411.Connection = Login.Connection
            Dim reader411 As MySqlDataReader
            reader411 = cmd411.ExecuteReader

            If reader411.HasRows Then
                While reader411.Read
                    id_bom_r = reader411(0).ToString
                End While
            End If

            reader411.Close()


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", open_job)
            cmd4.Parameters.AddWithValue("@id_bom", id_bom_r)
            cmd4.CommandText = "Select shipping_ad ,Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_type, MBOM from Material_Request.mr where released = 'Y' and job = @job and id_bom = @id_bom order by release_date"
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
                    Panel_qty = If(reader4(6) Is Nothing, 1, reader4(6).ToString)
                    Panel_desc = If(reader4(7) Is Nothing, "", reader4(7).ToString)


                    iz = iz + 1
                End While
            End If

            reader4.Close()

            name_rev = name_rev & iz  'last part of name of file ex: filename_revx
            Dim indexof_s = MB.IndexOf("_rev")

            If indexof_s < 0 Then
                indexof_s = MB.Count
            End If

            '---------- put fulfilled_qty in revision mr -------

            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("qty_fulfilled", GetType(String))

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", MB)
            cmd41.CommandText = "Select Part_No, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    dimen_table2.Rows.Add(reader41(0).ToString, If(reader41(1) Is DBNull.Value, 0, reader41(1)))
                End While
            End If

            reader41.Close()

            new_MB = MB.Remove(indexof_s, MB.Count - indexof_s) & name_rev
            real_mr = new_MB

            '/////// start inserting revision to table /////////////

            Dim main_cmd As New MySqlCommand
            main_cmd.Parameters.AddWithValue("@mr_name", MB.Remove(indexof_s, MB.Count - indexof_s) & name_rev)
            main_cmd.Parameters.AddWithValue("@released_by", current_user)
            main_cmd.Parameters.AddWithValue("@job", open_job)
            main_cmd.Parameters.AddWithValue("@Date_Created", Date_Created)
            main_cmd.Parameters.AddWithValue("@created_by", created_by)
            main_cmd.Parameters.AddWithValue("@id_bom", id_bom)
            main_cmd.Parameters.AddWithValue("@shipping_ad", shipping)
            main_cmd.Parameters.AddWithValue("@MBOM", "")
            main_cmd.Parameters.AddWithValue("@BOM_type", "MB")

            main_cmd.Parameters.AddWithValue("@need_date", need_date)
            main_cmd.Parameters.AddWithValue("@Panel_name", Panel_name)
            main_cmd.Parameters.AddWithValue("@Panel_desc", Panel_desc)
            main_cmd.Parameters.AddWithValue("@Panel_qty", Panel_qty)

            main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, released, released_by, release_date, job,  Date_Created, created_by, id_bom, shipping_ad, BOM_type, MBOM, need_date, Panel_name, Panel_desc, Panel_qty) VALUES (@mr_name,'Y', @released_by, now(), @job, @Date_Created, @created_by, @id_bom, @shipping_ad, @BOM_type ,@MBOM, @need_date, @Panel_name, @Panel_desc, @Panel_qty)"
            main_cmd.Connection = Login.Connection
            main_cmd.ExecuteNonQuery()



            '-------- enter data to mr_data
            For i = 0 To dimen_table.Rows.Count - 1

                Dim Create_cmd6 As New MySqlCommand
                Create_cmd6.Parameters.Clear()
                Create_cmd6.Parameters.AddWithValue("@mr_name", MB.Remove(indexof_s, MB.Count - indexof_s) & name_rev)
                Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table.Rows(i).Item(0).ToString Is Nothing, "", dimen_table.Rows(i).Item(0).ToString))
                Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table.Rows(i).Item(1).ToString Is Nothing, "", dimen_table.Rows(i).Item(1).ToString))
                Create_cmd6.Parameters.AddWithValue("@ADA_Number", "")
                Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(dimen_table.Rows(i).Item(2).ToString Is Nothing, "", dimen_table.Rows(i).Item(2).ToString))
                Create_cmd6.Parameters.AddWithValue("@Vendor", If(dimen_table.Rows(i).Item(3).ToString Is Nothing, "", dimen_table.Rows(i).Item(3).ToString))
                Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(4).ToString Is Nothing, "0", dimen_table.Rows(i).Item(4).ToString.Replace("$", "")))
                Create_cmd6.Parameters.AddWithValue("@Qty", If(dimen_table.Rows(i).Item(5).ToString Is Nothing, "0", dimen_table.Rows(i).Item(5).ToString))
                Create_cmd6.Parameters.AddWithValue("@subtotal", If(IsNumeric(dimen_table.Rows(i).Item(6).ToString) = True, CType(dimen_table.Rows(i).Item(6).ToString, Double), 0) * If(IsNumeric(dimen_table.Rows(i).Item(6).ToString) = True, CType(dimen_table.Rows(i).Item(6).ToString, Double), 0))
                Create_cmd6.Parameters.AddWithValue("@mfg_type", If(dimen_table.Rows(i).Item(7).ToString Is Nothing, "", dimen_table.Rows(i).Item(7).ToString))
                Create_cmd6.Parameters.AddWithValue("@Assembly_name", "")
                Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table.Rows(i).Item(8).ToString Is Nothing, "", dimen_table.Rows(i).Item(8).ToString))
                Create_cmd6.Parameters.AddWithValue("@Part_type", If(dimen_table.Rows(i).Item(9).ToString Is Nothing, "", dimen_table.Rows(i).Item(9).ToString))
                Create_cmd6.Parameters.AddWithValue("@full_panel", "")
                Create_cmd6.Parameters.AddWithValue("@need_by_date", If(dimen_table.Rows(i).Item(10).ToString Is Nothing, "", dimen_table.Rows(i).Item(10).ToString))
                Create_cmd6.Parameters.AddWithValue("@Notes", If(dimen_table.Rows(i).Item(11).ToString Is Nothing, "", dimen_table.Rows(i).Item(11).ToString))

                For j = 0 To dimen_table2.Rows.Count - 1
                    If String.Equals(dimen_table2.Rows(j).Item(0), dimen_table.Rows(i).Item(0).ToString) = True Then
                        Create_cmd6.Parameters.AddWithValue("@qty_fulfilled", dimen_table2.Rows(j).Item(1))
                        Exit For
                    End If
                Next


                Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released, latest_r ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y', 'x' )"
                Create_cmd6.Connection = Login.Connection
                Create_cmd6.ExecuteNonQuery()
            Next

            Call clean_x(open_job, MB.Remove(indexof_s, MB.Count - indexof_s) & name_rev)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '------------------------------------------------------------------------------------------------------------

        '-updating MBOM for the latest BOM members
        For i = 0 To all_latest.Count - 1

            Dim Create_cmd2 As New MySqlCommand
            Create_cmd2.Parameters.Clear()
            Create_cmd2.Parameters.AddWithValue("@mr_name", all_latest.Item(i).ToString)
            Create_cmd2.Parameters.AddWithValue("@MB", new_MB)

            Create_cmd2.CommandText = "UPDATE Material_Request.mr SET MBOM = @MB where mr_name = @mr_name"
            Create_cmd2.Connection = Login.Connection
            Create_cmd2.ExecuteNonQuery()
        Next


        '------------------------------------------------------------
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
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
                For i As Integer = 0 To compare_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = compare_grid.Columns(i).HeaderText
                    For j As Integer = 0 To compare_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = compare_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Compare_data_" & ".xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Table exported successfully!")
            End If
        End If
    End Sub

    Private Sub OpenMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenMaterialRequestToolStripMenuItem.Click
        If Log_out = True Then
            Me.Visible = False
            BOM_init.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

        If TabControl1.SelectedTab Is TabPage1 Then

            If PR_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(PR_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage3 Then
            If single_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(single_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage4 Then
            If compare_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(compare_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage2 Then
            If rev_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(rev_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If
        End If
    End Sub

    Private Sub ChangeNeedByDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeNeedByDateToolStripMenuItem.Click
        '--- open need by date --------
        Need_by_date.ShowDialog()
    End Sub

    Private Sub FeatureCodesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FeatureCodesToolStripMenuItem.Click
        If String.Equals(Me.Text, "My Material Requests") = False And rev_mode = True Then

            BOM_types.feature_s = True
            Feature_sel.ShowDialog()

        Else
            MessageBox.Show("You are not in revision mode")
        End If
    End Sub


    Sub clean_x(job As String, mr_name As String)
        ' this sub unmarked the mr_data with x's
        Try
            Dim mr_list As New List(Of String)()

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", job)
            cmd4.Parameters.AddWithValue("@mr", mr_name)
            cmd4.CommandText = "select mr_name from Material_Request.mr where job = @job and BOM_type = 'MB' and mr_name not like @mr"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    mr_list.Add(reader4(0).ToString)
                End While
            End If

            reader4.Close()

            For i = 0 To mr_list.Count - 1

                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@mr_name", mr_list.Item(i).ToString)

                Create_cmd.CommandText = "UPDATE Material_Request.mr_data SET latest_r = null where mr_name = @mr_name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Sub clean_old(job As String, mr_name As String)
        ' this sub unmarked the mr_data with x's
        Try
            Dim mr_list As New List(Of String)()

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", job)
            cmd4.Parameters.AddWithValue("@mr", mr_name)
            cmd4.CommandText = "select mr_name from Material_Request.mr where job = @job and BOM_type = 'old_BOM' and mr_name not like @mr"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    mr_list.Add(reader4(0).ToString)
                End While
            End If

            reader4.Close()

            For i = 0 To mr_list.Count - 1

                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@mr_name", mr_list.Item(i).ToString)

                Create_cmd.CommandText = "UPDATE Material_Request.mr_data SET latest_r = null where mr_name = @mr_name"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub ChangePanelDescriptionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangePanelDescriptionToolStripMenuItem.Click
        If String.Equals(Me.Text, "My Material Requests") = False Then

            Description_panels.Text = Me.Text
            Description_panels.ShowDialog()

        End If
    End Sub


    Sub send_confirmation_r(username As String, BOM_name As String, indexof_s As Integer, name_rev As String)

        Try

            Dim send_e As Boolean : send_e = False
            Dim Smtp_Server As New SmtpClient


            'host properties
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

            '---- get user email --------'
            Dim cmd41 As New MySqlCommand
            Dim email_user As String : email_user = "notfound"
            cmd41.Parameters.AddWithValue("@user", username)
            cmd41.CommandText = "SELECT email from users where username = @user"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then

                While reader41.Read
                    If reader41(0).ToString.Contains("@") Then
                        email_user = reader41(0)
                    End If
                End While

            End If

            reader41.Close()

            If String.Equals(email_user, "notfound") = False Then

                Dim mail_n As String : mail_n = "=====  THE FOLLOWING IS A APL CONFIRMATION MESSAGE FOR PROJECT : " & job_label.Text & " ======" & vbCrLf & vbCrLf

                mail_n = mail_n & vbCrLf & "Material Request Revision for Project " & job_label.Text & "  has been released" & vbCrLf & vbCrLf _
             & "Material Request Revised: " & Me.Text.Remove(indexof_s, Me.Text.Count - indexof_s) & name_rev & vbCrLf _
             & "Material Request Revised: " & real_mr & vbCrLf _
             & "Revised by: " & current_user

                '---------------------------------
                Try
                    Dim e_mail As New MailMessage()
                    e_mail = New MailMessage()
                    e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                    e_mail.To.Add(email_user)
                    e_mail.Subject = "MATERIAL REQUEST REVISION FOR PROJECT " & job_label.Text
                    e_mail.IsBodyHtml = False
                    e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                    Smtp_Server.Send(e_mail)

                Catch error_t As Exception
                    MsgBox(error_t.ToString)
                End Try

            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged

        If String.Equals(open_job, "") = False Then

            Try
                Dim string_q As String : string_q = "select distinct mr_name, release_date from Material_Request.mr where  released = 'Y' and job = '" & open_job & "' "

                ComboBox2.Items.Clear()


                Dim check_cmd As New MySqlCommand
                '  check_cmd.Parameters.AddWithValue("@job", job_label.Text)


                If String.Equals(ComboBox3.Text, "All BOMs") = True Then

                    string_q = string_q

                ElseIf String.Equals(ComboBox3.Text, "Master BOM") = True Then

                    string_q = string_q & " and BOM_type = 'MB' order by release_date"

                ElseIf String.Equals(ComboBox3.Text, "Panel") = True Then

                    string_q = string_q & " and BOM_type = 'Panel' order by release_date"

                ElseIf String.Equals(ComboBox3.Text, "Field") = True Then

                    string_q = string_q & " and BOM_type = 'Field' order by release_date"

                ElseIf String.Equals(ComboBox3.Text, "Assembly") = True Then
                    string_q = string_q & " and BOM_type = 'Assembly' order by release_date"

                ElseIf String.Equals(ComboBox3.Text, "Spare") = True Then

                    string_q = string_q & " and BOM_type = 'Spare' order by release_date"

                End If

                check_cmd.CommandText = string_q



                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        ComboBox2.Items.Add(reader(0))

                    End While
                End If

                reader.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub
End Class