Imports MySql.Data.MySqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Reflection

Public Class Form1

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Login_l.MouseEnter

        Login_l.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Login_l.MouseLeave

        Login_l.BackColor = Color.FromArgb(61, 60, 78)
    End Sub

    Private Sub Label8_MouseEnter(sender As Object, e As EventArgs) Handles Table_Label.MouseEnter

        Table_Label.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label8_MouseLeave(sender As Object, e As EventArgs) Handles Table_Label.MouseLeave

        Table_Label.BackColor = Color.FromArgb(61, 60, 78)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Timer1.Start()
        Search_field_box.Text = "Part_Number"
        enable_mess = True
        Call EnableDoubleBuffered(fc_grid)
        Call EnableDoubleBuffered(DataGridView1)
        Call EnableDoubleBuffered(Device_Table)

    End Sub


    Sub Login_pass()
        '============== DISPLAY THE DATABASE LOGIN SCREEN (NOT USER LOGIN SCREEN )==================
        Login.Visible = True
        Login.TopMost = True
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        '================ CLOSE CONNECTION WHEN CLOSING PROGRAM =====================
        If ok_press = True Then


            'If String.Equals(current_user, "") = False Then
            '    ' change status to no active if the user forgot to log out
            '    Try
            '        Dim status_cmd As New MySqlCommand
            '        status_cmd.Parameters.AddWithValue("@User", current_user)
            '        status_cmd.CommandText = "UPDATE management.status_user SET Status = 'Not Active', Log_in_time = now() where User = @User"
            '        status_cmd.Connection = Login.Connection
            '        status_cmd.ExecuteNonQuery()
            '    Catch ex As Exception
            '        MessageBox.Show(ex.ToString)
            '    End Try
            'End If

            Login.Connection.Close()
        End If
    End Sub


    Private Sub Form1_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        '============= WHEN FORM1 (MAIN FORM) APPEARS CALL PASSWORD LOGIN SCREEN =======================

        If Me.Visible = True Then
            '  Call Login_pass()
            Call Login.Access_database()
        End If
    End Sub


    Private Sub Label2_Click(sender As Object, e As EventArgs)

        '================== BOM MENU VISIBLE ICON=============================
        If ok_press = True Then
            BOM.Visible = True
        End If
    End Sub


    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Login_l.Click

        ' ============ USER LOGIN MENU logic ====================

        If Log_out = False And ok_press = True Then
            '  Login_rights.Visible = True
            Login_rights.ShowDialog()
        Else
            Login_l.Text = "User Log in"
            MessageBox.Show("You have logged out!")
            Log_out = False
            Me.Text = "APL v2.0"
            TabControl1.TabPages.Remove(TabPage10) 'remove cost tab

            '--close all forms except main one
            MyDash.Visible = False : MyDash.Close()
            Connections_win.Visible = False : Connections_win.Close()
            EnterPO.Visible = False : EnterPO.Close()
            User_form.Visible = False : User_form.Close()
            ADA_Setup.Visible = False : ADA_Setup.Close()
            Projects_m.Visible = False : Projects_m.Close()
            EnterPart.Visible = False : EnterPart.Close()
            UR_part.Visible = False : UR_part.Close()
            ImportCSV.Visible = False : ImportCSV.Close()
            add_KIT.Visible = False : add_KIT.Close()
            edit_KIT.Visible = False : edit_KIT.Close()
            add_DEVICE.Visible = False : add_DEVICE.Close()
            edit_DEVICE.Visible = False : edit_DEVICE.Close()
            Update_vendors.Visible = False : Update_vendors.Close()
            Reports.Visible = False : Reports.Close()
            myQuote.Visible = False : myQuote.Close()
            Quote_form.Visible = False : Quote_form.Close()
            PartCost_summary.Visible = False : PartCost_summary.Close()
            Setup_sheet.Visible = False : Setup_sheet.Close()
            Barc_Search.Visible = False : Barc_Search.Close()
            MailBox.Visible = False : MailBox.Close()
            Summary_devices.Visible = False : Summary_devices.Close()
            FMR_form.Visible = False : FMR_form.Close()
            My_Material_r.Visible = False : My_Material_r.Close()
            Project_dash.Visible = False : Project_dash.Close()
            Quick_inventory.Visible = False : Quick_inventory.Close()
            M_manager.Visible = False : M_manager.Close()
            My_Material_r.Visible = False : My_Material_r.Close()
            Repor_track.Visible = False : Repor_track.Close()
            Report_tracking_mn.Visible = False : Report_tracking_mn.Close()
            Inv_Material_r.Visible = False : Inv_Material_r.Close()
            Inventory_manage.Visible = False : Inventory_manage.Close()
            MPL_eng.Visible = False : MPL_eng.Close()
            MPL_mfg.Visible = False : MPL_mfg.Close()
            MPL_pc.Visible = False : MPL_pc.Close()
            Monitor_MR.Visible = False : Monitor_MR.Close()
            Assembly_report.Visible = False : Assembly_report.Close()
            Build_mfg.Visible = False : Build_mfg.Close()
            BOM_build.Visible = False : BOM_build.Close()
            Bom_packing.Visible = False : Bom_packing.Close()
            Add_return.Visible = False : Add_return.Close()
            BOM_init.Visible = False : BOM_init.Close()
            MR_init.Visible = False : MR_init.Close()
            INV_init.Visible = False : INV_init.Close()
            BOM_types.Visible = False : BOM_types.Close()
            Open_package.Visible = False : Open_package.Close()
            Part_selector.Visible = False : Part_selector.Close()
            Procurement_Overview.Visible = False : Procurement_Overview.Close()
            upload_file.Visible = False : upload_file.Close()


            'insert all tabs to dashboard
            MyDash.TabControl1.TabPages.Insert(0, MyDash.tab1_gm)
            MyDash.TabControl1.TabPages.Insert(1, MyDash.tab1_engm)
            MyDash.TabControl1.TabPages.Insert(2, MyDash.tab1_engineer)
            MyDash.TabControl1.TabPages.Insert(3, MyDash.tab1_pm)
            MyDash.TabControl1.TabPages.Insert(4, MyDash.tab1_proc)
            MyDash.TabControl1.TabPages.Insert(5, MyDash.tab1_inv)
            MyDash.TabControl1.TabPages.Insert(6, MyDash.tab1_manu)

            ' make all roles false
            Engineer = False
            Engineer_management = False
            Procurement = False
            Procurement_management = False
            Inventory_m = False
            Manufacturing = False
            General_management = False


            'change status to no active
            If ok_press = True Then
                Dim status_cmd As New MySqlCommand
                status_cmd.Parameters.AddWithValue("@User", current_user)
                status_cmd.CommandText = "UPDATE management.status_user SET Status = 'Not Active' where User = @User"
                status_cmd.Connection = Login.Connection
                status_cmd.ExecuteNonQuery()
                current_user = ""
            End If


        End If

    End Sub


    Private Sub Table_Label_Click(sender As Object, e As EventArgs) Handles Table_Label.Click

        '==================== CONTROL MENU FORM VISIBLE =================

        If Log_out = True Then

            'remove tabs according to user role


            If Engineer = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)



            ElseIf Engineer_management = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)

            ElseIf Procurement = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)

            ElseIf Procurement_management = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)

            ElseIf General_management = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)

            ElseIf Inventory_m = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_manu)

            ElseIf Manufacturing = True Then
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_gm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engineer)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_engm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_pm)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_proc)
                MyDash.TabControl1.TabPages.Remove(MyDash.tab1_inv)
            End If

            MyDash.Visible = True
            MyDash.WindowState = FormWindowState.Normal
        Else
                MessageBox.Show("Please, Log in")
        End If
    End Sub

    Private Sub Settings_Label_Click(sender As Object, e As EventArgs)

        ' ========== SHOW TRACKING FORM ===============
        If ok_press = True Then
            ' Tracking.Visible = True
            '  PO_Viewer.Visible = True
        End If

    End Sub

    Private Sub Login_l_Click(sender As Object, e As EventArgs)

        '=========== SHOW LOGIN SCREEN IF USER HAVENT LOG IN YET ========================

        If Log_out = False And ok_press = True Then
            Login_rights.Visible = True
        Else
            Login_l.Text = "User Log In"
            MessageBox.Show("You have logged out!")
            Log_out = False
            Engineer = False
            Engineer_management = False
            Procurement = False
            Procurement_management = False

            Me.Text = "ADA Database System"
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        '============== SYSTEM TIME =================
        current_Time.Text = "Current Time is: " & Now
    End Sub

    '  Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    ' ============ GOOGLE PART ==================
    ' Process.Start("www.google.com/#q=" & Component_n.Text)

    ' End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        '============== PRINTING DOC DIALOG ================

        PrintDialog1.Document = PrintDocument1
        Dim result As DialogResult
        result = PrintDialog1.ShowDialog()

        If (result = DialogResult.OK) Then
            PrintDocument1.Print()
        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        ' ============== PRINTING =====================
        Dim bm As New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)
        DataGridView1.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))
        e.Graphics.DrawImage(bm, 0, 0)

    End Sub



    ' Private Sub Attribute_Table_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Specs_Table.CellContentClick

    '========= DOUBLE CLICK PART NAME TO DISPLAY PART IMAGE =========

    'If (Specs_Table.CurrentCell.ColumnIndex = 0) Then

    'Dim type_p As String : type_p = ""
    'Dim path_im As String = "Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\"
    'Dim component As String : component = ""
    '  component = Specs_Table.CurrentCell.Value.ToString.Replace("/", "-")


    'Try

    'If Not Image.FromFile(path_im & component & ".JPG") Is Nothing Then
    '                PictureBox1.Image = Image.FromFile(path_im & component & ".JPG")
    '                Component_n.Text = Specs_Table.CurrentCell.Value.ToString
    '            End If

    '            'Check if specs pdf exist
    '            If File.Exists("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = False Then
    '                Label9.Text = "Not Available"
    '            Else
    '                Label9.Text = "Available"
    '            End If

    '        Catch ex As Exception

    '            Try
    '                'Display a generic image part type
    '                type_p = Specs_Table.Rows(Specs_Table.CurrentRow.Index).Cells(5).Value.ToString
    '                type_p = type_p.Replace("/", "-")
    '                PictureBox1.Image = Image.FromFile(path_im & type_p & ".JPG")
    '                Component_n.Text = Specs_Table.CurrentCell.Value.ToString
    '                'Check if specs pdf exist
    '                If File.Exists("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = False Then
    '                    Label9.Text = "Not Available"
    '                Else
    '             Label9.Text = "Available"
    '         End If

    '       Catch ex2 As Exception
    '           MessageBox.Show("Cannot Access Images Directory")
    '        End Try


    '      End Try
    '   End If

    ' End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

    '========= OPEN SPECS SHEET =================

    '  Dim component As String
    '  component = Component_n.Text.Replace("/", "-")

    '  Try
    ' System.Diagnostics.Process.Start("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
    ' Catch ex As Exception
    ' MessageBox.Show("Specs Sheet Not Found")
    ' End Try

    ' End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click

        '============ COPY MENU ================
        If DataGridView1.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(DataGridView1.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown

        '========= LEFT CLICK SELECT A CELL ===============
        If e.Button = MouseButtons.Right Then
            Try
                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)
                '  ToolStripTextBox1.Text = "1"
            Catch ex As Exception
            End Try

        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        '====================== ADD PARTS TO BOM ============================


        'If DataGridView1.CurrentCell.Value.ToString <> String.Empty Then

        '    If IsNumeric(ToolStripTextBox1.Text) Then  'If the amount is a number then

        '        BOM.DataGridView1.Rows.Add(New DataGridViewRow)
        '        BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(0).Value = DataGridView1.CurrentCell.Value.ToString


        '        'Get info about the part and put it in the BOM datagridview
        '        Try
        '            Dim cmd As New MySqlCommand
        '            cmd.CommandText = "SELECT Part_Description,  Primary_Vendor, Part_Type, Min_Order_Qty from parts_table where Part_Name  = """ & DataGridView1.CurrentCell.Value.ToString & """"

        '            cmd.Connection = Login.Connection
        '            Dim reader As MySqlDataReader
        '            reader = cmd.ExecuteReader

        '            If reader.HasRows Then

        '                While reader.Read
        '                    'WRITE PART INFO TO BOM DATAGRIDVIEW
        '                    If Not IsDBNull(reader(0)) Then
        '                        BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(1).Value = reader(0)  'Part description                          
        '                    End If

        '                    'If Not IsDBNull(reader(1)) Then
        '                    'BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(2).Value = reader(1)  'Primary Vendor   
        '                    'End If

        '                    BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(3).Value = reader(2)  'Part type

        '                    If Not IsDBNull(reader(3)) Then
        '                        If CType(ToolStripTextBox1.Text, Double) < reader(3) Then
        '                            MessageBox.Show("Entered Value is lower than the Minimum Order Quantity. Quantity has been adjusted accordingly ")
        '                            BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(5).Value = reader(3)
        '                        Else
        '                            BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(5).Value = ToolStripTextBox1.Text
        '                        End If
        '                    Else
        '                        BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(5).Value = ToolStripTextBox1.Text
        '                    End If

        '                End While
        '            End If

        '            reader.Close()

        '            BOM.DataGridView1.CurrentCell = BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(0) 'update selected cell
        '            BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(4).Value = Get_Latest_Cost(Login.Connection, DataGridView1.CurrentCell.Value.ToString, BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(2).Value)
        '            BOM.DataGridView1.RowTemplate.Height = 44
        '            MessageBox.Show(DataGridView1.CurrentCell.Value.ToString & " entered succesfully!")
        '            menu_BOM.Visible = False

        '        Catch ex As Exception
        '            MessageBox.Show(ex.ToString)
        '        End Try

        '    Else
        '        MessageBox.Show("Wrong Quantity Value")
        '    End If

        'Else
        '    MessageBox.Show("No part found!")
        'End If
    End Sub

    Function Get_Latest_Cost(myconnection As MySqlConnection, Part_Name As String, Vendor As String) As Decimal

        '============ GET THE LATEST COST OF THE PART SPECIFIED BY VENDOR ===========================

        Get_Latest_Cost = 0

        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.Parameters.AddWithValue("@part", Part_Name)
            cmd_v.Parameters.AddWithValue("@vendor", Vendor)
            cmd_v.CommandText = "SELECT Cost, Vendor_Name from vendors_table where Part_Name  = @part and Vendor_Name = @vendor  ORDER BY Purchase_Date DESC LIMIT 1"

            cmd_v.Connection = myconnection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader


            If readerv.HasRows Then
                While readerv.Read
                    Get_Latest_Cost = CType(readerv(0), Decimal)
                    '  BOM.DataGridView1.Rows(BOM.DataGridView1.Rows.Count - 2).Cells(2).Value = readerv(1)
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        ' ======== SEARCH PART NAME ===================
        'Searching Button find and highlight an update table with Part Name

        If ok_press = True Then

            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer
            Dim field As String : field = ""
            field = Search_field_box.Text

            If String.Equals(field, "") = True Or IsNothing(Search_field_box.SelectedItem) = True Then
                field = "Part_Number"
            End If

            If Match_b.Checked = True Then

                ' ================ Exact Match ====================
                If String.Equals(field, "Part_Name") = True Then
                    field = "Part_Number"
                End If


                If String.Equals(field, "Legacy_ADA_Number") = True Then
                    field = "ADA_Number"
                End If


                For Each row As DataGridViewRow In DataGridView1.Rows
                    If String.Compare(row.Cells.Item(field).Value.ToString, TextBox1.Text, True) = 0 Then
                        rowindex = row.Index
                        DataGridView1.CurrentCell = DataGridView1.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If
                Next
                If found_po = False Then
                    MsgBox("Part not found!")
                End If

            Else

                If String.Equals(field, "Part_Number") = True Then
                    field = "Part_Name"
                End If

                If String.Equals(field, "ADA_Number") = True Then
                    field = "Legacy_ADA_Number"
                End If
                '============== Partial match ============================
                Try
                    'Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where " & field & " LIKE  @search"
                    'Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                    'cmd.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")

                    '-------------  new  ----------------
                    Dim words As String() = TextBox1.Text.Replace("'", "").Split(New Char() {" "c})

                    Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where 1 = 1 "

                    If words.Length > 0 Then
                        For i = 0 To words.Length - 1
                            If String.IsNullOrEmpty(words(i)) = False Then
                                words(i).Replace("""", "")
                                words(i).Replace("'", "")
                                cmdstr = cmdstr & " and " & field & " LIKE '%" & words(i) & "%'"
                            End If
                        Next

                    End If


                    Dim cmd As New MySqlCommand(cmdstr, Login.Connection)



                    '-----------------------------------------------
                    Dim table_s As New DataTable
                    Dim adapter_s As New MySqlDataAdapter(cmd)

                    adapter_s.Fill(table_s)
                    DataGridView1.DataSource = table_s
                    'Setting Columns size for Parts Datagrid
                    For i = 0 To DataGridView1.ColumnCount - 1
                        With DataGridView1.Columns(i)
                            .Width = 200
                        End With
                    Next i

                    DataGridView1.Columns(0).Width = 480
                    DataGridView1.Columns(2).Width = 480

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If ok_press = True Then
            ' SHOW ALL PARTS IN THE (PARTS TABLE)
            Try
                Dim table As New DataTable
                Dim adapter As New MySqlDataAdapter("SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table order by Part_Name", Login.Connection)

                adapter.Fill(table)
                DataGridView1.DataSource = table

                'Setting Columns size for Parts Datagrid
                For i = 0 To DataGridView1.ColumnCount - 1
                    With DataGridView1.Columns(i)
                        .Width = 200
                    End With
                Next i

                DataGridView1.Columns(0).Width = 480
                DataGridView1.Columns(2).Width = 480

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded

        'UPDATE RECORDS
        Records_found.Text = "Parts Found: " & DataGridView1.Rows.Count.ToString

    End Sub

    Private Sub DataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        Records_found.Text = "Parts Found: " & DataGridView1.Rows.Count.ToString
    End Sub

    Private Sub PartTypebox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles PartTypebox1.SelectedValueChanged

        'UPDATE THE SELECTED PART TYPE AND STATUS ON THE DISPLAYED DATA

        If ok_press = True Then

            Try
                Dim selectedItem_type As String
                Dim selectedItem_status As String
                Dim part_query As String : part_query = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where 1=1"

                '--check if the combo boxes are null
                If Not PartTypebox1.SelectedItem Is Nothing Then
                    selectedItem_type = PartTypebox1.SelectedItem.ToString
                Else
                    selectedItem_type = ""
                End If

                If Not PartStatusbox1.SelectedItem Is Nothing Then
                    selectedItem_status = PartStatusbox1.SelectedItem.ToString
                Else
                    selectedItem_status = ""
                End If

                ' Concatenate query
                If String.Equals(selectedItem_type, "") = False Then
                    part_query = part_query & " and Part_Type = """ & selectedItem_type & """"
                End If

                If String.Equals(selectedItem_status, "") = False Then
                    part_query = part_query & " and Part_Status = """ & selectedItem_status & """"
                End If


                Try
                    Dim ta As New DataTable
                    Dim adapt As New MySqlDataAdapter(part_query, Login.Connection)
                    adapt.Fill(ta)
                    DataGridView1.DataSource = ta

                    For i = 0 To DataGridView1.ColumnCount - 1
                        With DataGridView1.Columns(i)
                            .Width = 200
                        End With
                    Next i

                    DataGridView1.Columns(0).Width = 480
                    DataGridView1.Columns(2).Width = 480
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If



    End Sub

    Private Sub PartStatusbox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles PartStatusbox1.SelectedValueChanged

        'UPDATE THE SELECTED PART TYPE AND STATUS ON THE DISPLAYED DATA (PARTS TAB)

        If ok_press = True Then

            Try
                Dim selectedItem_type As String
                Dim selectedItem_status As String
                Dim part_query As String : part_query = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where 1=1"


                If Not PartTypebox1.SelectedItem Is Nothing Then
                    selectedItem_type = PartTypebox1.SelectedItem.ToString
                Else
                    selectedItem_type = ""
                End If

                If Not PartStatusbox1.SelectedItem Is Nothing Then
                    selectedItem_status = PartStatusbox1.SelectedItem.ToString
                Else
                    selectedItem_status = ""
                End If

                '====== CONCATENATE QUERY ======================


                If String.Equals(selectedItem_type, "") = False Then
                    part_query = part_query & " and Part_Type = """ & selectedItem_type & """"

                End If

                If String.Equals(selectedItem_status, "") = False Then
                    part_query = part_query & " and Part_Status = """ & selectedItem_status & """"
                End If


                Try
                    Dim ta As New DataTable
                    Dim adapt As New MySqlDataAdapter(part_query, Login.Connection)
                    adapt.Fill(ta)
                    DataGridView1.DataSource = ta
                    For i = 0 To DataGridView1.ColumnCount - 1
                        With DataGridView1.Columns(i)
                            .Width = 200
                        End With
                    Next i

                    DataGridView1.Columns(0).Width = 480
                    DataGridView1.Columns(2).Width = 480
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If


    End Sub

    'Private Sub Attribute_Table_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)

    'UPDATE RECORDS
    '    Attributes_found.Text = "Records Found: " & Specs_Table.Rows.Count.ToString
    ' End Sub

    ' Private Sub Attribute_Table_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)

    'UPDATE RECORDS
    '  Attributes_found.Text = "Records Found: " & Specs_Table.Rows.Count.ToString

    ' End Sub


    ' Private Sub Button5_Click(sender As Object, e As EventArgs)

    'If ok_press = True Then
    'SHOW ALL PARTS IN THE (SPECIFICATIONS TABLE)
    ' Try
    'Dim table_a As New DataTable
    'Dim adapter_a As New MySqlDataAdapter("SELECT * from  parts_table order by Part_Name", Login.Connection)

    '     adapter_a.Fill(table_a)
    '     Specs_Table.DataSource = table_a
    'Catch ex As Exception
    '      MessageBox.Show(ex.ToString)
    'End Try
    'End If

    ' End Sub

    'Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click

    '    ' =========== KIT/ASSEMBLY VIEWER SEARCH =======================================
    '    If ok_press = True Then

    '        If Not device_show.Text.Equals("") Then

    '            Dim Atronix_n As String : Atronix_n = ""

    '            If IsInTable(device_show.Text, "Kits") = True Then
    '                ' If the KIT exist in the table then



    '                '--------------------------- Get AKN number --------------------
    '                Try
    '                    Dim cmd_k As New MySqlCommand
    '                    ' cmd_k.CommandText = "Select AKN_Number from Kits where KIT_Name = " & "'" & device_show.Text & "'"
    '                    cmd_k.CommandText = "Select AKN_Number from Kits where Legacy_ADA_Number = " & "'" & device_show.Text & "'"
    '                    cmd_k.Connection = Login.Connection

    '                    Dim reader_k As MySqlDataReader
    '                    reader_k = cmd_k.ExecuteReader


    '                    If reader_k.HasRows Then
    '                        While reader_k.Read
    '                            Atronix_n = reader_k(0)
    '                        End While
    '                    End If

    '                    reader_k.Close()


    '                Catch ex As Exception
    '                    MessageBox.Show(ex.ToString)
    '                End Try
    '                '-------------------------------------------------------------------------

    '                Dim table_akn As New DataTable
    '                Dim adapter_akn As New MySqlDataAdapter("SELECT p1.Part_Name, akn.Qty, p1.Manufacturer, p1.Part_Description, p1.Part_Status, p1.Part_Type,  p1.Primary_Vendor from parts_table as p1 INNER JOIN akn ON p1.Part_Name = akn.Part_Name where akn.AKN_Number = '" & Atronix_n & "' ", Login.Connection)

    '                adapter_akn.Fill(table_akn)
    '                part_assembly.DataSource = table_akn
    '                ' part_assembly.DataSource = table_akn

    '                'Setting Columns size 
    '                For i = 0 To part_assembly.ColumnCount - 1
    '                    With part_assembly.Columns(i)
    '                        .Width = 230
    '                    End With
    '                Next i

    '                part_assembly.Columns(1).Width = 60

    '                count_dk.Text = "# Components:   " & part_assembly.Rows.Count.ToString
    '                name_dk.Text = "Name:   " & device_show.Text

    '                '========== DISPLAY IMAGE OF THE KIT =============

    '                Dim path_vm As String = "Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\"
    '                Dim component_k As String : component_k = ""
    '                component_k = device_show.Text.ToString.Replace("/", "-")


    '                Try

    '                    If Not Image.FromFile(path_vm & component_k & ".JPG") Is Nothing Then
    '                        PictureBox3.Image = Image.FromFile(path_vm & component_k & ".JPG")
    '                    Else
    '                        PictureBox3.Image = PictureBox3.InitialImage
    '                    End If

    '                Catch ex As Exception
    '                    PictureBox3.Image = PictureBox3.InitialImage
    '                End Try


    '            ElseIf IsInTable(device_show.Text, "Devices") = True Then

    '                '=================================== If the device exist in the Devices table then ===================================

    '                '--------------------------- Get ADV number --------------------
    '                Try
    '                    Dim cmd_a As New MySqlCommand
    '                    '   cmd_a.CommandText = "Select ADV_Number from Devices where DEVICE_Name = " & "'" & device_show.Text & "'"
    '                    cmd_a.CommandText = "Select ADV_Number from Devices where Legacy_ADA_Number = " & "'" & device_show.Text & "'"
    '                    cmd_a.Connection = Login.Connection

    '                    Dim reader_k As MySqlDataReader
    '                    reader_k = cmd_a.ExecuteReader


    '                    If reader_k.HasRows Then
    '                        While reader_k.Read
    '                            Atronix_n = reader_k(0)
    '                        End While
    '                    End If

    '                    reader_k.Close()


    '                Catch ex As Exception
    '                    MessageBox.Show(ex.ToString)
    '                End Try
    '                '-------------------------------------------------------------------------

    '                Dim table_adv As New DataTable
    '                Dim adapter_adv As New MySqlDataAdapter("SELECT adv.Qty, p1.Part_Name, p1.Legacy_ADA_Number as ADA_Number, p1.Part_Description, p1.Manufacturer, p1.Part_Status, p1.Part_Type,  p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = '" & Atronix_n & "' ", Login.Connection)

    '                adapter_adv.Fill(table_adv)
    '                part_assembly.DataSource = table_adv
    '                part_assembly.DataSource = table_adv

    '                'Setting Columns size 
    '                For i = 0 To part_assembly.ColumnCount - 1
    '                    With part_assembly.Columns(i)
    '                        .Width = 230
    '                    End With
    '                Next i

    '                part_assembly.Columns(0).Width = 60

    '                count_dk.Text = "# Components:   " & part_assembly.Rows.Count.ToString
    '                name_dk.Text = "Name:   " & device_show.Text

    '                '========== DISPLAY IMAGE OF THE DEVICE =============

    '                Dim path_vm As String = "Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\"
    '                Dim component_k As String : component_k = ""
    '                component_k = device_show.Text.ToString.Replace("/", "-")


    '                Try

    '                    If Not Image.FromFile(path_vm & component_k & ".JPG") Is Nothing Then
    '                        PictureBox3.Image = Image.FromFile(path_vm & component_k & ".JPG")
    '                    Else
    '                        PictureBox3.Image = PictureBox3.InitialImage
    '                    End If

    '                Catch ex As Exception
    '                    PictureBox3.Image = PictureBox3.InitialImage
    '                End Try

    '            Else
    '                MessageBox.Show("Part Not Found!")
    '            End If
    '        End If
    '    End If

    'End Sub


    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    '============= SPECIFICATIONS TABLE SEARCH ====================

    'Searching Button find and highlight an update SPECS TABLE

    'If ok_press = True Then

    'Dim found_po As Boolean : found_po = False
    'Dim rowindex As Integer

    'If e_match.Checked = True Then

    ' ================ Exact Match ====================

    'For Each row As DataGridViewRow In Specs_Table.Rows
    'If String.Compare(row.Cells.Item("Part_Name").Value.ToString, Specs_search.Text) = 0 Then
    '         rowindex = row.Index
    '         Specs_Table.CurrentCell = Specs_Table.Rows(rowindex).Cells(0)
    '        found_po = True
    'Exit For
    'End If
    'Next
    'If found_po = False Then
    '       MsgBox(" Part not found!")
    'End If

    'Else

    '============== Partial match ============================
    'Try
    'Dim cmdstr As String : cmdstr = "SELECT * from parts_table where Part_Name LIKE  @search"
    'Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
    '    cmd.Parameters.AddWithValue("@search", "%" & Specs_search.Text & "%")

    'Dim adapter_spec As New MySqlDataAdapter(cmd)
    'Dim table_spec As New DataTable
    ' Dim adapter_spec As New MySqlDataAdapter("SELECT * from  parts_table where Part_Name LIKE ""%" & Specs_search.Text & "%""", Login.Connection)

    '     adapter_spec.Fill(table_spec)
    '     Specs_Table.DataSource = table_spec
    '     Catch ex As Exception
    '        MessageBox.Show(ex.ToString)
    '      End Try
    '   End If
    ' End If
    ' End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click

        '======= VENDOR TABLE SEARCH BUTTON ======================

        'Searching Button find and highlight an update Vendor table
        If ok_press = True Then

            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer
            Dim fieldv As String : fieldv = ""
            fieldv = vendor_f.Text

            If String.Compare(fieldv, "") = True Or IsNothing(vendor_f.SelectedItem) = True Then
                fieldv = "Part_Name"
            End If

            If v_match.Checked = True Then

                ' ============================= Exact Match =============================

                For Each row As DataGridViewRow In Vendors_grid.Rows
                    If String.Compare(row.Cells.Item(fieldv).Value.ToString, part_search_ven.Text) = 0 Then
                        rowindex = row.Index
                        Vendors_grid.CurrentCell = Vendors_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If
                Next
                If found_po = False Then
                    MsgBox(" Part not found!")
                End If

            Else

                '============================ Partial match ============================
                Try
                    Dim cmdstr As String : cmdstr = "SELECT * from  vendors_table where " & fieldv & " LIKE @search"
                    Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                    cmd.Parameters.AddWithValue("@search", "%" & part_search_ven.Text & "%")
                    Dim table_v As New DataTable
                    Dim adapter_v As New MySqlDataAdapter(cmd)

                    adapter_v.Fill(table_v)
                    Vendors_grid.DataSource = table_v
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If
        End If

    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        If ok_press = True Then
            'SHOW ALL PARTS IN THE (VENDOR TABLE)
            Try
                Dim table_v As New DataTable
                Dim adapter_v As New MySqlDataAdapter("SELECT Part_name, vendor_name, Cost, Purchase_date from  vendors_table order by Part_Name", Login.Connection)

                adapter_v.Fill(table_v)
                Vendors_grid.DataSource = table_v
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click

        '================ PRICES AND DATES SEARCH ========================

        If ok_press = True Then

            Try
                Dim selectedPrice As String
                Dim selectedDate As String
                Dim vendor_query As String : vendor_query = "SELECT * FROM vendors_table where 1=1"
                Dim price_value As Decimal
                Dim date_value As String


                ' Equal, greater ,less
                If Not Price_box.SelectedItem Is Nothing Then
                    selectedPrice = Price_box.SelectedItem.ToString
                    selectedPrice = Inequality(selectedPrice)
                Else
                    selectedPrice = ""
                End If

                If Not Date_box.SelectedItem Is Nothing Then
                    selectedDate = Date_box.SelectedItem.ToString
                    selectedDate = Inequality(selectedDate)
                Else
                    selectedDate = ""
                End If


                '------ values verification

                'check if cost is a number
                If IsNumeric(Price_T.Text) Then
                    price_value = Price_T.Text
                Else
                    price_value = 0
                End If

                'check if date is in the right format yyyy-mm-dd
                If Date_t.Text Like "####-##-##" Then
                    date_value = Date_t.Text
                Else
                    date_value = "1990-01-01"
                End If


                '==================== CONCATENATE QUERY ======================

                If String.Equals(selectedPrice, "") = False Then
                    vendor_query = vendor_query & " and Cost " & selectedPrice & " " & price_value
                End If

                If String.Equals(selectedDate, "") = False Then
                    vendor_query = vendor_query & " and Purchase_Date " & selectedDate & " '" & date_value & "'"
                End If


                Try
                    Dim tad As New DataTable
                    Dim adaptd As New MySqlDataAdapter(vendor_query, Login.Connection)
                    adaptd.Fill(tad)
                    Vendors_grid.DataSource = tad
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If



    End Sub

    Function Inequality(phrase As String) As String

        'Map string to equality symbols

        Inequality = ""

        Select Case phrase
            Case "Equal to"
                Inequality = "="
            Case "Greater than"
                Inequality = ">="
            Case "Less than"
                Inequality = "<="
        End Select
        Return Inequality

    End Function

    ' Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

    '============= ATRONIX NUMBER TABLE SEARCH ====================

    'Searching Button find and highlight an update ATRONIX NUMBER TABLE

    'If ok_press = True Then

    'Dim found_po As Boolean : found_po = False
    'Dim rowindex As Integer
    'Dim field As String : field = ""
    'field = number_box.Text

    'If String.Compare(field, "") = True Or IsNothing(number_box.SelectedItem) = True Then
    '    field = "Atronix_Number"
    'End If


    'If exact_a.Checked = True Then

    ' ================ Exact Match ====================

    'For Each row As DataGridViewRow In Cross_Table.Rows
    '                If String.Compare(row.Cells.Item(field).Value.ToString, Atronix_box.Text) = 0 Then
    '                    rowindex = row.Index
    '                    Cross_Table.CurrentCell = Cross_Table.Rows(rowindex).Cells(0)
    '                    found_po = True
    '                    Exit For
    '                End If
    '            Next
    '            If found_po = False Then
    '                MsgBox(" Atronix Number not found!")
    '            End If

    '        Else

    '            '============== Partial match ============================
    '            Try
    '                Dim cmdstr As String : cmdstr = "select * from ((Select ACN_Number As Atronix_Number, Part_Name, Legacy_ADA from acn)  UNION ALL (Select AKN_Number As Atronix_Number, Part_Name, Legacy_ADA from akn)  UNION ALL (Select ADV_Number As Atronix_Number, Part_Name, Legacy_ADA from adv) ) as t  where t." & field & " Like  @search"
    '                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
    '                cmd.Parameters.AddWithValue("@search", "%" & Atronix_box.Text & "%")

    '                Dim adapter_at As New MySqlDataAdapter(cmd)

    '                Dim table_at As New DataTable
    '                adapter_at.Fill(table_at)
    '                Cross_Table.DataSource = table_at
    ' Catch ex As Exception
    '      MessageBox.Show(ex.ToString)
    '   End Try
    ' End If
    '  End If

    '  End Sub

    'Private Sub Cross_Table_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
    '    'UPDATE RECORDS
    '    AN_found.Text = "Records Found: " & Cross_Table.Rows.Count.ToString
    'End Sub

    'Private Sub Cross_Table_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
    '    'UPDATE RECORDS
    '    AN_found.Text = "Records Found: " & Cross_Table.Rows.Count.ToString
    'End Sub

    'Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

    '    '============== Show alternate parts ======================

    '    If Not Alternate_find.Text.Equals("") Then

    '        Dim acn_ref As String : acn_ref = ""

    '        Try
    '            Dim cmd_al As New MySqlCommand
    '            cmd_al.CommandText = "Select ACN_Number from acn where Part_Name = " & "'" & Alternate_find.Text & "'"
    '            cmd_al.Connection = Login.Connection

    '            Dim reader_al As MySqlDataReader
    '            reader_al = cmd_al.ExecuteReader


    '            If reader_al.HasRows Then
    '                While reader_al.Read
    '                    acn_ref = reader_al(0)
    '                End While
    '            End If

    '            reader_al.Close()


    '        Catch ex As Exception
    '            MessageBox.Show(ex.ToString)
    '        End Try

    '        acn_label.Text = acn_ref

    '        '========================= POPULATE ALTERNATE TABLE ==============================

    '        Dim table_an As New DataTable
    '        Dim adapter_an As New MySqlDataAdapter("SELECT p1.Part_Name, acn.Legacy_ADA, p1.Manufacturer, p1.Part_Description, p1.Part_Status, p1.Part_Type,  p1.Primary_Vendor from parts_table as p1 INNER JOIN acn ON p1.Part_Name = acn.Part_Name where acn.ACN_Number = '" & acn_ref & "' and p1.Part_Name not like '" & Alternate_find.Text & "'", Login.Connection)

    '        adapter_an.Fill(table_an)
    '        Alter_table.DataSource = table_an
    '        Alter_table.DataSource = table_an

    '        'Setting Columns size 
    '        For i = 0 To Alter_table.ColumnCount - 1
    '            With Alter_table.Columns(i)
    '                .Width = 230
    '            End With
    '        Next i

    '    End If


    'End Sub

    'Private Sub Alter_table_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
    '    'UPDATE ALTERNATE PARTS
    '    Label10.Text = Alter_table.Rows.Count.ToString
    'End Sub

    'Private Sub Alter_table_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
    '    'UPDATE ALTERNATE PARTS
    '    Label10.Text = Alter_table.Rows.Count.ToString
    'End Sub

    'Private Sub Button7_Click(sender As Object, e As EventArgs)

    '    '========== Show all Atronix Numbers ===============

    '    If ok_press = True Then

    '        Dim table_an As New DataTable
    '        Dim adapter_an As New MySqlDataAdapter("(SELECT ACN_Number as Atronix_Number, Part_Name, Legacy_ADA from acn) UNION ALL (SELECT AKN_Number as Atronix_Number, Part_Name, Legacy_ADA from akn) UNION ALL SELECT ADV_Number as Atronix_Number, Part_Name, Legacy_ADA from adv", Login.Connection)

    '        adapter_an.Fill(table_an)     'Atronix Number fill
    '        Cross_Table.DataSource = table_an
    '        Cross_Table.DataSource = table_an   'Atronix Number table fill

    '        'Setting Columns size for Atronix Number Datagrid
    '        For i = 0 To Cross_Table.ColumnCount - 1
    '            With Cross_Table.Columns(i)
    '                .Width = 370
    '            End With
    '        Next i


    '    End If
    'End Sub

    '   Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

    '==================== View all KITS ==========================

    'If ok_press = True Then

    '    Dim table_kits As New DataTable
    '    Dim adapter_kits As New MySqlDataAdapter("SELECT  Legacy_ADA_Number as ADA_Number, Description from kits", Login.Connection)

    '    adapter_kits.Fill(table_kits)     'DataViewGrid1 fill
    '    Kit_Table.DataSource = table_kits
    '    Kit_Table.DataSource = table_kits   'Specifications table fill

    '    'Setting Columns size for Parts Datagrid
    '    For i = 0 To Kit_Table.ColumnCount - 1
    '        With Kit_Table.Columns(i)
    '            .Width = 850
    '        End With
    '    Next i

    'End If


    ' End Sub

    'Private Sub Kit_Table_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles Kit_Table.RowsAdded
    '    'Update Number of Kits found
    '    Kits_found.Text = "Kits Found: " & Kit_Table.Rows.Count.ToString
    'End Sub

    'Private Sub Kit_Table_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles Kit_Table.RowsRemoved
    '    'Update Number of Kits found
    '    Kits_found.Text = "Kits Found: " & Kit_Table.Rows.Count.ToString
    'End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        '==================== View all DEVICES ==========================

        If ok_press = True Then

            Cursor.Current = Cursors.WaitCursor
            Dim table_dev As New DataTable
            Dim adapter_dev As New MySqlDataAdapter("SELECT  Legacy_ADA_Number as ADA_Number, Description, Bulk_Cost, Labor_Cost, material_cost, total_cost from devices", Login.Connection) '(Material_Cost + Labor_Cost) as Cost

            adapter_dev.Fill(table_dev)     'DataViewGrid1 fill
            Device_Table.DataSource = table_dev
            Device_Table.DataSource = table_dev   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Device_Table.ColumnCount - 1
                With Device_Table.Columns(i)
                    .Width = 330
                End With
            Next i


            Device_Table.Columns(4).Visible = True
            Device_Table.Columns(5).Visible = True

            Call complete_costs()
            CheckBox2.Checked = True
            Cursor.Current = Cursors.Default

        End If

    End Sub

    Private Sub Device_Table_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles Device_Table.RowsAdded
        'Update Number of Devices found
        devices_found.Text = "Assemblies Found: " & Device_Table.Rows.Count.ToString
    End Sub

    Private Sub Device_Table_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles Device_Table.RowsRemoved
        'Update Number of Devices found
        devices_found.Text = "Assemblies Found: " & Device_Table.Rows.Count.ToString
    End Sub


    Function IsInTable(value As String, table As String) As Boolean

        ' ==============  Check if the device or kit exist ========================

        '  Dim xPart As String : xPart = "KIT_Name"
        IsInTable = False

        Try
            Dim cmd_i As New MySqlCommand
            '   cmd_i.CommandText = "Select * from " & table & " Where " & xPart & " = '" & value & "'"
            cmd_i.CommandText = "Select * from " & table & " Where Legacy_ADA_Number = '" & value & "'"
            cmd_i.Connection = Login.Connection

            Dim reader_i As MySqlDataReader
            reader_i = cmd_i.ExecuteReader


            If reader_i.HasRows Then
                IsInTable = True
            End If

            reader_i.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    'Private Sub Device_Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
    '    ' =========== Display name of device/kit when KIT/DEVICE Name column cell clicked =============

    '    If (Device_Grid.CurrentCell.ColumnIndex = 0) Then
    '        device_show.Text = Device_Grid.CurrentCell.Value.ToString
    '    End If
    'End Sub

    'Private Sub part_assembly_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    '    '=========== Click to display image on the Visualizer ========================
    '    If (part_assembly.CurrentCell.ColumnIndex = 1) Then

    '        Dim path_vm As String = "Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\"
    '        Dim component_v As String : component_v = ""
    '        component_v = part_assembly.CurrentCell.Value.ToString.Replace("/", "-")


    '        Try

    '            If Not Image.FromFile(path_vm & component_v & ".JPG") Is Nothing Then
    '                PictureBox4.Image = Image.FromFile(path_vm & component_v & ".JPG")
    '            Else
    '                PictureBox4.Image = PictureBox4.ErrorImage
    '            End If

    '        Catch ex As Exception
    '            PictureBox4.Image = PictureBox4.ErrorImage
    '        End Try
    '    End If
    'End Sub

    'Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
    '    '============= KITS TABLE SEARCH ====================

    '    'Searching Button find and highlight KIT TABLE

    '    If ok_press = True Then

    '        Dim found_po As Boolean : found_po = False
    '        Dim rowindex As Integer

    '        If kit_m.Checked = True Then

    '            ' ================ Exact Match ====================

    '            For Each row As DataGridViewRow In Kit_Table.Rows
    '                If String.Compare(row.Cells.Item("ADA_Number").Value.ToString, KIT_box.Text) = 0 Then
    '                    rowindex = row.Index
    '                    Kit_Table.CurrentCell = Kit_Table.Rows(rowindex).Cells(0)
    '                    found_po = True
    '                    Exit For
    '                End If
    '            Next
    '            If found_po = False Then
    '                MsgBox(" Kit not found!")
    '            End If

    '        Else



    '            '==========================  Partial match  ============================
    '            Try

    '                Dim cmdstr As String : cmdstr = "SELECT Legacy_ADA_Number As ADA_Number, Description from  kits  where Legacy_ADA_Number LIKE @search"
    '                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
    '                cmd.Parameters.AddWithValue("@search", "%" & KIT_box.Text & "%")

    '                Dim adapter_kits As New MySqlDataAdapter(cmd)
    '                Dim table_kits As New DataTable

    '                adapter_kits.Fill(table_kits)
    '                Kit_Table.DataSource = table_kits

    '                For i = 0 To Kit_Table.ColumnCount - 1
    '                    With Kit_Table.Columns(i)
    '                        .Width = 850
    '                    End With
    '                Next i

    '            Catch ex As Exception
    '                MessageBox.Show(ex.ToString)
    '            End Try
    '        End If
    '    End If
    'End Sub

    'Private Sub kit_draw_DoubleClick(sender As Object, e As EventArgs) Handles kit_draw.DoubleClick

    '    '============ Open the spec sheet of the selected KIT ================

    '    If (Kit_Table.CurrentCell.ColumnIndex = 0) Then

    '        Dim component As String : component = ""
    '        component = Kit_Table.CurrentCell.Value.ToString.Replace("/", "-")


    '        Try
    '            If File.Exists("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
    '                System.Diagnostics.Process.Start("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
    '            Else
    '                MessageBox.Show("Specification Sheet not available")
    '            End If

    '        Catch ex As Exception
    '        End Try

    '    End If
    'End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        '============= DEVICE TABLE SEARCH ====================

        'Searching Button find and highlight DEVICE TABLE

        If ok_press = True Then

            Dim found_po As Boolean : found_po = False
            Dim rowindex As Integer


            If d_match.Checked = True Then

                ' ================ Exact Match ====================

                For Each row As DataGridViewRow In Device_Table.Rows
                    If String.Compare(row.Cells.Item("ADA_Number").Value.ToString, device_box_t.Text) = 0 Then
                        rowindex = row.Index
                        Device_Table.CurrentCell = Device_Table.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If
                Next
                If found_po = False Then
                    MsgBox(" Assembly not found!")
                End If

            Else

                '============== Partial match ============================
                Try

                    Dim cmdstr As String : cmdstr = "SELECT  Legacy_ADA_Number as ADA_Number, Description, Bulk_Cost, Labor_Cost from devices where Legacy_ADA_Number LIKE @search"
                    Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                    cmd.Parameters.AddWithValue("@search", "%" & device_box_t.Text & "%")

                    Dim table_devices As New DataTable
                    '   Dim adapter_devices As New MySqlDataAdapter("SELECT DEVICE_Name, ADV_Number, Description, Material_Cost, Labor_Cost, (Material_Cost + Labor_Cost) as Cost from  devices where " & field & " LIKE ""%" & device_box_t.Text & "%""", Login.Connection)
                    Dim adapter_Devices As New MySqlDataAdapter(cmd)

                    adapter_Devices.Fill(table_devices)
                    Device_Table.DataSource = table_devices
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                '----------------------------------------------------------
            End If
        End If
    End Sub

    'Private Sub d_draw_DoubleClick(sender As Object, e As EventArgs) Handles d_draw.DoubleClick
    '    '============ Open the spec sheet of the selected DEVICE ================

    '    If (Device_Table.CurrentCell.ColumnIndex = 0) Then

    '        Dim component As String : component = ""
    '        component = Device_Table.CurrentCell.Value.ToString.Replace("/", "-")


    '        Try
    '            If File.Exists("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
    '                System.Diagnostics.Process.Start("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
    '            Else
    '                MessageBox.Show("Specification Sheet not available")
    '            End If

    '        Catch ex As Exception
    '        End Try

    '    End If
    'End Sub

    '  Private Sub Button9_Click(sender As Object, e As EventArgs)
    ' ========== SEARCH ADA NUMBER ==============


    'If ok_press = True Then
    '    Try
    '        Dim table_ada As New DataTable

    '        Dim adapter_ada As New MySqlDataAdapter("select * from ((Select Legacy_ADA_Number, Part_Name as Reference_Name from parts_table)  UNION ALL (Select Legacy_ADA_Number, KIT_Name as Reference_Name from kits)  UNION ALL (Select Legacy_ADA_Number, DEVICE_Name as Reference_Name from Devices) ) as t  where t.LEGACY_ADA_Number Like '%" & ADA_box.Text & "%'", Login.Connection)

    '        adapter_ada.Fill(table_ada)
    '        ADA_grid.DataSource = table_ada
    '        ADA_grid.Columns(0).Width = 720
    '        ADA_grid.Columns(1).Width = 850
    '    Catch ex As Exception
    '        MessageBox.Show(ex.ToString)
    '    End Try

    'End If
    '   End Sub

    '' Private Sub Button10_Click(sender As Object, e As EventArgs)
    'If ok_press = True Then
    '        ' SHOW ALL PARTS IN THE (SPECS TABLE)
    '        Try
    '            Dim table As New DataTable
    '            Dim adapter As New MySqlDataAdapter("SELECT * from  parts_table order by Part_Name", Login.Connection)

    '            adapter.Fill(table)
    '            Specs_Table.DataSource = table
    '        Catch ex As Exception
    '            MessageBox.Show(ex.ToString)
    '        End Try

    '    End If
    'End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        'Export data to CSV file

        Select Case grid_selector.Text
            Case "(Parts Table)"
                Call Convert_csv(DataGridView1)
                ' Case "(ADA Table)"
            '    Call Convert_csv(ADA_grid)
            Case "(Kits Table)"
              '  Call Convert_csv(Kit_Table)
            Case "(Assembly Table)"
                Call Convert_csv(Device_Table)
            Case "(Prices/Vendor)"
                If Log_out = True Then
                    Call Convert_csv(Vendors_grid)
                End If
        End Select

    End Sub

    Sub Convert_csv(temp_grid As DataGridView)

        'Creates a Text file from DataGridView

        Dim rows = From row As DataGridViewRow In temp_grid.Rows.Cast(Of DataGridViewRow)()
                   Where Not row.IsNewRow
                   Select Array.ConvertAll(row.Cells.Cast(Of DataGridViewCell).ToArray, Function(c) If(c.Value IsNot Nothing, c.Value.ToString, ""))

        Using sw As New IO.StreamWriter("csv.txt")
            For Each r In rows
                sw.WriteLine(String.Join(",", r))
            Next

        End Using
        Process.Start("csv.txt")

    End Sub


    'Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    '    Dim component As String : component = DataGridView1.CurrentCell.Value.ToString.Replace("/", "-")


    '    If File.Exists("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
    '        AxAcroPDF1.LoadFile("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
    '        MessageBox.Show("SPECS LOADED! SEE SPECIFICATIONS TAB")
    '    Else
    '        MessageBox.Show("NO SPECS FOUND")
    '    End If

    'End Sub


    'When a cell is clicked it will find the right PDF spec sheet and display it
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim component As String : component = DataGridView1.CurrentCell.Value.ToString.Replace("/", "-")


        If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
            System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
            '  AxAcroPDF1.LoadFile("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
            '  TabControl1.SelectedTab = TabPage2
            '  MessageBox.Show("SPECS LOADED! SEE SPECIFICATIONS TAB")
        Else
            ' AxAcroPDF1.LoadFile("NOTEXIST.pdf")
            MessageBox.Show("NO SPECS FOUND")
        End If
    End Sub



    Private Sub Device_Table_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Device_Table.CellDoubleClick
        '------- Display drawing and activate specs tab
        Dim component As String : component = Device_Table.CurrentCell.Value.ToString.Replace("/", "-")


        If File.Exists("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf") = True Then
            System.Diagnostics.Process.Start("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
            ' AxAcroPDF1.LoadFile("Z:\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & component & ".pdf")
            ' TabControl1.SelectedTab = TabPage2
            '  MessageBox.Show("SPECS LOADED! SEE SPECIFICATIONS TAB")
        Else
            'AxAcroPDF1.LoadFile("NOTEXIST.pdf")
            MessageBox.Show("NO SPECS FOUND")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Break down a part

        If ok_press = True Then

            Visualizer.mat_l.Text = "Material Cost: $"
            Dim component As String : component = ""
            component = Device_Table.CurrentCell.Value.ToString

            Call Break_part(component)

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        'break down feature code
        If ok_press = True Then

            Dim component As String : component = ""
            component = fc_grid.CurrentCell.Value.ToString

            Call FC_part(component)

        End If
    End Sub


    Sub FC_part(component As String)
        '-------- Show BOM for feature code --------------
        Dim my_assemblies = New List(Of String)()

        Try

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



            Dim cmd_po As New MySqlCommand
            cmd_po.Parameters.AddWithValue("@fq", component)
            cmd_po.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd_po.CommandText = "SELECT quote_table.feature_parts.qty, p1.Part_Name, p1.Manufacturer, p1.Part_Description, p1.Part_Status,  p1.Primary_Vendor from parts.parts_table as p1 INNER JOIN quote_table.feature_parts ON p1.Part_Name =  quote_table.feature_parts.part_name where  quote_table.feature_parts.Feature_code = @fq and quote_table.feature_parts.solution = @sol "

            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    Visualizer.part_assembly.Rows.Add()  'add a new row
                    Visualizer.part_assembly.Rows(i).Cells(1).Value = readerv(0).ToString
                    Visualizer.part_assembly.Rows(i).Cells(2).Value = readerv(1).ToString
                    Visualizer.part_assembly.Rows(i).Cells(3).Value = readerv(2).ToString
                    Visualizer.part_assembly.Rows(i).Cells(4).Value = readerv(3).ToString
                    Visualizer.part_assembly.Rows(i).Cells(5).Value = readerv(4).ToString
                    Visualizer.part_assembly.Rows(i).Cells(6).Value = readerv(5).ToString

                    i = i + 1
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        For j = 0 To Visualizer.part_assembly.Rows.Count - 1

            Try
                If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".jpg") Is Nothing Then
                    Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".jpg")
                    Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1
                End If

            Catch ex As Exception
                '----------------- default image -----------------
                Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\random_p.png")
                Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1
            End Try

            If my_assemblies.Contains(Visualizer.part_assembly.Rows(j).Cells(2).Value) = True Then
                Visualizer.part_assembly.Rows(j).Cells(7).Value = myQuote.Cost_of_Assem(Visualizer.part_assembly.Rows(j).Cells(2).Value)
            Else
                Visualizer.part_assembly.Rows(j).Cells(7).Value = Get_Latest_Cost(Login.Connection, Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString, Visualizer.part_assembly.Rows(j).Cells(6).Value.ToString)
            End If



        Next

        Visualizer.Text = component
        Visualizer.Visible = True
    End Sub


    Sub Break_part(component As String)

        If Not String.Equals(component, "") Then

            '  Visualizer.Visible = True
            Visualizer.part_assembly.Rows.Clear()
            Dim Atronix_n As String : Atronix_n = ""

            If IsInTable(component, "kits") = True Then
                ' If the KIT exist in the table then

                '--------------------------- Get AKN number --------------------
                Try
                    Dim cmd_k As New MySqlCommand
                    cmd_k.CommandText = "Select AKN_Number from kits where Legacy_ADA_Number = " & "'" & component & "'"
                    cmd_k.Connection = Login.Connection

                    Dim reader_k As MySqlDataReader
                    reader_k = cmd_k.ExecuteReader


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

                    Dim cmd_po As New MySqlCommand
                    cmd_po.CommandText = "SELECT akn.Qty, p1.Part_Name, p1.Manufacturer, p1.Part_Description, p1.Part_Status,  p1.Primary_Vendor from parts_table as p1 INNER JOIN akn ON p1.Part_Name = akn.Part_Name where akn.AKN_Number = '" & Atronix_n & "' "

                    cmd_po.Connection = Login.Connection
                    Dim readerv As MySqlDataReader
                    readerv = cmd_po.ExecuteReader

                    If readerv.HasRows Then
                        Dim i As Integer : i = 0
                        While readerv.Read

                            Visualizer.part_assembly.Rows.Add()  'add a new row
                            Visualizer.part_assembly.Rows(i).Cells(1).Value = readerv(0).ToString
                            Visualizer.part_assembly.Rows(i).Cells(2).Value = readerv(1).ToString
                            Visualizer.part_assembly.Rows(i).Cells(3).Value = readerv(2).ToString
                            Visualizer.part_assembly.Rows(i).Cells(4).Value = readerv(3).ToString
                            Visualizer.part_assembly.Rows(i).Cells(5).Value = readerv(4).ToString
                            Visualizer.part_assembly.Rows(i).Cells(6).Value = readerv(5).ToString

                            i = i + 1
                        End While
                    End If

                    readerv.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                For j = 0 To Visualizer.part_assembly.Rows.Count - 1

                    Try
                        If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".jpg") Is Nothing Then
                            Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".jpg")
                            Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1
                        End If

                    Catch ex As Exception
                        '----------------- default image -----------------
                        Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\random_p.png")
                        Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1
                    End Try

                    Visualizer.part_assembly.Rows(j).Cells(7).Value = Get_Latest_Cost(Login.Connection, Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString, Visualizer.part_assembly.Rows(j).Cells(6).Value.ToString)

                Next

                Visualizer.Text = component
                Visualizer.Visible = True


            ElseIf IsInTable(component, "devices") = True Then

                '=================================== If the device exist in the Devices table then ===================================

                '--------------------------- Get ADV number --------------------
                Try
                    Dim cmd_a As New MySqlCommand
                    cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = " & "'" & component & "'"
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
                    cmd_pd.CommandText = "SELECT adv.Qty, p1.Part_Name, p1.Part_Description, p1.Manufacturer, p1.Part_Status,  p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = '" & Atronix_n & "' "
                    cmd_pd.Connection = Login.Connection
                    Dim readerv As MySqlDataReader
                    readerv = cmd_pd.ExecuteReader

                    If readerv.HasRows Then
                        Dim i As Integer : i = 0
                        While readerv.Read

                            Visualizer.part_assembly.Rows.Add()  'add a new row
                            Visualizer.part_assembly.Rows(i).Cells(1).Value = readerv(0).ToString
                            Visualizer.part_assembly.Rows(i).Cells(2).Value = readerv(1).ToString
                            Visualizer.part_assembly.Rows(i).Cells(3).Value = readerv(2).ToString
                            Visualizer.part_assembly.Rows(i).Cells(4).Value = readerv(3).ToString
                            Visualizer.part_assembly.Rows(i).Cells(5).Value = readerv(4).ToString
                            Visualizer.part_assembly.Rows(i).Cells(6).Value = readerv(5).ToString

                            i = i + 1
                        End While
                    End If

                    readerv.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


                For j = 0 To Visualizer.part_assembly.Rows.Count - 1

                    Try
                        If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".JPG") Is Nothing Then
                            Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString.Replace(",", "").Replace("/", "") & ".JPG")
                            Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1

                        End If

                    Catch ex As Exception
                        '------------- default image -----------------
                        Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\random_p.png")
                        Visualizer.part_assembly.Rows(j).Cells(0).Value = bmp1
                    End Try

                    Visualizer.part_assembly.Rows(j).Cells(7).Value = Get_Latest_Cost(Login.Connection, Visualizer.part_assembly.Rows(j).Cells(2).Value.ToString, Visualizer.part_assembly.Rows(j).Cells(6).Value.ToString) * Visualizer.part_assembly.Rows(j).Cells(1).Value

                Next

                Visualizer.Text = component
                Visualizer.Visible = True
            Else
                MessageBox.Show("Part Not Found!")
            End If
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '------------- PO VIEWER SECTION ---------
        '================== Populate PO Info in datagridview =====================
        If ok_press = True And String.IsNullOrEmpty(po_box.Text) = False And String.IsNullOrEmpty(projects_box.Text) = False Then

            Dim po_name As String : po_name = ""
            Dim project_name As String : project_name = ""
            Dim company_name As String : company_name = ""


            DataGridView2.Rows.Clear() 'clear datagridview

            If Not po_box.SelectedItem Is Nothing Then
                po_name = po_box.SelectedItem.ToString()
            End If

            If Not projects_box.SelectedItem Is Nothing Then
                project_name = "job_" & projects_box.SelectedItem.ToString()
            End If


            Try

                Dim cmd_po As New MySqlCommand
                cmd_po.Parameters.AddWithValue("@PO", po_name)
                cmd_po.CommandText = "SELECT Part_Description, Qty_Ordered, sum(Qty_Received), Company from jobs." & project_name & " where PO = @PO group by Part_Description, Qty_Ordered, Company"

                cmd_po.Connection = Login.Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_po.ExecuteReader

                If readerv.HasRows Then
                    Dim i As Integer : i = 0
                    While readerv.Read

                        DataGridView2.Rows.Add()  'add a new row

                        DataGridView2.Rows(i).Cells(1).Value = readerv(0).ToString
                        DataGridView2.Rows(i).Cells(2).Value = readerv(1).ToString
                        DataGridView2.Rows(i).Cells(3).Value = readerv(2).ToString
                        'DataGridView2.Rows(i).Cells(4).Value = readerv(3).ToString

                        'company_name = readerv(4).ToString
                        company_name = readerv(3).ToString

                        i = i + 1
                    End While
                End If

                readerv.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try


            For j = 0 To DataGridView2.Rows.Count - 1

                Try

                    '------------------------- look for the part in parts table to get picture ----------------------------------
                    Dim cmd_img As New MySqlCommand

                    cmd_img.Parameters.AddWithValue("@Part_Name", DataGridView2.Rows(j).Cells(1).Value)
                    cmd_img.CommandText = "SELECT Part_Name, Part_Type FROM parts_table where Part_Name = @Part_Name"
                    cmd_img.Connection = Login.Connection
                    Dim reader_image As MySqlDataReader
                    reader_image = cmd_img.ExecuteReader


                    If reader_image.HasRows Then
                        '--------- real image-----------
                        While reader_image.Read
                            Try

                                If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & reader_image(0).ToString.Replace(",", "").Replace("/", "") & ".jpg") Is Nothing Then
                                    Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & reader_image(0).ToString.Replace(",", "").Replace("/", "") & ".jpg")
                                    DataGridView2.Rows(j).Cells(0).Value = bmp1
                                End If

                            Catch ex As Exception


                                '-------- else type image -----------
                                Try
                                    'Display a generic image part type

                                    Dim type_p As String : type_p = ""
                                    type_p = reader_image(1).ToString.Replace("/", "-") & ".jpg"
                                    Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\" & type_p)
                                    DataGridView2.Rows(j).Cells(0).Value = bmp1

                                Catch ex2 As Exception

                                End Try
                            End Try

                        End While
                    Else
                        '------------- default image -----------------
                        Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\Default_package.jpg")
                        DataGridView2.Rows(j).Cells(0).Value = bmp1
                    End If

                    reader_image.Close()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                'enter qty received
                DataGridView2.Rows(j).Cells(4).Value = DataGridView2.Rows(j).Cells(2).Value - DataGridView2.Rows(j).Cells(3).Value

            Next
            '--------------------------------------------------------------------------------------------------

            '---------- get total values ---------------
            Call Sum_parts(2)  'total ordered labels
            Call Sum_parts(3) 'total received labels

            po_label.Text = "PO Number:  " & po_name
            company_label.Text = "Company:     " & company_name

            '----------------- Display completeness image and percentage --------------------


            Dim percentage As Long : percentage = 0

            If String.Equals(po_name, "") = False Then
                percentage = (CType(received_label_v.Text, Long) * 100.0) / CType(ordered_label_v.Text, Long)

                If percentage > 100 Then
                    percentage = 100
                End If

                status_desc.Text = "Purchase Order Status:   " & CType(percentage, String) & "% Received"
            End If


            If percentage < 100 Then
                Try
                    Dim img As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\cross.png")
                    status_label.Image = img
                Catch ex As Exception
                    MessageBox.Show("Cannot access the server")
                End Try
                completeness.Text = "Incomplete"

            Else
                Try
                    Dim img As New System.Drawing.Bitmap("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\check.png")
                    status_label.Image = img
                Catch ex As Exception
                    MessageBox.Show("Cannot access the server")
                End Try
                completeness.Text = "Complete"
            End If

        End If

    End Sub

    Private Sub projects_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles projects_box.SelectedValueChanged
        '-------------------------- populate PO combobox ----------------------------------
        po_box.Items.Clear()

        Try
            Dim project_name As String : project_name = ""
            Dim cmd_po As New MySqlCommand

            If Not projects_box.SelectedItem Is Nothing Then
                project_name = "job_" & projects_box.SelectedItem.ToString()
            End If

            cmd_po.CommandText = "Select distinct PO from jobs." & project_name & " order by PO"
            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    po_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub DataGridView2_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView2.DoubleClick
        '--- when double click image in datagrid view, open tracking form
        Dim po_name As String : po_name = ""
        Dim project_name = ""

        If Not po_box.SelectedItem Is Nothing Then
            po_name = po_box.SelectedItem.ToString()
        End If

        If Not projects_box.SelectedItem Is Nothing Then
            project_name = "job_" & projects_box.SelectedItem.ToString()
        End If

        If (DataGridView2.CurrentCell.ColumnIndex = 0) Then
            Tracking.Visible = True

            Tracking.PO_Num.Text = po_label.Text
            Tracking.Part_name_t.Text = "Part Name:    " & DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(1).Value
            Tracking.Part_box.Image = DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(0).Value

            '------------------ populate date tracking grid ------------------
            Try

                Dim cmd_dt As New MySqlCommand
                cmd_dt.Parameters.AddWithValue("@Part_name", DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(1).Value)
                cmd_dt.Parameters.AddWithValue("@PO", po_name)
                cmd_dt.CommandText = "SELECT Date_Received as 'Received Date', QTY_Received as 'QTY Received' from jobs." & project_name & " where Part_Description = @Part_Name and PO = @PO"
                cmd_dt.Connection = Login.Connection
                Dim reader_date As MySqlDataReader
                reader_date = cmd_dt.ExecuteReader
                Dim z As Integer : z = 0

                If reader_date.HasRows Then
                    While reader_date.Read
                        If String.Equals(reader_date(1).ToString, "0") = False Then
                            Tracking.DataGridView1.Rows.Add()  'add a new row
                            Tracking.DataGridView1.Rows(z).Cells(0).Value = date_complete(reader_date(0).ToString)
                            Tracking.DataGridView1.Rows(z).Cells(1).Value = reader_date(1).ToString
                            z = z + 1
                        End If
                    End While
                End If

                reader_date.Close()

                Tracking.DataGridView1.Columns(0).Width = 690
                Tracking.DataGridView1.Columns(1).Width = 260


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Function date_complete(date_t As String) As String

        ' -------------------- Convert mm/dd/yyyy to Friday, April 20, 2018 --------------------------
        '-------- date_t should have mm/dd/yyyy format ------------

        Dim day As String : day = ""
        Dim name_day As String : name_day = ""
        Dim month As String : month = ""
        Dim month_name As String : month_name = ""
        Dim year As String : year = ""
        Dim f_date As String()
        Dim temp As String()

        temp = date_t.Split(New Char() {" "c})


        f_date = temp(0).ToString.Split(New Char() {"/"c})

        month = f_date(0)
        day = f_date(1)
        year = f_date(2)

        name_day = DateTime.Parse(year & "/" & month & "/" & day).DayOfWeek.ToString
        month_name = MonthName(month, True)

        date_complete = name_day & ", " & month_name & " " & day & ", " & year
        Return date_complete

    End Function

    Sub Sum_parts(n As Integer)

        '---- Sum quantities of a column in datagridview

        Dim sum_t As Double : sum_t = 0
        Try
            For i = 0 To DataGridView2.Rows.Count - 1
                sum_t = sum_t + CType(DataGridView2.Rows(i).Cells(n).Value(), Integer)
            Next

            If n = 2 Then
                ordered_label_v.Text = sum_t

            Else
                received_label_v.Text = sum_t
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        If ok_press = True Then
            Search_order.Visible = True
        End If
    End Sub

    '-------------- COMMAND HISTORY ------------------
    Sub Command_h(user As String, action As String, job As String)

        Dim role As String : role = "APL"

        Try
            '-- get role ----
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@user", user)
            cmd4.CommandText = "SELECT Role from users where username = @user"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                While reader4.Read
                    role = reader4(0).ToString
                End While

            End If

            reader4.Close()

            '----------------------

            '--- enter data to history -------
            Dim main_cmd As New MySqlCommand
            main_cmd.Parameters.AddWithValue("@username", user)
            main_cmd.Parameters.AddWithValue("@role", role)
            main_cmd.Parameters.AddWithValue("@action", action)
            main_cmd.Parameters.AddWithValue("@date", Now())
            main_cmd.Parameters.AddWithValue("@job", job)

            main_cmd.CommandText = "INSERT INTO management.history(username, role , action_m, date_m, job) VALUES (@username, @role, @action, now(), @job)"
            main_cmd.Connection = Login.Connection
            main_cmd.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        '-- autofind parts
        If Match_b.Checked = False Then
            Try

                Dim field As String : field = ""
                field = Search_field_box.Text


                If String.Equals(field, "Part_Number") = True Then
                    field = "Part_Name"
                End If

                If String.Equals(field, "ADA_Number") = True Then
                    field = "Legacy_ADA_Number"
                End If
                '============== Partial match ============================

                'Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where " & field & " LIKE  @search"
                'Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                'cmd.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")

                '----- new partial search ---
                Dim words As String() = TextBox1.Text.Replace("'", "").Split(New Char() {" "c})

                Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table where 1 = 1 "

                If words.Length > 0 Then
                    For i = 0 To words.Length - 1
                        If String.IsNullOrEmpty(words(i)) = False Then
                            words(i).Replace("""", "")
                            words(i).Replace("'", "")
                            cmdstr = cmdstr & " and " & field & " LIKE '%" & words(i) & "%'"
                        End If
                    Next
                End If


                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)

                '------------------

                Dim table_s As New DataTable
                Dim adapter_s As New MySqlDataAdapter(cmd)

                adapter_s.Fill(table_s)
                DataGridView1.DataSource = table_s
                'Setting Columns size for Parts Datagrid
                For i = 0 To DataGridView1.ColumnCount - 1
                    With DataGridView1.Columns(i)
                        .Width = 200
                    End With
                Next i

                DataGridView1.Columns(0).Width = 480
                DataGridView1.Columns(2).Width = 480

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub DataFilterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DataFilterToolStripMenuItem.Click
        '--- call filter data
        Call filter_data(DataGridView1.CurrentCell.ColumnIndex)

    End Sub

    Private Sub MoreInfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoreInfoToolStripMenuItem.Click

        Extra_info.Text = DataGridView1.CurrentCell.Value.ToString
        Extra_info.ShowDialog()
    End Sub

    Sub filter_data(header_index As Integer)

        '--- this function populates checklist box of filter parts form
        Dim my_collection = New List(Of String)()
        Filter_parts.CheckedListBox1.Items.Clear()
        Filter_parts.my_index = header_index


        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(header_index).Value Is DBNull.Value = False Then
                If my_collection.Contains(DataGridView1.Rows(i).Cells(header_index).Value) = False Then
                    my_collection.Add(DataGridView1.Rows(i).Cells(header_index).Value)
                End If
            End If
        Next

        my_collection.Sort()

        For i = 0 To my_collection.Count - 1
            Filter_parts.CheckedListBox1.Items.Add(my_collection.Item(i))
        Next

        Filter_parts.ShowDialog()
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

        If (Login.Connection.State <> ConnectionState.Open) Then
            Try
                Login.Connection = New MySqlConnection("datasource=10.21.2.8;port=3306;username=root;password=aapl_1369;database=parts;Allow User Variables=True")
                Login.Connection.Open()
                MessageBox.Show("APL is back again! :)")
            Catch ex As Exception
                MessageBox.Show("Cannot Connect to APL, :(")
            End Try
        Else
            MessageBox.Show("APL is online")
        End If
    End Sub

    Private Sub Label7_MouseEnter(sender As Object, e As EventArgs) Handles Label7.MouseEnter
        Label7.ForeColor = Color.Teal
    End Sub

    Private Sub Label7_MouseLeave(sender As Object, e As EventArgs) Handles Label7.MouseLeave
        Label7.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub fc_box_TextChanged(sender As Object, e As EventArgs) Handles fc_box.TextChanged
        '-- search feature code when text change
        fc_grid.Rows.Clear()

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Feature_code", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("type", GetType(String))
        dimen_table.Columns.Add("specific", GetType(String))
        dimen_table.Columns.Add("labor", GetType(String))
        dimen_table.Columns.Add("bulk", GetType(String))
        dimen_table.Columns.Add("total", GetType(String))


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@solution", sol_box.Text)
            cmd4.Parameters.AddWithValue("@search", "%" & fc_box.Text & "%")
            cmd4.CommandText = "Select Feature_code, description, type, specific_type, labor_cost, bulk_cost from quote_table.feature_codes where solution = @solution and Feature_code LIKE @search"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then

                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, If(IsDBNull(reader4(4)), "0", reader4(4).ToString), If(IsDBNull(reader4(5)), "0", reader4(5).ToString), 0)
                End While

            End If

            reader4.Close()

            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                fc_grid.Rows.Add(New String() {})
                fc_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                fc_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                fc_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                fc_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                fc_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                fc_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                fc_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString

            Next

            Call update_total_fc()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub sol_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles sol_box.SelectedValueChanged
        '-- change solution
        fc_grid.Rows.Clear()

        Dim cmd4 As New MySqlCommand
        cmd4.Parameters.AddWithValue("@solution", sol_box.Text)
        cmd4.CommandText = "Select Feature_code, description, type, specific_type, labor_cost, bulk_cost from quote_table.feature_codes where solution = @solution"
        cmd4.Connection = Login.Connection
        Dim reader4 As MySqlDataReader
        reader4 = cmd4.ExecuteReader

        If reader4.HasRows Then
            While reader4.Read
                fc_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, If(IsDBNull(reader4(4)), "0", reader4(4).ToString), If(IsDBNull(reader4(5)), "0", reader4(5).ToString), 0})
            End While
        End If

        reader4.Close()

        Call update_total_fc()

    End Sub


    Sub update_total_fc()
        '--- update feature code total tables
        Dim value_d As String : value_d = "0"


        For i = 0 To fc_grid.Rows.Count - 1

            If IsNumeric(fc_grid.Rows(i).Cells(4).Value) = True And IsNumeric(fc_grid.Rows(i).Cells(5).Value) = True And IsNumeric(fc_grid.Rows(i).Cells(6).Value) = True Then

                fc_grid.Rows(i).Cells(7).Value = CType(fc_grid.Rows(i).Cells(4).Value, Double) + CType(fc_grid.Rows(i).Cells(5).Value, Double) + CType(fc_grid.Rows(i).Cells(6).Value, Double)

                fc_grid.Rows(i).Cells(7).Value = "$" & Decimal.Round(fc_grid.Rows(i).Cells(7).Value, 2, MidpointRounding.AwayFromZero).ToString("N")

            End If
        Next
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub


    Function Feature_code_cost(feature_code As String, solution As String) As Double

        Feature_code_cost = 0

        Dim table_values As DataTable

        table_values = New DataTable
        table_values.Columns.Add("qty", GetType(String))
        table_values.Columns.Add("part_name", GetType(String))
        table_values.Columns.Add("vendor", GetType(String))
        table_values.Columns.Add("cost", GetType(String))
        table_values.Columns.Add("total", GetType(String))


        '------------ Calculate cost of feature_code BOM ---------------------
        Dim my_assemblies = New List(Of String)()

        Try

            '--------------  add to device ------------------------------
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
            '--------------------------------------------------------

            Dim cmd_po As New MySqlCommand
            cmd_po.Parameters.AddWithValue("@fq", feature_code)
            cmd_po.Parameters.AddWithValue("@sol", solution)
            cmd_po.CommandText = "SELECT quote_table.feature_parts.qty, p1.Part_Name, p1.Primary_Vendor from parts.parts_table as p1 INNER JOIN quote_table.feature_parts ON p1.Part_Name =  quote_table.feature_parts.part_name where  quote_table.feature_parts.Feature_code = @fq and quote_table.feature_parts.solution = @sol "

            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read
                    table_values.Rows.Add(readerv(0).ToString, readerv(1).ToString, readerv(2).ToString, 0, 0)
                End While
            End If

            readerv.Close()

            For j = 0 To table_values.Rows.Count - 1
                If my_assemblies.Contains(table_values.Rows(j).Item(1)) = True Then
                    table_values.Rows(j).Item(3) = myQuote.Cost_of_Assem(table_values.Rows(j).Item(1))
                Else
                    table_values.Rows(j).Item(3) = Get_Latest_Cost(Login.Connection, table_values.Rows(j).Item(1), table_values.Rows(j).Item(2))
                End If

                table_values.Rows(j).Item(4) = CType(table_values.Rows(j).Item(3), Double) * CType(table_values.Rows(j).Item(0), Double)

            Next

            For j = 0 To table_values.Rows.Count - 1
                Feature_code_cost = Feature_code_cost + CType(table_values.Rows(j).Item(4), Double)
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '---------------------------------------------------------------------


    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            For i = 0 To fc_grid.Rows.Count - 1
                fc_grid.Rows(i).Cells(6).Value = Feature_code_cost(fc_grid.Rows(i).Cells(0).Value, sol_box.Text)
            Next
        Else

            For i = 0 To fc_grid.Rows.Count - 1
                fc_grid.Rows(i).Cells(6).Value = 0
            Next

        End If

        update_total_fc()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        If CheckBox2.Checked = True Then

            Cursor.Current = Cursors.WaitCursor
            Device_Table.Columns(4).Visible = True
            Device_Table.Columns(5).Visible = True

            Call complete_costs()
            Cursor.Current = Cursors.Default

        Else
            If Device_Table.Columns.Count > 4 Then
                Device_Table.Columns(4).Visible = False
                Device_Table.Columns(5).Visible = False
            End If
        End If
    End Sub

    Sub complete_costs()

        '----------- fill material and total costs -------------

        For i = 0 To Device_Table.Rows.Count - 1

            Device_Table.Rows(i).Cells(4).Value = Convert.ToDecimal(CType(Cost_of_Assem(Device_Table.Rows(i).Cells(0).Value), Double).ToString("N"))
            Device_Table.Rows(i).Cells(5).Value = Convert.ToDecimal(CType((Device_Table.Rows(i).Cells(2).Value + Device_Table.Rows(i).Cells(3).Value + Device_Table.Rows(i).Cells(4).Value), Double).ToString("N"))
        Next



    End Sub

    Function Cost_of_Assem(assembly As String) As Double

        Cost_of_Assem = 0

        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("qty", GetType(Double))
        datatable.Columns.Add("primary_vendor", GetType(String))
        datatable.Columns.Add("total_cost", GetType(Double))


        Dim mat_c As Double : mat_c = 0


        Try
            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@ADA", assembly)
            cmd3.CommandText = "SELECT p1.Part_Name, adv.Qty, p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.Legacy_ADA  = @ADA"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    datatable.Rows.Add(reader3(0).ToString, reader3(1).ToString, reader3(2).ToString, 0)
                End While
            End If

            reader3.Close()

            For i = 0 To datatable.Rows.Count - 1
                datatable.Rows(i).Item(3) = Get_Latest_Cost(Login.Connection, datatable.Rows(i).Item(0), datatable.Rows(i).Item(2)) * datatable.Rows(i).Item(1)
                mat_c = mat_c + datatable.Rows(i).Item(3)
            Next



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Cost_of_Assem = mat_c

    End Function

    Sub get_email_list(job As String, ByVal lst As List(Of String))

        '------ return a list with all the emails associated to the passed job ------
        Try

            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@job", job)
            cmd3.CommandText = "SELECT p1.email from users as p1 INNER JOIN management.assignments ON p1.username = management.assignments.Employee_Name where management.assignments.Job_number = @job"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    If (IsDBNull(reader3(0))) = False Then
                        If String.IsNullOrEmpty(reader3(0).ToString) = False Then
                            lst.Add(reader3(0))
                        End If
                    End If
                End While
            End If

            reader3.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub


End Class
