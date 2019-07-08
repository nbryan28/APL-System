Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Revision_Control

    Public master_bom As String
    Public Smtp_Server As New SmtpClient


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim result As DialogResult = MessageBox.Show("Are you sure, you want to Push all this Revisions?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            Cursor.Current = Cursors.WaitCursor

            If check_duplicate_rev() = True Then

                Dim allow_m As Boolean : allow_m = False
                Dim mr_name_for_merge As String : mr_name_for_merge = ""

                For i = 0 To open_grid.Rows.Count - 1
                    If Convert.ToBoolean(open_grid.Rows(i).Cells(5).Value) = True Then
                        Call create_rev(open_grid.Rows(i).Cells(0).Value, open_grid.Rows(i).Cells(1).Value, If(String.IsNullOrEmpty(open_grid.Rows(i).Cells(4).Value) = True, "", open_grid.Rows(i).Cells(4).Value), master_bom, If(String.IsNullOrEmpty(open_grid.Rows(i).Cells(6).Value) = True, "", open_grid.Rows(i).Cells(6).Value), If(String.IsNullOrEmpty(open_grid.Rows(i).Cells(8).Value) = True, "", open_grid.Rows(i).Cells(8).Value), If(String.IsNullOrEmpty(open_grid.Rows(i).Cells(7).Value) = True, "", open_grid.Rows(i).Cells(7).Value))
                        allow_m = True
                        mr_name_for_merge = open_grid.Rows(i).Cells(0).Value
                    End If
                Next

                If allow_m = True Then

                    Call My_Material_r.Merge_and_release_MB(mr_name_for_merge)

                    Call BOM_types.Create_build_request(My_Material_r.open_job)    '-- create build_request revision
                    Call BOM_types.Create_MPL(My_Material_r.open_job)

                    '////////////////////// ---------------------  notify -----------------------
                    If enable_mess = True Then

                        Dim mail_n As String : mail_n = "Material Request Revision for Project " & My_Material_r.open_job & "  has been released" & vbCrLf & vbCrLf _
             & "Material Request Revised: " & mb_label.Text & vbCrLf _
             & "Revised by: " & current_user



                        Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request Revision has been Released for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request Revision has been Released for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request Revision has been Released for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("General Management", "Material Request Revision has been Released for Project " & My_Material_r.open_job, mail_n)


                        '--- sent email-------
                        'add email addresses
                        Dim emails_addr As New List(Of String)()

                        'Procurement
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

                            e_mail.Subject = "Material Request Revision for Project " & My_Material_r.open_job & "  has been Released"
                            e_mail.IsBodyHtml = False
                            e_mail.Body = mail_n & vbCrLf & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                            Smtp_Server.Send(e_mail)

                        Catch error_t As Exception
                            MsgBox(error_t.ToString)
                        End Try
                        '   Next


                        '-------- confirmation message -------------'
                        Call send_confirmation_r_n(current_user, mb_label.Text, My_Material_r.open_job)
                        '------------------------------------------

                    End If
                    '--------------------


                    MessageBox.Show("Revision Released Succesfully!")

                    Me.Visible = False

                    My_Material_r.job_label.Text = "Open Project:"
                    My_Material_r.PR_grid.Rows.Clear()
                    My_Material_r.Text = "My Material Requests"
                    My_Material_r.isitreleased = False
                    My_Material_r.rev_mode = False
                    Call Inventory_manage.General_inv_cal()   'recalculate inventory values

                    If My_Material_r.TabControl1.SelectedTab Is My_Material_r.TabPage2 Then
                        My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
                    End If
                End If
            End If


            Cursor.Current = Cursors.Default
        End If
    End Sub


    Sub send_confirmation_r_n(username As String, mbom As String, job As String)
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

                Dim mail_n As String : mail_n = "=====  THE FOLLOWING IS A APL CONFIRMATION MESSAGE FOR PROJECT : " & job & " ======" & vbCrLf & vbCrLf

                mail_n = mail_n & vbCrLf & "Material Request Revision for Project " & job & "  has been released" & vbCrLf & vbCrLf _
             & "Material Request Revised: " & mbom & vbCrLf _
             & "Revised by: " & current_user

                '---------------------------------
                Try
                    Dim e_mail As New MailMessage()
                    e_mail = New MailMessage()
                    e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")
                    e_mail.To.Add(email_user)
                    e_mail.Subject = "MATERIAL REQUEST REVISION FOR PROJECT " & job
                    e_mail.IsBodyHtml = False
                    e_mail.Body = mail_n & vbCrLf & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                    Smtp_Server.Send(e_mail)

                Catch error_t As Exception
                    MsgBox(error_t.ToString)
                End Try

            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub



    Private Sub Revision_Control_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '-- get target MBOM ---
        open_grid.Rows.Clear()
        Call init_revision_control(My_Material_r.Text)

        open_grid.Columns(6).Visible = False
        open_grid.Columns(7).Visible = False
        open_grid.Columns(8).Visible = False

        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"
    End Sub


    Sub init_revision_control(mr_name As String)

        '--get the name of the next master bom revision
        Dim MB As String : MB = ""
        Dim new_MB As String : new_MB = ""

        Try
            '-------- get MBOM ------------
            Dim cmdm As New MySqlCommand
            cmdm.Parameters.AddWithValue("@job", My_Material_r.open_job)
            cmdm.Parameters.AddWithValue("@mr_name", mr_name)
            cmdm.CommandText = "SELECT MBOM from Material_Request.mr where mr_name = @mr_name  and job = @job"
            cmdm.Connection = Login.Connection
            Dim readerm As MySqlDataReader
            readerm = cmdm.ExecuteReader

            If readerm.HasRows Then
                While readerm.Read
                    MB = readerm(0).ToString
                    master_bom = MB
                End While
            End If

            readerm.Close()
            '----------------------------

            If String.IsNullOrEmpty(MB) = False Then

                Dim name_rev As String = "_rev"
                Dim iz As Integer : iz = 0 'counter

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

                '-----------------------------------------------------------------------
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@job", My_Material_r.open_job)
                cmd4.Parameters.AddWithValue("@id_bom", id_bom_r)
                cmd4.CommandText = "Select shipping_ad ,Date_Created, created_by, id_bom, need_date, Panel_name, Panel_qty, Panel_desc, BOM_type, MBOM from Material_Request.mr where released = 'Y' and job = @job and id_bom = @id_bom order by release_date"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        iz = iz + 1
                    End While
                End If

                reader4.Close()

                name_rev = name_rev & iz  'last part of name of file ex: filename_revx
                Dim indexof_s = MB.IndexOf("_rev")

                If indexof_s < 0 Then
                    indexof_s = MB.Count
                End If

                new_MB = MB.Remove(indexof_s, MB.Count - indexof_s) & name_rev
                mb_label.Text = new_MB

                '--fill table ---

                Dim cmd48 As New MySqlCommand
                cmd48.Parameters.AddWithValue("@MB", MB)
                cmd48.CommandText = "SELECT mr_name, rev_name, created_date, created_by, new_panel, panel_name, desc_name, qty_name from Revisions.mr_rev where MB = @MB"
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
                        open_grid.Rows(i).Cells(4).Value = reader48(4).ToString
                        open_grid.Rows(i).Cells(6).Value = reader48(5).ToString
                        open_grid.Rows(i).Cells(7).Value = reader48(6).ToString
                        open_grid.Rows(i).Cells(8).Value = reader48(7).ToString
                        i = i + 1
                    End While
                End If

                reader48.Close()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub



    Function check_duplicate_rev() As Boolean
        'check if two revision with the same bom are selected
        check_duplicate_rev = True

        Dim boms As New Dictionary(Of String, String)


        For i = 0 To open_grid.Rows.Count - 1

            If Convert.ToBoolean(open_grid.Rows(i).Cells(5).Value) = True Then

                If boms.ContainsKey(open_grid.Rows(i).Cells(0).Value) = False Then
                    boms.Add(open_grid.Rows(i).Cells(0).Value, open_grid.Rows(i).Cells(0).Value)

                Else
                    MessageBox.Show("Please pick one revision for " & open_grid.Rows(i).Cells(0).Value)

                    check_duplicate_rev = False
                    Exit For
                End If
            End If

        Next

    End Function

    Sub create_rev(mr_name As String, rev_name As String, mode_t As String, MB As String, panel As String, qty_p As String, desc As String)

        Select Case mode_t

            Case "Add"
                Call add_rev(mr_name, rev_name, MB, panel, qty_p, desc)
            Case "Zero Out"
                Call zero_rev(mr_name, rev_name, MB, panel, desc)
            Case ""
                Call plain_rev(mr_name, rev_name, MB)
            Case Else
                Call change_rev(mr_name, rev_name, MB, panel, qty_p, desc)

        End Select
    End Sub

    '----------/////// Add a panel, create a revision -------
    Sub add_rev(mr_name As String, rev_name As String, MB As String, panel As String, qty_p As String, desc As String)

        Panel_grid.Rows.Clear()

        Try

            '---------------------------- add data to panel_grid
            Dim check_cmd1 As New MySqlCommand
            check_cmd1.Parameters.AddWithValue("@mr_name", mr_name)
            check_cmd1.Parameters.AddWithValue("@rev_name", rev_name)
            check_cmd1.CommandText = "select Part_No, description_t, Manufacturer, Vendor, Price, new_qty, mfg_type, part_status, part_type, Notes, need_date  from Revisions.mr_rev_data where mr_name = @mr_name and rev_name = @rev_name"

            check_cmd1.Connection = Login.Connection
            check_cmd1.ExecuteNonQuery()

            Dim reader1 As MySqlDataReader
            reader1 = check_cmd1.ExecuteReader

            If reader1.HasRows Then

                While reader1.Read
                    Panel_grid.Rows.Add(New String() {reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString, reader1(5).ToString, 0, reader1(6).ToString, reader1(7).ToString, reader1(8).ToString, reader1(9).ToString, reader1(10).ToString})
                End While
            End If

            reader1.Close()
            '----------------------------

            Call BOM_types.Save_bom(Panel_grid, mr_name, My_Material_r.open_job, "Panel", False, panel, desc, 1, MB)

            Dim n_bom As Double : n_bom = 0
            Dim check_cmd8 As New MySqlCommand
            check_cmd8.Parameters.AddWithValue("@job", My_Material_r.open_job)
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
            Create_cmd.Parameters.AddWithValue("@mr_name", mr_name)
            Create_cmd.Parameters.AddWithValue("@released", "Y")
            Create_cmd.Parameters.AddWithValue("@released_by", current_user)
            Create_cmd.Parameters.AddWithValue("@id_bom", n_bom)
            Create_cmd.Parameters.AddWithValue("@Panel_qty", If(IsNumeric(qty_p) = True, qty_p, 1))

            Create_cmd.CommandText = "UPDATE Material_Request.mr SET  released = @released,  released_by = @released_by, release_date = now(), id_bom = @id_bom, Panel_qty = @Panel_qty where mr_name = @mr_name"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Panel_grid.Rows.Clear()

    End Sub


    '----------/////////// zero panel, create a revision -------
    Sub zero_rev(mr_name As String, rev_name As String, MB As String, panel As String, desc As String)


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
                dimen_table.Rows.Add(readera(0).ToString, readera(1).ToString, readera(2).ToString, readera(3).ToString, readera(4).ToString, readera(5).ToString, readera(6).ToString, readera(7).ToString, If(readera(8) Is DBNull.Value, "0", readera(8)), If(readera(9) Is DBNull.Value, "0", readera(9)), If(readera(10) Is DBNull.Value, "", readera(10)), readera(11).ToString, readera(12).ToString, readera(13).ToString, readera(14).ToString)
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
        cmd4.Parameters.AddWithValue("@job", My_Material_r.open_job)
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
        main_cmd.Parameters.AddWithValue("@job", My_Material_r.open_job)
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
            Create_cmd6.Parameters.AddWithValue("@qty_fullfilled", If(dimen_table.Rows(i).Item(8) Is Nothing, "0", dimen_table.Rows(i).Item(8)))
            Create_cmd6.Parameters.AddWithValue("@qty_needed", If(dimen_table.Rows(i).Item(9) Is Nothing, "0", dimen_table.Rows(i).Item(9)))
            Create_cmd6.Parameters.AddWithValue("@Assembly_name", If(dimen_table.Rows(i).Item(10) Is Nothing, "", dimen_table.Rows(i).Item(10)))
            Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table.Rows(i).Item(11) Is Nothing, "", dimen_table.Rows(i).Item(11)))
            Create_cmd6.Parameters.AddWithValue("@Part_Type", If(dimen_table.Rows(i).Item(12) Is Nothing, "", dimen_table.Rows(i).Item(12)))
            Create_cmd6.Parameters.AddWithValue("@Notes", If(dimen_table.Rows(i).Item(13) Is Nothing, "", dimen_table.Rows(i).Item(13)))
            Create_cmd6.Parameters.AddWithValue("@Need_by_date", If(dimen_table.Rows(i).Item(14) Is Nothing, "", dimen_table.Rows(i).Item(14)))


            Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, qty_fullfilled, qty_needed, Assembly_name, Part_status, Part_Type, Notes, need_by_date ) VALUES (@mr_name, @Part_No, @Description, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fullfilled, @qty_needed, @Assembly_name, @Part_status, @Part_Type, @Notes,  @Need_by_date  )"
            Create_cmd6.Connection = Login.Connection
            Create_cmd6.ExecuteNonQuery()

        Next


    End Sub


    '----------////////// create main revision -------
    Sub plain_rev(mr_name As String, rev_name As String, MB As String)

        Try
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


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", My_Material_r.open_job)
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
            Dim indexof_s = mr_name.IndexOf("_rev")

            If indexof_s < 0 Then
                indexof_s = mr_name.Count
            End If

            '---------- put all data including qty fulfilled -------

            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Part_No", GetType(String))
            dimen_table.Columns.Add("Description", GetType(String))
            dimen_table.Columns.Add("ADA_Number", GetType(String))
            dimen_table.Columns.Add("Manufacturer", GetType(String))
            dimen_table.Columns.Add("Vendor", GetType(String))
            dimen_table.Columns.Add("Price", GetType(String))
            dimen_table.Columns.Add("Qty", GetType(String))
            dimen_table.Columns.Add("subtotal", GetType(String))
            dimen_table.Columns.Add("mfg_type", GetType(String))
            dimen_table.Columns.Add("Assembly_name", GetType(String))
            dimen_table.Columns.Add("Notes", GetType(String))
            dimen_table.Columns.Add("Part_status", GetType(String))
            dimen_table.Columns.Add("Part_type", GetType(String))
            dimen_table.Columns.Add("full_panel", GetType(String))
            dimen_table.Columns.Add("need_by_date", GetType(String))
            dimen_table.Columns.Add("qty_fulfilled", GetType(String))

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.Parameters.AddWithValue("@rev_name", rev_name)
            cmd41.CommandText = "Select Part_No, description_t, ADA_Number, Manufacturer, Vendor, Price, new_qty, '0', mfg_type, '', Notes, part_status, part_type, isitfull, need_date from Revisions.mr_rev_data where mr_name = @mr_name and rev_name = @rev_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    dimen_table.Rows.Add(reader41(0).ToString, If(reader41(1) Is DBNull.Value, 0, reader41(1)), If(reader41(2) Is DBNull.Value, 0, reader41(2)), If(reader41(3) Is DBNull.Value, 0, reader41(3)), If(reader41(4) Is DBNull.Value, 0, reader41(4)), If(reader41(5) Is DBNull.Value, 0, reader41(5)), If(reader41(6) Is DBNull.Value, 0, reader41(6)), If(reader41(7) Is DBNull.Value, 0, reader41(7)), If(reader41(8) Is DBNull.Value, 0, reader41(8)), If(reader41(9) Is DBNull.Value, 0, reader41(9)), If(reader41(10) Is DBNull.Value, 0, reader41(10)), If(reader41(11) Is DBNull.Value, 0, reader41(11)), If(reader41(12) Is DBNull.Value, 0, reader41(12)), If(reader41(13) Is DBNull.Value, 0, reader41(13)), If(reader41(14) Is DBNull.Value, 0, reader41(14)), 0)
                End While
            End If

            reader41.Close()


            '--- qty fulfilled ----
            For i = 0 To dimen_table.Rows.Count - 1

                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", mr_name)
                cmd419.Parameters.AddWithValue("@Part_No", dimen_table.Rows(i).Item(0))
                cmd419.CommandText = "Select qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name and Part_No = @Part_No"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read
                        dimen_table.Rows(i).Item(15) = reader419(0)
                    End While
                End If

                reader419.Close()
            Next

            '/////// start inserting revision to table /////////////

            Dim main_cmd As New MySqlCommand
            main_cmd.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
            main_cmd.Parameters.AddWithValue("@released_by", current_user)
            main_cmd.Parameters.AddWithValue("@job", My_Material_r.open_job)
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



            For i = 0 To dimen_table.Rows.Count - 1

                Dim Create_cmd6 As New MySqlCommand
                Create_cmd6.Parameters.Clear()
                Create_cmd6.Parameters.AddWithValue("@mr_name", mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table.Rows(i).Item(0) Is Nothing, "", dimen_table.Rows(i).Item(0)))
                Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table.Rows(i).Item(1) Is Nothing, "", dimen_table.Rows(i).Item(1)))
                Create_cmd6.Parameters.AddWithValue("@ADA_Number", If(dimen_table.Rows(i).Item(2) Is Nothing, "", dimen_table.Rows(i).Item(2)))
                Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(dimen_table.Rows(i).Item(3) Is Nothing, "", dimen_table.Rows(i).Item(3)))
                Create_cmd6.Parameters.AddWithValue("@Vendor", If(dimen_table.Rows(i).Item(4) Is Nothing, "", dimen_table.Rows(i).Item(4)))
                Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table.Rows(i).Item(5) Is Nothing, "", dimen_table.Rows(i).Item(5).ToString.Replace("$", "")))
                Create_cmd6.Parameters.AddWithValue("@Qty", If(dimen_table.Rows(i).Item(6) Is Nothing, "", dimen_table.Rows(i).Item(6)))
                Create_cmd6.Parameters.AddWithValue("@subtotal", If(IsNumeric(dimen_table.Rows(i).Item(5)) = True, CType(dimen_table.Rows(i).Item(5), Double), 0) * If(IsNumeric(dimen_table.Rows(i).Item(6)) = True, CType(dimen_table.Rows(i).Item(6), Double), 0))
                Create_cmd6.Parameters.AddWithValue("@mfg_type", If(dimen_table.Rows(i).Item(8) Is Nothing, "", dimen_table.Rows(i).Item(8).ToString))
                Create_cmd6.Parameters.AddWithValue("@Assembly_name", If(dimen_table.Rows(i).Item(9) Is Nothing, "", dimen_table.Rows(i).Item(9).ToString))
                Create_cmd6.Parameters.AddWithValue("@Notes", If(dimen_table.Rows(i).Item(10) Is Nothing, "", dimen_table.Rows(i).Item(10).ToString))
                Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table.Rows(i).Item(11) Is Nothing, "", dimen_table.Rows(i).Item(11).ToString))
                Create_cmd6.Parameters.AddWithValue("@Part_type", If(dimen_table.Rows(i).Item(12) Is Nothing, "", dimen_table.Rows(i).Item(12).ToString))
                Create_cmd6.Parameters.AddWithValue("@full_panel", If(dimen_table.Rows(i).Item(13) Is Nothing, "", dimen_table.Rows(i).Item(13).ToString))
                Create_cmd6.Parameters.AddWithValue("@need_by_date", If(dimen_table.Rows(i).Item(14) Is Nothing, "", dimen_table.Rows(i).Item(14).ToString))


                If String.Equals(BOM_type, "old_BOM") = True Then
                    Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released, latest_r ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y', 'x' )"
                    Call My_Material_r.clean_old(My_Material_r.open_job, mr_name.Remove(indexof_s, mr_name.Count - indexof_s) & name_rev)
                Else
                    Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, qty_fullfilled, Assembly_name, Notes, Part_status, Part_Type, full_panel, need_by_date, released ) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @qty_fulfilled, @Assembly_name, @Notes, @Part_status, @Part_Type, @full_panel, @need_by_date, 'Y' )"
                End If


                Create_cmd6.Connection = Login.Connection
                Create_cmd6.ExecuteNonQuery()
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub


    '----------////////// change qty of panels -------
    Sub change_rev(mr_name As String, rev_name As String, MB As String, panel As String, qty_p As String, desc As String)


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
        Dim Panel_qty As String : Panel_qty = qty_p
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
        cmd4.Parameters.AddWithValue("@job", My_Material_r.open_job)
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
        main_cmd.Parameters.AddWithValue("@job", My_Material_r.open_job)
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


    End Sub


    Private Sub open_grid_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles open_grid.CellClick
        '--- open panel revision -----

        Panel_grid.Rows.Clear()

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
                    Panel_grid.Rows.Add(New String() {reader1(0).ToString, reader1(1).ToString, reader1(2).ToString, reader1(3).ToString, reader1(4).ToString, reader1(5).ToString, 0, reader1(6).ToString, reader1(7).ToString, reader1(8).ToString, reader1(9).ToString, reader1(10).ToString})
                End While
            End If

            reader1.Close()



        End If
    End Sub
End Class