Imports MySql.Data.MySqlClient
Imports System.Net.Mail

Public Class Assign_box

    Public Smtp_Server As New SmtpClient

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'convert mr in progress to release version

        If check_need_date() = False Then

            '--------- Panel presence checker --------------
            Try
                If does_it_has_panels(My_Material_r.Text) = False Then
                    Dim result2 As DialogResult = MessageBox.Show("There are no Panels associated with this BOM. Would you like to add panels ?", "No Panels Found", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

                    If (result2 = DialogResult.Yes) Then
                        Full_Panels.Text = My_Material_r.Text
                        Full_Panels.ShowDialog()
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
            '---------------------------------------------


            Dim result As DialogResult = MessageBox.Show("Are you sure you want to release this Material Request?  Releasing a Material Request will notify Procurement and Inventory to start allocating parts for the job. Please, make sure the Material Request is correct before releasing it", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

            If (result = DialogResult.Yes) Then

                If ComboBox1.SelectedItem Is Nothing = False And String.Equals(My_Material_r.Text, "My Material Requests") = False Then

                    Try
                        '----------- check if its already released -----------

                        ' My_Material_r.wait_la.Visible = True
                        '  Application.DoEvents()

                        Dim exist_sp As Boolean = False

                        'Dim cmd4 As New MySqlCommand
                        'cmd4.Parameters.AddWithValue("@job", ComboBox1.SelectedItem)
                        'cmd4.CommandText = "SELECT * from Material_Request.mr where job = @job and released = 'Y'"
                        'cmd4.Connection = Login.Connection
                        'Dim reader4 As MySqlDataReader
                        'reader4 = cmd4.ExecuteReader

                        'If reader4.HasRows Then
                        '    exist_c = True
                        'End If

                        'reader4.Close()

                        'check if specific mr has been released 

                        Dim cmd5 As New MySqlCommand
                        cmd5.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                        cmd5.CommandText = "SELECT released from Material_Request.mr where mr_name = @mr_name"
                        cmd5.Connection = Login.Connection
                        Dim reader5 As MySqlDataReader
                        reader5 = cmd5.ExecuteReader

                        If reader5.HasRows Then
                            While reader5.Read
                                If reader5(0) Is Nothing = False Then
                                    If String.Equals(reader5(0).ToString, "Y") Then
                                        exist_sp = True
                                    End If
                                End If
                            End While
                        End If

                        reader5.Close()

                        '---------------------------------

                        If exist_sp = False Then

                            '--- find out the number of BOM for the specified project ------

                            Dim id_bom As Double : id_bom = 0

                            Dim Create_cmd41 As New MySqlCommand
                            Create_cmd41.Parameters.Clear()
                            Create_cmd41.Parameters.AddWithValue("@job", ComboBox1.SelectedItem)
                            Create_cmd41.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"
                            Create_cmd41.Connection = Login.Connection
                            Dim reader As MySqlDataReader
                            reader = Create_cmd41.ExecuteReader

                            If reader.HasRows Then
                                While reader.Read
                                    id_bom = reader(0).ToString
                                End While
                            End If

                            reader.Close()

                            id_bom = id_bom + 1

                            '----------------------------------------------------------------
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.Clear()
                            Create_cmd.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
                            Create_cmd.Parameters.AddWithValue("@user", current_user)
                            Create_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem)
                            Create_cmd.Parameters.AddWithValue("@id_bom", id_bom)
                            Create_cmd.Parameters.AddWithValue("@ship_box", ship_box.Text)

                            Create_cmd.CommandText = "UPDATE Material_Request.mr SET released = 'Y', released_by = @user, release_date = now(), job = @job, id_bom = @id_bom, shipping_ad = @ship_box where mr_name = @mr_name"
                            Create_cmd.Connection = Login.Connection
                            Create_cmd.ExecuteNonQuery()

                            Dim Create_cmd2 As New MySqlCommand
                            Create_cmd2.Parameters.AddWithValue("@mr_name", My_Material_r.Text)

                            Create_cmd2.CommandText = "UPDATE Material_Request.mr_data SET released = 'Y' where mr_name = @mr_name"
                            Create_cmd2.Connection = Login.Connection
                            Create_cmd2.ExecuteNonQuery()


                            '--------- MPL released -------
                            'Dim Create_cmd3 As New MySqlCommand
                            'Create_cmd3.Parameters.Clear()
                            'Create_cmd3.Parameters.AddWithValue("@mpl_name", My_Material_r.Text)
                            'Create_cmd3.Parameters.AddWithValue("@user", current_user)
                            'Create_cmd3.Parameters.AddWithValue("@job", ComboBox1.SelectedItem)

                            'Create_cmd3.CommandText = "UPDATE Master_Packing_List.mpl SET released = 'Y', released_by = @user, release_date = now(), job = @job where mpl_name = @mpl_name"
                            'Create_cmd3.Connection = Login.Connection
                            'Create_cmd3.ExecuteNonQuery()

                            'Dim Create_cmd4 As New MySqlCommand
                            'Create_cmd4.Parameters.AddWithValue("@mpl_name", My_Material_r.Text)

                            'Create_cmd4.CommandText = "UPDATE Master_Packing_List.mpl_data SET released = 'Y' where mpl_name = @mpl_name"
                            'Create_cmd4.Connection = Login.Connection
                            'Create_cmd4.ExecuteNonQuery()
                            '--------------------------------

                            If enable_mess = True Then

                                'send APL notification to procurement and inventory and mfg

                                'write mail
                                Dim mail_n As String : mail_n = "Material Request for Project " & ComboBox1.SelectedItem & "  has been Released to Procurement" & vbCrLf & vbCrLf _
                         & "Released by: " & current_user & vbCrLf _
                         & "Released Date: " & Now & vbCrLf _
                         & "Project: " & ComboBox1.SelectedItem & vbCrLf _
                         & "BOM File Name: " & My_Material_r.Text & vbCrLf



                                Call Sent_mail.Sent_multiple_emails("Procurement", "Material Request and MPL has been Released for Project " & ComboBox1.SelectedItem, mail_n)
                                Call Sent_mail.Sent_multiple_emails("Procurement Management", "Material Request and MPL has been Released for Project " & ComboBox1.SelectedItem, mail_n)
                                Call Sent_mail.Sent_multiple_emails("Inventory", "Material Request and MPL has been Released for Project " & ComboBox1.SelectedItem, mail_n)
                                Call Sent_mail.Sent_multiple_emails("Manufacturing", "Material Request and MPL has been Released for Project " & ComboBox1.SelectedItem, mail_n)
                                Call Sent_mail.Sent_multiple_emails("General Management", "Material Request and MPL has been Released for Project " & ComboBox1.SelectedItem, mail_n)

                                'add email addresses
                                Dim emails_addr As New List(Of String)()

                                emails_addr.Add("TBullard@atronixengineering.com")
                                emails_addr.Add("dshipman@atronixengineering.com")

                                ''procurement
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

                                    For i = 0 To emails_addr.Count - 1
                                        e_mail.To.Add(emails_addr.Item(i))
                                    Next

                                    e_mail.Subject = "Material Request for Project " & ComboBox1.SelectedItem & "  has been Released"
                                    e_mail.IsBodyHtml = False
                                    e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

                                    Smtp_Server.Send(e_mail)

                                Catch error_t As Exception
                                    MsgBox(error_t.ToString)
                                End Try
                                ' Next

                            End If
                            '----------------------------------
                            Call Inventory_manage.General_inv_cal()   'recalculate inventory values
                            Call Form1.Command_h(current_user, "Material Request Released", ComboBox1.SelectedItem)

                            ' My_Material_r.wait_la.Visible = False
                            MessageBox.Show(My_Material_r.Text & " was released successfully!")
                            Me.Visible = False

                        Else
                            MessageBox.Show("Material Request has already been released!")
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.ToString)
                    End Try

                Else
                    MessageBox.Show("Please select a job and make sure a Material Request is open")

                End If
            End If

        Else
            MessageBox.Show("Please enter a need by date!")
        End If
    End Sub

    Private Sub Assign_box_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            Dim cmd_j As New MySqlCommand
            cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

            cmd_j.Connection = Login.Connection
            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    ComboBox1.Items.Add(readerj(0))

                End While
            End If

            readerj.Close()

            'host properties
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Assign_box_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = True Then
            ComboBox1.Items.Clear()

            Try
                Dim cmd_j As New MySqlCommand
                cmd_j.CommandText = "SELECT distinct Job_number from management.projects order by Job_number desc"

                cmd_j.Connection = Login.Connection
                Dim readerj As MySqlDataReader
                readerj = cmd_j.ExecuteReader

                If readerj.HasRows Then
                    While readerj.Read
                        ComboBox1.Items.Add(readerj(0))

                    End While
                End If

                readerj.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Enter_spec_parts()

        ''-------- enter project specific into inventory -----------
        For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
            If IsNumeric(My_Material_r.PR_grid.Rows(i).Cells(6).Value) = True And My_Material_r.PR_grid.Rows(i).Cells(0).Value Is Nothing = False And My_Material_r.PR_grid.Rows(i).Cells(8).Value Is Nothing = False Then

                If String.Equals(My_Material_r.PR_grid.Rows(i).Cells(8).Value, "Project_specific") = True And String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(i).Cells(9).Value) = True Then
                    '  Call Inventory_manage.add_specific(My_Material_r.PR_grid.Rows(i).Cells(0).Value, My_Material_r.PR_grid.Rows(i).Cells(1).Value, "Project_specific")
                End If
            End If
        Next
    End Sub

    Function check_need_date() As Boolean

        check_need_date = False

        For i = 0 To My_Material_r.PR_grid.Rows.Count - 1

            If String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(i).Cells(14).Value) = True And String.IsNullOrEmpty(My_Material_r.PR_grid.Rows(i).Cells(0).Value) = False Then
                check_need_date = True
            End If

        Next

    End Function

    Function does_it_has_panels(mr_name As String) As Boolean
        '--- check if this BOM is associated with panels

        does_it_has_panels = False

        Try
            Dim cmd41 As New MySqlCommand
            Dim count_m As Integer : count_m = 0

            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT * from Material_Request.full_panels where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    does_it_has_panels = True
                End While
            End If

            reader41.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function
End Class