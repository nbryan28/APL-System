Imports MySql.Data.MySqlClient
Imports System.Net.Mail

Public Class Shipping_Edit

    Public released_t As String
    Public job As String
    Public Smtp_Server As New SmtpClient
    Public project_ship As String

    Private Sub Shipping_Edit_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        shipbox.Text = "Project Address"
        Ship_b.ReadOnly = True

        Try
            'host properties
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.gmail.com"


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
            cmd4.CommandText = "SELECT shipping_ad, released, job from Material_Request.mr where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    Ship_b.Text = reader4(0).ToString
                    project_ship = reader4(0).ToString
                    released_t = reader4(1).ToString
                    job = reader4(2).ToString
                End While

            End If

            reader4.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
            Dim Create_cmd As New MySqlCommand
            Create_cmd.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
            Create_cmd.Parameters.AddWithValue("@ship", Ship_b.Text)
            Create_cmd.Parameters.AddWithValue("@notes", TextBox1.Text)

            Create_cmd.CommandText = "UPDATE Material_Request.mr SET shipping_ad = @ship, ship_notes = @notes where mr_name = @mr_name"
            Create_cmd.Connection = Login.Connection
            Create_cmd.ExecuteNonQuery()


            ''sending a message
            'If enable_mess = True And String.Equals("Y", released_t) = True Then
            '    'send APL notification to procurement and inventory and mfg

            '    'write mail
            '    Dim mail_n As String : mail_n = "Shipping Address has been updated for Project:  " & job & vbCrLf & vbCrLf _
            ' & "New Shipping Address: " & Ship_b.Text & vbCrLf _
            ' & "Updated by: " & current_user & vbCrLf _
            ' & "Updated Date: " & Now & vbCrLf


            '    mail_n = mail_n.Trim()


            '    Call Sent_mail.Sent_multiple_emails("Procurement", "Shipping Address has been updated for Project:  " & job, mail_n)
            '    Call Sent_mail.Sent_multiple_emails("Procurement Management", "Shipping Address has been updated for Project:  " & job, mail_n)
            '    Call Sent_mail.Sent_multiple_emails("Inventory", "Shipping Address has been updated for Project:  " & job, mail_n)
            '    Call Sent_mail.Sent_multiple_emails("Manufacturing", "Shipping Address has been updated for Project:  " & job, mail_n)
            '    Call Sent_mail.Sent_multiple_emails("General Management", "Shipping Address has been updated for Project:  " & job, mail_n)

            '    'add email addresses
            '    Dim emails_addr As New List(Of String)()

            '    'procurement
            '    emails_addr.Add("ecoy@atronixengineering.com")
            '    emails_addr.Add("fvargas@atronixengineering.com")
            '    emails_addr.Add("mmorris@atronixengineering.com")
            '    emails_addr.Add("sowens@atronixengineering.com")

            '    'mfg
            '    emails_addr.Add("shenley@atronixengineering.com")
            '    emails_addr.Add("mowens@atronixengineering.com")



            '    ' For i = 0 To emails_addr.Count - 1
            '    Try

            '        Dim e_mail As New MailMessage()
            '        e_mail = New MailMessage()
            '        e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

            '        For i = 0 To emails_addr.Count - 1
            '            e_mail.To.Add(emails_addr.Item(i))
            '        Next
            '        e_mail.Subject = "Shipping Address has been updated for Project:  " & job
            '        e_mail.IsBodyHtml = False
            '        e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

            '        Smtp_Server.Send(e_mail)

            '    Catch error_t As Exception
            '        MsgBox(error_t.ToString)
            '    End Try
            '    ' Next
            'End If

            MessageBox.Show("Shipping Address Updated")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub shipbox_SelectedValueChanged(sender As Object, e As EventArgs) Handles shipbox.SelectedValueChanged
        '--check if the combo boxes are null

        Dim choice_ship As String : choice_ship = "Project Address"

        If Not shipbox.SelectedItem Is Nothing Then
            choice_ship = shipbox.SelectedItem.ToString
        Else
            choice_ship = "Project Address"
        End If

        If String.Equals(choice_ship, "Other") = True Then
            Ship_b.ReadOnly = False
            Ship_b.Text = "Please type your Shipping Address"


        ElseIf String.Equals(choice_ship, "Project Address") = True Then
            Ship_b.ReadOnly = True
            Ship_b.Text = project_ship

        ElseIf String.Equals(choice_ship, "Office suite 110") = True Then
            Ship_b.ReadOnly = True
            Ship_b.Text = "3100 Medlock Bridge Rd NW # 110, Norcross, GA 30071"

        ElseIf String.Equals(choice_ship, "Office suite 220") = True Then
            Ship_b.ReadOnly = True
            Ship_b.Text = "3100 Medlock Bridge Rd NW # 220, Norcross, GA 30071"
        End If

    End Sub
End Class