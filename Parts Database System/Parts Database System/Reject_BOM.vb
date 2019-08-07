Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Reject_BOM

    Public Smtp_Server As New SmtpClient


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim result As DialogResult = MessageBox.Show("Are you sure, you want to reject this BOM?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message

        If (result = DialogResult.Yes) Then

            'write mail
            Dim mail_n As String : mail_n = "BOM Rejected!!!" & vbCrLf _
                & "Project: " & Procurement_Overview.job_label.Text & vbCrLf _
                 & "Reason for rejection: " & reason_t.Text & vbCrLf _
                 & "No further action from the Procurement, Inventory and Manufacturing department will take place until the BOM gets amended" & vbCrLf

            Call Sent_mail.Sent_multiple_emails("General Management", Procurement_Overview.job_label.Text & " BOM Rejected!!!", mail_n)

            '--- get who send it -----------
            Dim cmd411 As New MySqlCommand
            Dim bom_creator As String : bom_creator = "notfound"
            cmd411.Parameters.AddWithValue("@mr_name", Procurement_Overview.Text)
            cmd411.CommandText = "SELECT released_by from Material_Request.mr where mr_name = @mr_name"
            cmd411.Connection = Login.Connection
            Dim reader411 As MySqlDataReader
            reader411 = cmd411.ExecuteReader

            If reader411.HasRows Then

                While reader411.Read
                    bom_creator = reader411(0)
                End While

            End If

            reader411.Close()

            '---- get user email --------'
            Dim cmd41 As New MySqlCommand
            Dim email_user As String : email_user = "notfound"
            cmd41.Parameters.AddWithValue("@user", bom_creator)
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


            'add email addresses
            Dim emails_addr As New List(Of String)()

            If String.Equals(email_user, "notfound") = False Then
                emails_addr.Add(email_user)
            End If

            emails_addr.Add("TBullard@atronixengineering.com")
            ''procurement
            emails_addr.Add("ecoy@atronixengineering.com")
            emails_addr.Add("fvargas@atronixengineering.com")
            ''mfg
            emails_addr.Add("shenley@atronixengineering.com")
            emails_addr.Add("mowens@atronixengineering.com")
            ''inventory
            emails_addr.Add("dnix@atronixengineering.com")
            emails_addr.Add("dmoore@atronixengineering.com")
            'test
            ' emails_addr.Add("bdahlqvist@atronixengineering.com")

            Try

                Dim e_mail As New MailMessage()
                e_mail = New MailMessage()
                e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                For i = 0 To emails_addr.Count - 1
                    e_mail.To.Add(emails_addr.Item(i))
                Next

                e_mail.Subject = Procurement_Overview.job_label.Text & " BOM Rejected!!!"
                e_mail.IsBodyHtml = False
                e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

                Smtp_Server.Send(e_mail)

            Catch error_t As Exception
                MsgBox(error_t.ToString)
            End Try

            MessageBox.Show("Done")
            Me.Visible = False


        End If

    End Sub

    Private Sub Reject_BOM_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"
    End Sub
End Class