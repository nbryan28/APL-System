Imports MySql.Data.MySqlClient
Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO

Public Class Need_by_date

    Public n_r As Integer
    Public br_name As String
    Public Smtp_Server As New SmtpClient


    Private Sub Need_by_date_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"

        br_name = "-"

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", My_Material_r.Text)
            cmd4.CommandText = "SELECT Panel_name from Material_Request.mr where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    panel_n.Text = reader4(0).ToString
                End While
            End If

            reader4.Close()

            '---------------------------------

            '--get latest revision --
            n_r = 0
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@job", My_Material_r.open_job)
            cmd41.CommandText = "SELECT distinct n_r from Build_request.build_r where job = @job order by n_r desc limit 1"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    n_r = reader41(0).ToString
                End While
            End If

            reader41.Close()
            '----------------------

            Dim cmd42 As New MySqlCommand
            cmd42.Parameters.AddWithValue("@job", My_Material_r.open_job)
            cmd42.Parameters.AddWithValue("@n_r", n_r)
            cmd42.Parameters.AddWithValue("@panel", panel_n.Text)
            cmd42.CommandText = "SELECT need_date, br_name from Build_request.build_r where job = @job and n_r = @n_r and panel = @panel"
            cmd42.Connection = Login.Connection
            Dim reader42 As MySqlDataReader
            reader42 = cmd42.ExecuteReader

            If reader42.HasRows Then
                While reader42.Read
                    date_n.Text = reader42(0).ToString
                    br_name = reader42(1).ToString
                End While
            End If

            reader42.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-------- update need_by_date in build_mr ----
        Try

            If String.IsNullOrEmpty(panel_n.Text) = False Then

                Dim Create_cmd2 As New MySqlCommand
                Create_cmd2.Parameters.AddWithValue("@br_name", br_name)
                Create_cmd2.Parameters.AddWithValue("@panel", panel_n.Text)
                Create_cmd2.Parameters.AddWithValue("@job", My_Material_r.open_job)
                Create_cmd2.Parameters.AddWithValue("@n_r", n_r)
                Create_cmd2.Parameters.AddWithValue("@need_date", MonthCalendar1.SelectionRange.Start)


                Create_cmd2.CommandText = "UPDATE Build_request.build_r SET need_date = @need_date where panel = @panel and job = @job and n_r = @n_r and br_name = @br_name"
                Create_cmd2.Connection = Login.Connection
                Create_cmd2.ExecuteNonQuery()

                '---------------- update MPL as well ------------------------
                Dim Create_cmd3 As New MySqlCommand
                Create_cmd3.Parameters.AddWithValue("@mpl_name", My_Material_r.open_job & "_Master_Packing_List")
                Create_cmd3.Parameters.AddWithValue("@panel", panel_n.Text)
                Create_cmd3.Parameters.AddWithValue("@job", My_Material_r.open_job)
                Create_cmd3.Parameters.AddWithValue("@n_r", n_r)
                Create_cmd3.Parameters.AddWithValue("@need_date", MonthCalendar1.SelectionRange.Start)


                Create_cmd3.CommandText = "UPDATE Master_Packing_List.packing_l SET need_date = @need_date where part_name = @panel and job = @job and n_r = @n_r and mpl_name = @mpl_name"
                Create_cmd3.Connection = Login.Connection
                Create_cmd3.ExecuteNonQuery()

                '----------------------------------------------

                If enable_mess = True Then
                    'send APL notification to procurement and inventory and mfg

                    'write mail
                    Dim mail_n As String : mail_n = "Need by date has been updated for Project:  " & My_Material_r.open_job & vbCrLf & vbCrLf _
                      & "old need by date: " & date_n.Text & vbCrLf _
                      & "new need by date: " & MonthCalendar1.SelectionRange.Start & vbCrLf _
                 & "Updated by: " & current_user & vbCrLf _



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
                    emails_addr.Add("jtuttle@atronixengineering.com")
                    emails_addr.Add("dshipman@atronixengineering.com")
                    emails_addr.Add("tbullard@atronixengineering.com")


                    ' For i = 0 To emails_addr.Count - 1
                    Try

                        Dim e_mail As New MailMessage()
                        e_mail = New MailMessage()
                        e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                        For i = 0 To emails_addr.Count - 1
                            e_mail.To.Add(emails_addr.Item(i))
                        Next
                        e_mail.Subject = "Need by date has been updated for Project:  " & My_Material_r.open_job
                        e_mail.IsBodyHtml = False
                        e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"

                        Smtp_Server.Send(e_mail)

                    Catch error_t As Exception
                        MsgBox(error_t.ToString)
                    End Try
                    ' Next

                End If
                MessageBox.Show("Need by Date Updated")

            End If

            Me.Visible = False

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Need_by_date_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class