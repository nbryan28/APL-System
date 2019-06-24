Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class Description_panels

    Public job_n As String
    Public Smtp_Server As New SmtpClient

    Private Sub Description_panels_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'host properties
        Smtp_Server.UseDefaultCredentials = False
        Smtp_Server.Credentials = New Net.NetworkCredential("inventory.atronix.system@gmail.com", "atronixatl123")
        Smtp_Server.Port = 587
        Smtp_Server.EnableSsl = True
        Smtp_Server.Host = "smtp.gmail.com"


        Panel_grid.Rows.Clear()

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@MBOM", Me.Text)
            cmd4.CommandText = "SELECT Panel_name, panel_desc, need_date, job from Material_Request.mr where MBOM = @MBOM and BOM_Type = 'Panel'"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    Panel_grid.Rows.Add(New String() {})
                    Panel_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    Panel_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    Panel_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    job_n = reader4(3)
                    i = i + 1
                End While

            End If

            reader4.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If Panel_grid.Rows.Count > 0 Then

            For i = 0 To Panel_grid.Rows.Count - 1

                If String.IsNullOrEmpty(Panel_grid.Rows(i).Cells(1).Value) = False Then

                    '---- update in mr ----------
                    Dim Create_cmd As New MySqlCommand
                    Create_cmd.Parameters.AddWithValue("@MBOM", Me.Text)
                    Create_cmd.Parameters.AddWithValue("@Panel_name", Panel_grid.Rows(i).Cells(0).Value)
                    Create_cmd.Parameters.AddWithValue("@Panel_desc", Panel_grid.Rows(i).Cells(1).Value)
                    Create_cmd.Parameters.AddWithValue("@need_date", Panel_grid.Rows(i).Cells(2).Value)

                    Create_cmd.CommandText = "UPDATE Material_Request.mr SET Panel_desc = @Panel_desc, need_date = @need_date where MBOM = @MBOM and Panel_name = @Panel_name"
                    Create_cmd.Connection = Login.Connection
                    Create_cmd.ExecuteNonQuery()

                    '-----  update in packing list --

                    Dim n_bom2 As Double : n_bom2 = 0
                    Dim check_cmd3 As New MySqlCommand
                    check_cmd3.Parameters.AddWithValue("@job", job_n)
                    check_cmd3.CommandText = "select count(distinct n_r) from Master_Packing_List.packing_l where job = @job"

                    check_cmd3.Connection = Login.Connection
                    check_cmd3.ExecuteNonQuery()

                    Dim reader3 As MySqlDataReader
                    reader3 = check_cmd3.ExecuteReader

                    If reader3.HasRows Then
                        While reader3.Read
                            n_bom2 = reader3(0) - 1
                        End While
                    End If

                    reader3.Close()

                    Dim Create_cmd3 As New MySqlCommand
                    Create_cmd3.Parameters.AddWithValue("@n_r", n_bom2)
                    Create_cmd3.Parameters.AddWithValue("@part_name", Panel_grid.Rows(i).Cells(0).Value)
                    Create_cmd3.Parameters.AddWithValue("@part_desc", Panel_grid.Rows(i).Cells(1).Value)
                    Create_cmd3.Parameters.AddWithValue("@need_date", Panel_grid.Rows(i).Cells(2).Value)
                    Create_cmd3.Parameters.AddWithValue("@job", job_n)

                    Create_cmd3.CommandText = "UPDATE Master_Packing_List.packing_l SET part_desc = @part_desc where part_name = @part_name and need_date = @need_date and job = @job and n_r = @n_r"
                    Create_cmd3.Connection = Login.Connection
                    Create_cmd3.ExecuteNonQuery()


                    '------ update in build request -----
                    Dim n_bom As Double : n_bom = 0
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.AddWithValue("@job", job_n)
                    check_cmd.CommandText = "select count(distinct n_r) from Build_request.build_r where job = @job"

                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        While reader.Read
                            n_bom = reader(0) - 1
                        End While
                    End If

                    reader.Close()

                    Dim Create_cmd2 As New MySqlCommand
                    Create_cmd2.Parameters.AddWithValue("@n_r", n_bom)
                    Create_cmd2.Parameters.AddWithValue("@Panel_name", Panel_grid.Rows(i).Cells(0).Value)
                    Create_cmd2.Parameters.AddWithValue("@Panel_desc", Panel_grid.Rows(i).Cells(1).Value)
                    Create_cmd2.Parameters.AddWithValue("@need_date", Panel_grid.Rows(i).Cells(2).Value)
                    Create_cmd2.Parameters.AddWithValue("@job", job_n)

                    Create_cmd2.CommandText = "UPDATE Build_request.build_r SET panel_desc = @Panel_desc where panel = @Panel_name and need_date = @need_date and job = @job and n_r = @n_r"
                    Create_cmd2.Connection = Login.Connection
                    Create_cmd2.ExecuteNonQuery()

                    '-----------------------------------

                End If
            Next

            '------- send email to MFG --------

            '////////////////////// ---------------------  notify -----------------------
            If enable_mess = True Then

                Dim mail_n As String : mail_n = "Panels Descriptions have been updated for project" & job_n & vbCrLf & vbCrLf _
             & "Revised by: " & current_user


                '--- sent email-------
                'add email addresses
                Dim emails_addr As New List(Of String)()

                'mfg
                emails_addr.Add("shenley@atronixengineering.com")
                emails_addr.Add("mowens@atronixengineering.com")


                Try
                    Dim e_mail As New MailMessage()
                    e_mail = New MailMessage()
                    e_mail.From = New MailAddress("inventory.atronix.system@gmail.com")

                    For j = 0 To emails_addr.Count - 1
                        e_mail.To.Add(emails_addr.Item(j))
                    Next

                    e_mail.Subject = "Panels Descriptions have been updated for project" & job_n
                    e_mail.IsBodyHtml = False
                    e_mail.Body = mail_n & vbCrLf & "APL (Atronix Pricing and Logistics System)  | Warehouse and Distribution" & vbCrLf & "3100 Medlock Bridge Rd  |  Ste 110  |  Peachtree Corners, GA 30071" & vbCrLf & "ATRONIX"
                    Smtp_Server.Send(e_mail)

                Catch error_t As Exception
                    MsgBox(error_t.ToString)
                End Try
                '   Next

            End If
            '---------------------

            '----------------------------------


            MessageBox.Show("Panel Update Complete")
        End If


        If BOM_types.Visible = True Then
            BOM_types.populate_panel_bom()
            BOM_types.Panel_grid.Rows.Clear()
        End If


        Me.Visible = False


    End Sub
End Class