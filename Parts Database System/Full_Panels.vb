Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Full_Panels

    Public index_s As Integer


    Private Sub Full_Panels_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        index_s = -1

        Try

            panel_grid.Rows.Clear()
            ' need_box.Text = My_Material_r.general_date.Text

            '------ get vendor if parts exist ----
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@mr_name", Me.Text)
            cmd.CommandText = "SELECT panel_name, panel_desc, qty, need_by_date, draw_description from Material_Request.full_panels where mr_name = @mr_name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    panel_grid.Rows.Add(New String() {reader(0).ToString, reader(1).ToString, reader(2).ToString, reader(3).ToString, reader(4).ToString})
                End While
            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub pdf_grid_DoubleClick(sender As Object, e As EventArgs) Handles panel_grid.DoubleClick
        'Show panel data in its corresponding textboxes while double clicking

        If panel_grid.Rows.Count > 0 Then

            For j = 0 To panel_grid.Rows.Count - 1
                panel_grid.Rows(j).DefaultCellStyle.BackColor = Color.WhiteSmoke
            Next

            Dim index_k = panel_grid.CurrentCell.RowIndex

            panel_box.Text = If(String.IsNullOrEmpty(panel_grid.Rows(index_k).Cells(0).Value) = True, "", panel_grid.Rows(index_k).Cells(0).Value)
            desc_box.Text = If(String.IsNullOrEmpty(panel_grid.Rows(index_k).Cells(1).Value) = True, "", panel_grid.Rows(index_k).Cells(1).Value)
            qty_box.Text = If(String.IsNullOrEmpty(panel_grid.Rows(index_k).Cells(2).Value) = True, "", panel_grid.Rows(index_k).Cells(2).Value)
            need_box.Text = If(String.IsNullOrEmpty(panel_grid.Rows(index_k).Cells(3).Value) = True, "", panel_grid.Rows(index_k).Cells(3).Value)

            panel_grid.Rows(index_k).DefaultCellStyle.BackColor = Color.Tan
            index_s = index_k

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If String.IsNullOrEmpty(panel_box.Text) = False Then

            'add panel
            Try
                Dim Create_cmd2 As New MySqlCommand
                Create_cmd2.Parameters.AddWithValue("@mr_name", Me.Text)
                Create_cmd2.Parameters.AddWithValue("@job", If(String.IsNullOrEmpty(My_Material_r.open_job) = True, "Unknown", My_Material_r.open_job))
                Create_cmd2.Parameters.AddWithValue("@panel_name", If(String.IsNullOrEmpty(panel_box.Text) = True, "Unknown", panel_box.Text))
                Create_cmd2.Parameters.AddWithValue("@panel_desc", If(String.IsNullOrEmpty(desc_box.Text) = True, "Unknown", desc_box.Text))
                Create_cmd2.Parameters.AddWithValue("@qty", If(IsNumeric(qty_box.Text) = True, qty_box.Text, 0))
                Create_cmd2.Parameters.AddWithValue("@need_by_date", If(String.IsNullOrEmpty(need_box.Text) = True, "Unknown", need_box.Text))
                Create_cmd2.Parameters.AddWithValue("@added_by", current_user)
                Create_cmd2.Parameters.AddWithValue("@draw_description", My_Material_r.open_job & "_" & panel_box.Text)
                Create_cmd2.CommandText = "INSERT INTO Material_Request.full_panels (mr_name, job, panel_name, panel_desc, qty, need_by_date, added_by, draw_description) VALUES (@mr_name, @job,  @panel_name, @panel_desc, @qty, @need_by_date, @added_by, @draw_description);"
                Create_cmd2.Connection = Login.Connection
                Create_cmd2.ExecuteNonQuery()
                Call Refresh_tables()

                '--send notification of revision
                If send_revision(Me.Text) = True Then

                    Dim mail_n As String : mail_n = "Build Request and Packing List for Project " & My_Material_r.open_job & "  has been Revised" & vbCrLf & vbCrLf _
                       & "Released by: " & current_user & vbCrLf _
                       & "Released Date: " & Now & vbCrLf _
                       & "Project: " & My_Material_r.open_job & vbCrLf _
                       & "BOM File Name: " & My_Material_r.Text & vbCrLf


                    Call Sent_mail.Sent_multiple_emails("Procurement", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("Procurement Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("Manufacturing", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("General Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)

                End If


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Else
            MessageBox.Show("Please enter a panel name")
        End If
    End Sub

    Sub Refresh_tables()

        ''---------- REFRESH DATA IN GRID ----------------------
        panel_grid.Rows.Clear()

        Dim cmd4 As New MySqlCommand
        cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
        cmd4.CommandText = "SELECT panel_name, panel_desc, qty, need_by_date from Material_Request.full_panels where mr_name = @mr_name"
        cmd4.Connection = Login.Connection
        Dim reader4 As MySqlDataReader
        reader4 = cmd4.ExecuteReader

        If reader4.HasRows Then

            While reader4.Read
                panel_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString})
            End While

        End If

        reader4.Close()
        index_s = -1
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '- Remove panel

        '-------- Delete entry --------

        If index_s > -1 Then
            Try

                Dim Create_vd As New MySqlCommand
                Dim dlgR As DialogResult
                dlgR = MessageBox.Show("Are you sure you want to delete this Panel entry?", "Attention!", MessageBoxButtons.YesNo)

                If dlgR = DialogResult.Yes Then

                    Create_vd.Parameters.AddWithValue("@mr_name", Me.Text)
                    Create_vd.Parameters.AddWithValue("@panel_name", If(String.IsNullOrEmpty(panel_box.Text) = True, "Unknown", panel_box.Text))
                    Create_vd.Parameters.AddWithValue("@panel_desc", If(String.IsNullOrEmpty(desc_box.Text) = True, "Unknown", desc_box.Text))
                    Create_vd.Parameters.AddWithValue("@qty", If(String.IsNullOrEmpty(qty_box.Text) = True, "Unknown", qty_box.Text))
                    Create_vd.Parameters.AddWithValue("@need_by_date", If(String.IsNullOrEmpty(need_box.Text) = True, "Unknown", need_box.Text))
                    Create_vd.CommandText = "DELETE FROM Material_Request.full_panels where mr_name = @mr_name and panel_name = @panel_name and panel_desc = @panel_desc and qty = @qty and need_by_date = @need_by_date"
                    Create_vd.Connection = Login.Connection
                    Create_vd.ExecuteNonQuery()

                    Call Refresh_tables()

                    '--send notification of revision
                    If send_revision(Me.Text) = True Then

                        Dim mail_n As String : mail_n = "Build Request and Packing List for Project " & My_Material_r.open_job & "  has been Revised" & vbCrLf & vbCrLf _
                           & "Released by: " & current_user & vbCrLf _
                           & "Released Date: " & Now & vbCrLf _
                           & "Project: " & My_Material_r.open_job & vbCrLf _
                           & "BOM File Name: " & My_Material_r.Text & vbCrLf


                        Call Sent_mail.Sent_multiple_emails("Procurement", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Procurement Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("Manufacturing", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                        Call Sent_mail.Sent_multiple_emails("General Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)

                    End If

                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            MessageBox.Show("Please double click the entry you want to delete, and click the 'Remove' button")
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '------- Update entry ---------
        If index_s > -1 Then

            Try
                Dim Create_cmd As New MySqlCommand

                Create_cmd.Parameters.AddWithValue("@old_mr_name", Me.Text)
                Create_cmd.Parameters.AddWithValue("@old_panel_name", If(String.IsNullOrEmpty(panel_grid.Rows(index_s).Cells(0).Value) = True, "", panel_grid.Rows(index_s).Cells(0).Value))
                Create_cmd.Parameters.AddWithValue("@old_panel_desc", If(String.IsNullOrEmpty(panel_grid.Rows(index_s).Cells(1).Value) = True, "", panel_grid.Rows(index_s).Cells(1).Value))
                Create_cmd.Parameters.AddWithValue("@old_qty", If(String.IsNullOrEmpty(panel_grid.Rows(index_s).Cells(2).Value) = True, "", panel_grid.Rows(index_s).Cells(2).Value))
                Create_cmd.Parameters.AddWithValue("@old_need_by_date", If(String.IsNullOrEmpty(panel_grid.Rows(index_s).Cells(3).Value) = True, "", panel_grid.Rows(index_s).Cells(3).Value))

                '--- new values
                Create_cmd.Parameters.AddWithValue("@panel_name", panel_box.Text)
                Create_cmd.Parameters.AddWithValue("@panel_desc", desc_box.Text)
                Create_cmd.Parameters.AddWithValue("@qty", qty_box.Text)
                Create_cmd.Parameters.AddWithValue("@need_by_date", need_box.Text)
                Create_cmd.Parameters.AddWithValue("@draw_description", My_Material_r.open_job & "_" & "_" & panel_box.Text)
                Create_cmd.CommandText = "UPDATE Material_Request.full_panels  SET  panel_name = @panel_name, panel_desc = @panel_desc, qty = @qty, need_by_date = @need_by_date, draw_description = @draw_description where mr_name = @old_mr_name and panel_name = @old_panel_name and panel_desc = @old_panel_desc and qty = @old_qty and need_by_date = @old_need_by_date"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

                Call Refresh_tables()

                If send_revision(Me.Text) = True Then

                    Dim mail_n As String : mail_n = "Build Request and Packing List for Project " & My_Material_r.open_job & "  has been Revised" & vbCrLf & vbCrLf _
                       & "Released by: " & current_user & vbCrLf _
                       & "Released Date: " & Now & vbCrLf _
                       & "Project: " & My_Material_r.open_job & vbCrLf _
                       & "BOM File Name: " & My_Material_r.Text & vbCrLf


                    Call Sent_mail.Sent_multiple_emails("Procurement", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("Procurement Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("Manufacturing", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)
                    Call Sent_mail.Sent_multiple_emails("General Management", "Build Request and Packing List has been revised for Project " & My_Material_r.open_job, mail_n)

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ''-- attached a pdf drawing to a full panel
        'If index_s > -1 Then

        '    If String.Equals(Me.Text, "Add Full Panels") = False Then

        '        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then

        '            Dim DestPath As String = "O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Drawings\"

        '            Dim draw_description As String : draw_description = If(String.IsNullOrEmpty(panel_grid.Rows(index_s).Cells(0).Value) = True, "", panel_grid.Rows(index_s).Cells(0).Value)

        '            Dim valid_mr_name As String = Me.Text

        '            For Each c In Path.GetInvalidFileNameChars()
        '                If valid_mr_name.Contains(c) Then
        '                    valid_mr_name = valid_mr_name.Replace("", c)
        '                End If
        '            Next

        '            draw_description = My_Material_r.open_job & "_" & valid_mr_name & "_" & draw_description

        '            Call CopyDirectory(DestPath, OpenFileDialog1.FileName, My_Material_r.open_job, draw_description)

        '        End If
        '    End If
        'Else
        '    MessageBox.Show("Please double click the Panel you want to attached a pdf drawing to")
        'End If
    End Sub

    Sub CopyDirectory(destPath As String, filename_s As String, job As String, draw_desc As String)

        '-- copy file to APL drawing folder

        Dim path_s As String : path_s = Path.GetDirectoryName(filename_s)
        Dim only_f As String : only_f = Path.GetFileName(filename_s)

        '------- check if a similar file exist
        Dim count_m As Integer : count_m = 0

        Try
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@draw", draw_desc)
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "SELECT count(*) from management.draw_pdf where draw_description = @draw and job = @job"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    count_m = reader41(0).ToString
                End While
            End If

            reader41.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '----------------------------------------

        If count_m > 0 Then
            '-- if there are existing revisions
            Dim dlgR As DialogResult
            dlgR = MessageBox.Show("An existing Drawing associated with this object has been detected, This Drawing will be consider a revision. Do you want to proceed?", "Existing Drawing detected!", MessageBoxButtons.YesNo)

            If dlgR = DialogResult.Yes Then

            End If

        Else

            Dim valid_mr_name As String = Me.Text

            For Each c In Path.GetInvalidFileNameChars()
                If valid_mr_name.Contains(c) Then
                    valid_mr_name = valid_mr_name.Replace("", c)
                End If
            Next

            Dim desc_path As String : desc_path = destPath & "Drawing_" & valid_mr_name & "_" & only_f.Replace("/", "")

            If File.Exists(desc_path) = True Then
                MessageBox.Show("A Drawing with the same name already exist")

            Else
                File.Copy(filename_s, desc_path)

                Try
                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@job", job)
                    Create_cmd6.Parameters.AddWithValue("@draw_description", draw_desc)
                    Create_cmd6.Parameters.AddWithValue("@file_name", "Drawing_" & valid_mr_name & "_" & only_f.Replace("/", ""))
                    Create_cmd6.Parameters.AddWithValue("@uploaded_by", current_user)

                    Create_cmd6.CommandText = "INSERT INTO management.draw_pdf(job, draw_description, file_name , date_upload, uploaded_by) VALUES (@job, @draw_description, @file_name, now(), @uploaded_By)"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

                MessageBox.Show("Drawing upload successfully")
            End If
        End If


    End Sub

    Function send_revision(mr_name As String) As Boolean
        '-- check if we are dealing with a released mr

        send_revision = False

        Try
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT * from Material_Request.mr where mr_name = @mr_name and released = 'Y'"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    send_revision = True
                End While
            End If

            reader41.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function
End Class