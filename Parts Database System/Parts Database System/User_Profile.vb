Imports MySql.Data.MySqlClient

Public Class User_Profile

    Public plabels() As Label
    Public alabels() As Label
    Public tronco() As Label
    Public tronco2() As Label


    Private Sub User_Profile_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            Dim usern As String : usern = "Unknown"
            '--- load user data ------
            Dim check As New MySqlCommand
            check.Parameters.AddWithValue("@username", current_user)
            check.CommandText = "select username, Role, email from users where username = @username"
            check.Connection = Login.Connection
            check.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check.ExecuteReader

            If reader.HasRows Then

                While reader.Read

                    usern = reader(0).ToString
                    user_l.Text = reader(0).ToString
                    role_l.Text = reader(1).ToString
                    email_l.Text = reader(2).ToString

                End While
            End If

            reader.Close()


            Try
                Dim ch As Char = usern(0)


                If Not Image.FromFile("O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\user_ini\" & ch & ".png") Is Nothing Then
                    Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\\Users\Bryan Dahlqvist\Parts Database System\user_ini\" & ch & ".png")
                    Label1.Image = bmp1
                End If

            Catch ex As Exception
                '----------------- default image -----------------
                Dim bmp1 As New System.Drawing.Bitmap("O:\atlanta\\Users\Bryan Dahlqvist\Parts Database System\user_ini\a.png")
                Label1.Image = bmp1
            End Try

            Me.Text = "Welcome " & usern

            '---load comboboxes ---

            If String.Equals(current_user, "tbullard") = True Then


            Else


                Dim check_cmd1 As New MySqlCommand
                check_cmd1.Parameters.AddWithValue("@Employee_Name", current_user)
                check_cmd1.CommandText = "select Job_number from management.assignments where Employee_Name = @Employee_Name"
                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    While reader1.Read
                        ComboBox1.Items.Add(reader1(0).ToString)
                        ComboBox2.Items.Add(reader1(0).ToString)
                        ComboBox3.Items.Add(reader1(0).ToString)
                    End While
                End If

                reader1.Close()


            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        Call Load_BOM_map("110019_Master_BOM_rev32", "110019")
        Call Color_Module("110019")
        ' Call Load_BOM_map("110019_Master_BOM_rev32", "110019")
        ' Call Color_Module("110019")

    End Sub

    Sub Load_BOM_map(mr As String, job As String)
        '--- Load BOM Map ---

        ToolTip1.SetToolTip(MBOM, BOM_info("Master BOM", job))
        ToolTip1.SetToolTip(F_BOM, BOM_info("Field BOM", job))
        ToolTip1.SetToolTip(SP_BOM, BOM_info("Spare Parts BOM", job))

        Dim list As New List(Of String)  'list of panels
        Dim list_as As New List(Of String)  'list of assemblies

        Try

            '--- add panel names to list
            Dim cmd_j As New MySqlCommand
            cmd_j.Parameters.AddWithValue("@MBOM", mr)
            cmd_j.CommandText = "SELECT Panel_name from Material_Request.mr where BOM_type = 'Panel' and MBOM = @MBOM"
            cmd_j.Connection = Login.Connection

            Dim readerj As MySqlDataReader
            readerj = cmd_j.ExecuteReader

            If readerj.HasRows Then
                While readerj.Read
                    list.Add(readerj(0).ToString)
                End While
            End If

            readerj.Close()

            '--------- add assemblies --------
            Dim cmd_j2 As New MySqlCommand
            cmd_j2.Parameters.AddWithValue("@MBOM", mr)
            cmd_j2.CommandText = "SELECT Panel_name from Material_Request.mr where BOM_type = 'Assembly' and MBOM = @MBOM"
            cmd_j2.Connection = Login.Connection

            Dim readerj2 As MySqlDataReader
            readerj2 = cmd_j2.ExecuteReader

            If readerj2.HasRows Then
                While readerj2.Read
                    list_as.Add(readerj2(0).ToString)
                End While
            End If

            readerj2.Close()
            '---------------------------------------
            DoubleBuffered = True
            Me.SetStyle(ControlStyles.UserPaint, True)
            Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
            Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

            '---- initial points ------
            Dim x_i As Integer : x_i = 0
            Dim y_i As Integer : y_i = 0
            Dim w_i As Integer : w_i = 0
            Dim l_i As Integer : l_i = 0

            Dim w_ia As Integer : w_ia = 0
            Dim l_ia As Integer : l_ia = 0
            Dim x_ia As Integer : x_ia = 0
            Dim y_ia As Integer : y_ia = 0

            w_i = P_BOM.Width
            l_i = P_BOM.Height

            w_ia = A_BOM.Width
            l_ia = A_BOM.Height

            x_i = P_BOM.Location.X
            y_i = P_BOM.Location.Y

            x_ia = A_BOM.Location.X
            y_ia = A_BOM.Location.Y

            Dim ini_y As Integer : ini_y = x_i + CType((w_i) / 2, Integer)
            Dim ini_z As Integer : ini_z = y_i + l_i

            '   Dim plabels(list.Count - 1) As Label

            ReDim plabels(list.Count - 1)
            ReDim tronco(list.Count - 1)

            '   Dim alabels(list_as.Count - 1) As Label
            ReDim alabels(list_as.Count - 1)
            ReDim tronco2(list_as.Count - 1)

            Dim y As Integer : y = ini_z '384
            Dim z As Integer : z = ini_z + 22 '406  

            Dim y2 As Integer : y2 = ini_z '384
            Dim z2 As Integer : z2 = ini_z + 22 '406  

            '--- add branches and panel blocks---
            For i = 0 To list.Count - 1

                'branches
                tronco(i) = New Label()
                tronco(i).Text = ""
                tronco(i).Location = New Point(ini_y, y) '286
                tronco(i).BackColor = Color.DimGray
                tronco(i).Size = New System.Drawing.Size(10, 22)
                tronco(i).Anchor = AnchorStyles.Top

                Panel1.Controls.Add(tronco(i))
                y = y + 56

                'panel blocks

                plabels(i) = New Label()
                plabels(i).Size = New System.Drawing.Size(97, 38)   '267,  74
                plabels(i).BackColor = Color.Teal
                plabels(i).Text = list.Item(i).ToString
                plabels(i).ForeColor = Color.WhiteSmoke
                plabels(i).Font = New Font("Segoe UI", 9, FontStyle.Regular)
                plabels(i).Anchor = AnchorStyles.Top
                plabels(i).Location = New Point(ini_y - CType((w_i) / 2, Integer), z)  '161
                plabels(i).TextAlign = ContentAlignment.MiddleCenter


                Panel1.Controls.Add(plabels(i))
                z = z + 56

                Dim temp_i As Integer : temp_i = i



                ToolTip1.SetToolTip(plabels(i), BOM_info(plabels(i).Text, job))
            Next

            '------------------


            '----- add assem -----
            For i = 0 To list_as.Count - 1

                'branches
                tronco2(i) = New Label()
                tronco2(i).Text = ""
                tronco2(i).Location = New Point(x_ia + CType((w_ia) / 2, Integer), y2)  '942
                tronco2(i).BackColor = Color.DimGray
                tronco2(i).Size = New System.Drawing.Size(10, 22)
                tronco2(i).Anchor = AnchorStyles.Top

                Panel1.Controls.Add(tronco2(i))
                y2 = y2 + 56

                'panel blocks

                alabels(i) = New Label()
                alabels(i).Size = New System.Drawing.Size(97, 38)   '267, 74
                alabels(i).BackColor = Color.Teal
                alabels(i).Text = list_as.Item(i).ToString
                alabels(i).ForeColor = Color.WhiteSmoke
                alabels(i).Font = New Font("Segoe UI", 9, FontStyle.Regular)
                alabels(i).Anchor = AnchorStyles.Top
                alabels(i).Location = New Point(x_ia, z2)  '817
                alabels(i).TextAlign = ContentAlignment.MiddleCenter

                Panel1.Controls.Add(alabels(i))
                z2 = z2 + 56

                Dim temp_i As Integer : temp_i = i



                ToolTip1.SetToolTip(alabels(i), BOM_info(alabels(i).Text, job))
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Function BOM_info(name As String, job As String) As String

        Dim created_by As String : created_by = ""
        Dim date_created As String : date_created = ""
        Dim date_released As String : date_released = ""
        Dim mr_name As String : mr_name = ""

        Try

            If String.Equals("Field BOM", name) = True Then

                Dim cmd_j2 As New MySqlCommand
                cmd_j2.Parameters.AddWithValue("@job", job)
                cmd_j2.CommandText = "SELECT created_by, Date_Created, release_date, mr_name  from Material_Request.mr where BOM_type = 'Field' and job = @job"
                cmd_j2.Connection = Login.Connection

                Dim readerj2 As MySqlDataReader
                readerj2 = cmd_j2.ExecuteReader

                If readerj2.HasRows Then
                    While readerj2.Read
                        created_by = readerj2(0).ToString
                        date_created = readerj2(1).ToString
                        date_released = readerj2(2).ToString
                        mr_name = readerj2(3).ToString
                    End While
                End If

                readerj2.Close()

            ElseIf String.Equals("Spare Parts BOM", name) = True Then

                Dim cmd_j2 As New MySqlCommand
                cmd_j2.Parameters.AddWithValue("@Panel_name", name)
                cmd_j2.Parameters.AddWithValue("@job", job)
                cmd_j2.CommandText = "SELECT created_by, Date_Created, release_date, mr_name  from Material_Request.mr where BOM_type = 'Spare_Parts' and job = @job"
                cmd_j2.Connection = Login.Connection

                Dim readerj2 As MySqlDataReader
                readerj2 = cmd_j2.ExecuteReader

                If readerj2.HasRows Then
                    While readerj2.Read
                        created_by = readerj2(0).ToString
                        date_created = readerj2(1).ToString
                        date_released = readerj2(2).ToString
                        mr_name = readerj2(3).ToString
                    End While
                End If

                readerj2.Close()

            ElseIf String.Equals("Master BOM", name) = True Then

                Dim cmd_j2 As New MySqlCommand
                cmd_j2.Parameters.AddWithValue("@Panel_name", name)
                cmd_j2.Parameters.AddWithValue("@job", job)
                cmd_j2.CommandText = "SELECT created_by, Date_Created, release_date, mr_name  from Material_Request.mr where BOM_type = 'MB' and job = @job"
                cmd_j2.Connection = Login.Connection

                Dim readerj2 As MySqlDataReader
                readerj2 = cmd_j2.ExecuteReader

                If readerj2.HasRows Then
                    While readerj2.Read
                        created_by = readerj2(0).ToString
                        date_created = readerj2(1).ToString
                        date_released = readerj2(2).ToString
                        mr_name = readerj2(3).ToString
                    End While
                End If

                readerj2.Close()

            Else
                Dim cmd_j2 As New MySqlCommand
                cmd_j2.Parameters.AddWithValue("@Panel_name", name)
                cmd_j2.Parameters.AddWithValue("@job", job)
                cmd_j2.CommandText = "SELECT created_by, Date_Created, release_date, mr_name  from Material_Request.mr where Panel_name = @Panel_name and job = @job"
                cmd_j2.Connection = Login.Connection

                Dim readerj2 As MySqlDataReader
                readerj2 = cmd_j2.ExecuteReader

                If readerj2.HasRows Then
                    While readerj2.Read
                        created_by = readerj2(0).ToString
                        date_created = readerj2(1).ToString
                        date_released = readerj2(2).ToString
                        mr_name = readerj2(3).ToString
                    End While
                End If

                readerj2.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        BOM_info = mr_name & vbCrLf & vbCrLf & "Created by: " & created_by & vbCrLf & vbCrLf & " Date Created: " & date_created _
        & vbCrLf & vbCrLf & " Date Released: " & date_released


    End Function

    Sub Color_Module(job As String)

        '--- This sub color the blocks of the BOM Map
        Dim MBOM_p As Double : MBOM_p = percentage_full("Master BOM", job)
        Dim Field_p As Double : Field_p = percentage_full("Field BOM", job)
        Dim SP_p As Double : SP_p = percentage_full("Spare Parts BOM", job)

        If MBOM_p = 100 Then
            MBOM.BackColor = Color.Teal
        ElseIf MBOM_p = 0 Then
            MBOM.BackColor = Color.DarkRed
        Else
            MBOM.BackColor = Color.DarkGoldenrod
        End If

        '-----------
        If Field_p = 100 Then
            F_BOM.BackColor = Color.Teal
        ElseIf Field_p = 0 And in_progress_bom_color("Field BOM", job) = False Then
            F_BOM.BackColor = Color.DarkRed
        Else
            F_BOM.BackColor = Color.DarkGoldenrod
        End If

        '----------
        If SP_p = 100 Then
            SP_BOM.BackColor = Color.Teal
        ElseIf SP_p = 0 And in_progress_bom_color("Spare Parts BOM", job) = False Then
            SP_BOM.BackColor = Color.DarkRed
        Else
            SP_BOM.BackColor = Color.DarkGoldenrod
        End If



        For i = 0 To plabels.Count - 1
            Dim P_p As Double : P_p = percentage_full(plabels(i).Text, job)

            If P_p = 100 Then
                plabels(i).BackColor = Color.Teal
            ElseIf P_p = 0 And in_progress_bom_color(plabels(i).Text, job) = False Then
                plabels(i).BackColor = Color.DarkRed
            Else
                plabels(i).BackColor = Color.DarkGoldenrod
            End If
        Next



        For i = 0 To alabels.Count - 1
            Dim AP_p As Double : AP_p = percentage_full(alabels(i).Text, job)

            If AP_p = 100 Then
                alabels(i).BackColor = Color.Teal
            ElseIf AP_p = 0 Then
                alabels(i).BackColor = Color.DarkRed
            Else
                alabels(i).BackColor = Color.DarkGoldenrod
            End If
        Next

    End Sub

    Function in_progress_bom_color(panel As String, job As String) As Boolean
        'if return false then no progress ha been made

        in_progress_bom_color = False


        Dim cmd3 As New MySqlCommand
        cmd3.Parameters.AddWithValue("@panel_name", panel)
        cmd3.Parameters.AddWithValue("@job", job)
        cmd3.CommandText = "select * from Material_Request.my_assem where panel_name = @panel_name and job = @job"
        cmd3.Connection = Login.Connection
        Dim reader3 As MySqlDataReader
        reader3 = cmd3.ExecuteReader

        If reader3.HasRows Then
            While reader3.Read
                in_progress_bom_color = True
            End While
        End If

        reader3.Close()

    End Function

    Function percentage_full(name As String, open_job As String) As Double

        percentage_full = 100

        Dim p_qty As Integer : p_qty = 1  'panel quantity multiplier

        Try
            '---------- calculate percentage ---------
            Dim complete_mr As Double : complete_mr = 0
            Dim total_qty As Double : total_qty = 0
            Dim fullf As Double : fullf = 0
            Dim mr_name As String : mr_name = ""

            '---get mr_name from name
            If String.Equals(name, "Field BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "Field")
                cmd5.Parameters.AddWithValue("@job", open_job)

                cmd5.CommandText = "SELECT mr_name from Material_Request.mr where  BOM_type = @BOM_type and job = @job"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        mr_name = reader5(0).ToString
                    End While
                End If
                reader5.Close()

                mr_name = Procurement_Overview.get_last_revision(mr_name)

            ElseIf String.Equals(name, "Master BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "MB")
                cmd5.Parameters.AddWithValue("@job", open_job)

                cmd5.CommandText = "SELECT mr_name from Material_Request.mr where  BOM_type = @BOM_type and job = @job"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        mr_name = reader5(0).ToString
                    End While
                End If
                reader5.Close()
                mr_name = Procurement_Overview.get_last_revision(mr_name)


            ElseIf String.Equals(name, "Spare Parts BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "Spare_Parts")
                cmd5.Parameters.AddWithValue("@job", open_job)

                cmd5.CommandText = "SELECT mr_name from Material_Request.mr where  BOM_type = @BOM_type and job = @job"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        mr_name = reader5(0).ToString
                    End While
                End If
                reader5.Close()
                mr_name = Procurement_Overview.get_last_revision(mr_name)



            Else

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@job", open_job)
                cmd5.Parameters.AddWithValue("@Panel_name", name)
                cmd5.CommandText = "SELECT mr_name from Material_Request.mr where Panel_name = @Panel_name and job = @job"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        mr_name = reader5(0).ToString
                    End While
                End If
                reader5.Close()

                mr_name = Procurement_Overview.get_last_revision(mr_name)

                '-- get panel qty -----
                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", mr_name)
                cmd419.CommandText = "SELECT Panel_qty, BOM_type from Material_Request.mr where mr_name = @mr_name"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read
                        If String.Equals(reader419(1).ToString, "Assembly") = False Then
                            p_qty = If(reader419(0) Is DBNull.Value, 1, reader419(0))
                        Else
                            p_qty = 1
                        End If
                    End While
                End If

                reader419.Close()
                '--------------------
            End If


            '-----------------------------

            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@mr_name", mr_name)
            cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    If IsDBNull(reader3(0)) Then
                        fullf = 0
                    Else
                        fullf = CType(reader3(0), Double)
                    End If
                End While
            End If

            reader3.Close()

            Dim cmdx As New MySqlCommand
            cmdx.Parameters.AddWithValue("@mr_name", mr_name)
            cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
            cmdx.Connection = Login.Connection
            Dim readerx As MySqlDataReader
            readerx = cmdx.ExecuteReader

            If readerx.HasRows Then
                While readerx.Read
                    If IsDBNull(readerx(0)) Then
                        total_qty = 0
                    Else
                        total_qty = CType(readerx(0), Double) * p_qty
                    End If
                End While
            End If

            readerx.Close()

            If total_qty > 0 Then
                percentage_full = Math.Floor((fullf / total_qty) * 100)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Function

    Private Sub ComboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged


        Panel1.Controls.Clear()


    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub User_Profile_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class