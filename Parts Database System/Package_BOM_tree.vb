Imports MySql.Data.MySqlClient


Public Class Package_BOM_tree


    Private Sub Label1_MouseEnter(sender As Object, e As EventArgs) Handles MBOM.MouseEnter
        MBOM.BackColor = Color.SlateGray
    End Sub

    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs) Handles MBOM.MouseLeave
        MBOM.BackColor = Color.Teal
    End Sub

    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles F_BOM.MouseEnter
        F_BOM.BackColor = Color.SlateGray
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles F_BOM.MouseLeave
        F_BOM.BackColor = Color.Teal
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs)
        A_BOM.BackColor = Color.SlateGray
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs)
        A_BOM.BackColor = Color.Teal
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles SP_BOM.MouseEnter
        SP_BOM.BackColor = Color.SlateGray
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles SP_BOM.MouseLeave
        SP_BOM.BackColor = Color.Teal
    End Sub

    Private Sub Package_BOM_tree_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        'Try
        '    Dim check_cmd As New MySqlCommand
        '    check_cmd.CommandText = "select mr_name, job from Material_Request.mr"

        '    check_cmd.Connection = Login.Connection
        '    check_cmd.ExecuteNonQuery()

        '    Dim reader As MySqlDataReader
        '    reader = check_cmd.ExecuteReader

        '    If reader.HasRows Then
        '        While reader.Read
        '            open_grid.Rows.Add(New String() {reader(0).ToString, reader(1).ToString})
        '        End While
        '    End If

        '    reader.Close()
        'Catch ex As Exception
        '    MessageBox.Show(ex.ToString)
        'End Try



        '-----------------tool tip -------------------

        Dim mydata As String : mydata = "Master BOM: 117812_Master_BOM" & vbCrLf & vbCrLf & "Created by: username" & vbCrLf & vbCrLf & " Date Created: " & Today _
        & vbCrLf & vbCrLf & " Date Released: " & Today

        ToolTip1.IsBalloon = True
        'ToolTip1.SetToolTip(MBOM, mydata)
        'ToolTip1.SetToolTip(P_BOM, mydata)
        'ToolTip1.SetToolTip(F_BOM, mydata)
        'ToolTip1.SetToolTip(SP_BOM, mydata)
        '------------------------------------------------------

        ''---- initial points ------
        'Dim x_i As Integer : x_i = 0
        'Dim y_i As Integer : y_i = 0
        'Dim w_i As Integer : w_i = 0
        'Dim l_i As Integer : l_i = 0

        'w_i = P_BOM.Width
        'l_i = P_BOM.Height

        'x_i = P_BOM.Location.X
        'y_i = P_BOM.Location.Y

        'Dim ini_y As Integer : ini_y = x_i + CType((w_i) / 2, Integer)
        'Dim ini_z As Integer : ini_z = y_i + l_i

        'Dim plabels(list.Count - 1) As Label
        'Dim tronco(list.Count - 1) As Label

        'Dim alabels(list_as.Count - 1) As Label
        'Dim tronco2(list_as.Count - 1) As Label

        'Dim y As Integer : y = ini_z '384
        'Dim z As Integer : z = ini_z + 22 '406

        'Dim y2 As Integer : y2 = ini_z '384
        'Dim z2 As Integer : z2 = ini_z + 22 '406

        ''--- add branches and panel blocks---
        'For i = 0 To list.Count - 1

        '    'branches
        '    tronco(i) = New Label()
        '    tronco(i).Text = ""
        '    tronco(i).Location = New Point(286, y)
        '    tronco(i).BackColor = Color.DimGray
        '    tronco(i).Size = New System.Drawing.Size(10, 22)
        '    tronco(i).Anchor = AnchorStyles.Top

        '    Me.Controls.Add(tronco(i))
        '    y = y + 96

        '    'panel blocks

        '    plabels(i) = New Label()
        '    plabels(i).Size = New System.Drawing.Size(267, 74)   '267,  74
        '    plabels(i).BackColor = Color.Teal
        '    plabels(i).Text = list.Item(i).ToString
        '    plabels(i).ForeColor = Color.WhiteSmoke
        '    plabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
        '    plabels(i).Anchor = AnchorStyles.Top
        '    plabels(i).Location = New Point(161, z)
        '    plabels(i).TextAlign = ContentAlignment.MiddleCenter


        '    Me.Controls.Add(plabels(i))
        '    z = z + 96

        '    plabels(i).ContextMenuStrip = ContextMenuStrip1

        '    AddHandler plabels(i).MouseEnter, AddressOf block_color
        '    AddHandler plabels(i).MouseLeave, AddressOf block_normal

        '    ToolTip1.SetToolTip(plabels(i), mydata)
        'Next

        ''------------------


        ''----- add assem -----
        'For i = 0 To list_as.Count - 1

        '    'branches
        '    tronco2(i) = New Label()
        '    tronco2(i).Text = ""
        '    tronco2(i).Location = New Point(942, y2)
        '    tronco2(i).BackColor = Color.DimGray
        '    tronco2(i).Size = New System.Drawing.Size(10, 22)
        '    tronco2(i).Anchor = AnchorStyles.Top

        '    Me.Controls.Add(tronco2(i))
        '    y2 = y2 + 96

        '    'panel blocks

        '    alabels(i) = New Label()
        '    alabels(i).Size = New System.Drawing.Size(267, 74)   '267, 74
        '    alabels(i).BackColor = Color.Teal
        '    alabels(i).Text = list_as.Item(i).ToString
        '    alabels(i).ForeColor = Color.WhiteSmoke
        '    alabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
        '    alabels(i).Anchor = AnchorStyles.Top
        '    alabels(i).Location = New Point(817, z2)
        '    alabels(i).TextAlign = ContentAlignment.MiddleCenter

        '    Me.Controls.Add(alabels(i))
        '    z2 = z2 + 96
        '    alabels(i).ContextMenuStrip = ContextMenuStrip1

        '    AddHandler alabels(i).MouseEnter, AddressOf block_color
        '    AddHandler alabels(i).MouseLeave, AddressOf block_normal

        '    ToolTip1.SetToolTip(alabels(i), mydata)
        'Next
        '---------------------


        'Dim l_a As New Label()

        'l_a.Text = ""
        'l_a.Location = New Point(286, 768)
        'l_a.BackColor = Color.DimGray
        'l_a.Size = New System.Drawing.Size(10, 22)
        'l_a.Anchor = AnchorStyles.Top

        'Me.Controls.Add(l_a)



        'Dim block As New Label
        'block.Size = New System.Drawing.Size(267, 74)
        'block.BackColor = Color.Teal
        'block.Text = "ADA-TEST"
        'block.ForeColor = Color.WhiteSmoke
        'block.Font = New Font("Segoe UI", 14, FontStyle.Regular)
        'block.Anchor = AnchorStyles.Top
        'block.Location = New Point(161, 790)
        'block.TextAlign = ContentAlignment.MiddleCenter

        'Me.Controls.Add(block)
        'block.ContextMenuStrip = ContextMenuStrip1



        'AddHandler block.MouseEnter, AddressOf block_color
        'AddHandler block.MouseLeave, AddressOf block_normal

    End Sub

    Private Sub block_color(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.BackColor = Color.SlateGray
    End Sub

    Private Sub block_normal(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.BackColor = Color.Teal
    End Sub

    Private Sub Package_BOM_tree_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Panel1.Width = 1
            Me.Close()
        End If
    End Sub

    'Private Sub AddObjectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddObjectToolStripMenuItem.Click
    '    '  Panel1.Size = New System.Drawing.Size(662, 826)
    '    'Panel1.Width = 662
    '    'Me.Refresh()


    '    'Label1.Text = "Selected Object: " & ContextMenuStrip1.SourceControl.Text
    'End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        ' Panel1.Size = New System.Drawing.Size(10, 826)
        Panel1.Width = 1
        Me.Refresh()
        MessageBox.Show("BOM added")
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Panel1.Width = 1
        Me.Refresh()
    End Sub

    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = Color.Red
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.Maroon
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

            Dim plabels(list.Count - 1) As Label
            Dim tronco(list.Count - 1) As Label

            Dim alabels(list_as.Count - 1) As Label
            Dim tronco2(list_as.Count - 1) As Label

            Dim y As Integer : y = ini_z '384
            Dim z As Integer : z = ini_z + 22 '406

            Dim y2 As Integer : y2 = ini_z '384
            Dim z2 As Integer : z2 = ini_z + 22 '406

            '--- add branches and panel blocks---
            For i = 0 To list.Count - 1

                'branches
                tronco(i) = New Label()
                tronco(i).Text = ""
                tronco(i).Location = New Point(ini_y, y)
                tronco(i).BackColor = Color.DimGray
                tronco(i).Size = New System.Drawing.Size(10, 22)
                tronco(i).Anchor = AnchorStyles.Top

                Me.Controls.Add(tronco(i))
                y = y + 96

                'panel blocks

                plabels(i) = New Label()
                plabels(i).Size = New System.Drawing.Size(267, 74)   '267,  74
                plabels(i).BackColor = Color.Teal
                plabels(i).Text = list.Item(i).ToString
                plabels(i).ForeColor = Color.WhiteSmoke
                plabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
                plabels(i).Anchor = AnchorStyles.Top
                plabels(i).Location = New Point(ini_y - CType((w_i) / 2, Integer), z)
                plabels(i).TextAlign = ContentAlignment.MiddleCenter


                Me.Controls.Add(plabels(i))
                z = z + 96

                plabels(i).ContextMenuStrip = ContextMenuStrip1

                AddHandler plabels(i).MouseEnter, AddressOf block_color
                AddHandler plabels(i).MouseLeave, AddressOf block_normal

                ToolTip1.SetToolTip(plabels(i), BOM_info(plabels(i).Text, job))
            Next

            '------------------


            '----- add assem -----
            For i = 0 To list_as.Count - 1

                'branches
                tronco2(i) = New Label()
                tronco2(i).Text = ""
                tronco2(i).Location = New Point(x_ia + CType((w_ia) / 2, Integer), y2)
                tronco2(i).BackColor = Color.DimGray
                tronco2(i).Size = New System.Drawing.Size(10, 22)
                tronco2(i).Anchor = AnchorStyles.Top

                Me.Controls.Add(tronco2(i))
                y2 = y2 + 96

                'panel blocks

                alabels(i) = New Label()
                alabels(i).Size = New System.Drawing.Size(267, 74)   '267, 74
                alabels(i).BackColor = Color.Teal
                alabels(i).Text = list_as.Item(i).ToString
                alabels(i).ForeColor = Color.WhiteSmoke
                alabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
                alabels(i).Anchor = AnchorStyles.Top
                alabels(i).Location = New Point(x_ia, z2)
                alabels(i).TextAlign = ContentAlignment.MiddleCenter

                Me.Controls.Add(alabels(i))
                z2 = z2 + 96
                alabels(i).ContextMenuStrip = ContextMenuStrip1

                AddHandler alabels(i).MouseEnter, AddressOf block_color
                AddHandler alabels(i).MouseLeave, AddressOf block_normal

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
        Dim p_desc As String : p_desc = "Not specified"
        Dim qty_p As String : qty_p = 1

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
                cmd_j2.CommandText = "SELECT created_by, Date_Created, release_date, mr_name, Panel_desc, Panel_qty from Material_Request.mr where Panel_name = @Panel_name and job = @job"
                cmd_j2.Connection = Login.Connection

                Dim readerj2 As MySqlDataReader
                readerj2 = cmd_j2.ExecuteReader

                If readerj2.HasRows Then
                    While readerj2.Read
                        created_by = readerj2(0).ToString
                        date_created = readerj2(1).ToString
                        date_released = readerj2(2).ToString
                        mr_name = readerj2(3).ToString
                        p_desc = If(IsDBNull(readerj2(4)) = True, "Not specified", readerj2(4))
                        qty_p = If(IsDBNull(readerj2(5)) = True, "1", readerj2(5))
                    End While
                End If

                readerj2.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        BOM_info = mr_name & vbCrLf & vbCrLf & "Created by: " & created_by & vbCrLf & vbCrLf & " Date Created: " & date_created _
        & vbCrLf & vbCrLf & " Date Released: " & date_released & vbCrLf & vbCrLf & "Description: " & p_desc & vbCrLf & vbCrLf & "Qty: " & qty_p


    End Function

    Private Sub CreateRevisionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateRevisionToolStripMenuItem.Click
        If My_Material_r.Visible = True Then
            '------- get mr name ------
            Dim name As String : name = ContextMenuStrip1.SourceControl.Text
            Dim mr_name As String : mr_name = ""

            If String.Equals(name, "Field BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "Field")
                cmd5.Parameters.AddWithValue("@job", My_Material_r.open_job)

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
                cmd5.Parameters.AddWithValue("@job", My_Material_r.open_job)

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
                cmd5.Parameters.AddWithValue("@job", My_Material_r.open_job)

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
                cmd5.Parameters.AddWithValue("@job", My_Material_r.open_job)
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

                '-get latest rev
                mr_name = Procurement_Overview.get_last_revision(mr_name)


            End If

            '---- open bom in material request

            My_Material_r.PR_grid.Rows.Clear()


            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", mr_name)
                cmd4.CommandText = "SELECT Part_No, Description, ADA_Number, Manufacturer, Vendor, Price, Qty, subtotal, mfg_type, Assembly_name ,Notes, Part_status, Part_type, full_panel, need_by_date, qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    Dim i As Integer : i = 0
                    While reader4.Read
                        My_Material_r.PR_grid.Rows.Add(New String() {})
                        My_Material_r.PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(4).Value = reader4(4).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(5).Value = reader4(5).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(6).Value = reader4(6).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(7).Value = reader4(7).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(8).Value = reader4(8).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(9).Value = reader4(9).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(10).Value = reader4(10).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(11).Value = reader4(11).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(12).Value = reader4(12).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(13).Value = reader4(13).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(14).Value = reader4(14).ToString
                        My_Material_r.PR_grid.Rows(i).Cells(15).Value = reader4(15).ToString

                        i = i + 1
                    End While

                End If

                reader4.Close()
                My_Material_r.Text = mr_name
                My_Material_r.single_grid.Rows.Clear()

                My_Material_r.TabControl1.SelectedTab = My_Material_r.TabPage1


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            For i = 0 To My_Material_r.PR_grid.Rows.Count - 1
                My_Material_r.PR_grid.Rows(i).ReadOnly = True
            Next

            My_Material_r.isitreleased = True
            My_Material_r.TabControl1.TabPages.Remove(My_Material_r.TabPage2)
            Inventory_manage.part_sel = ""
            My_Material_r.rev_mode = False
            Me.Visible = False
            '-----------------------------------

        Else
            MessageBox.Show("Revision not allowed in Unreleased BOM")
        End If
    End Sub
End Class