Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class Manufacturing_Report

    Public previous_f As Integer  '
    Public counter_i As Integer
    Public my_assemblies As List(Of String)




    Private Sub Manufacturing_Report_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        Call EnableDoubleBuffered(MO_grid)
        Call EnableDoubleBuffered(open_grid)
        Call EnableDoubleBuffered(total_grid)
        Call EnableDoubleBuffered(asm_grid)


        open_grid.Columns(5).DisplayIndex = 4
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        TabControl1.TabPages.Remove(TabPage5)
        TabControl1.TabPages.Remove(TabPage6)

        Try
            job_box.Items.Clear()

            my_assemblies = New List(Of String)()

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
            '-------------------------------------------------

            Dim check_cmd21 As New MySqlCommand
            check_cmd21.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
            check_cmd21.Connection = Login.Connection
            check_cmd21.ExecuteNonQuery()

            Dim reader21 As MySqlDataReader
            reader21 = check_cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    job_box.Items.Add(reader21(0))
                End While
            End If

            reader21.Close()

            '---------------------------------
            Call load_all()
            Call load_inv_as()
            Call remove_no_order()
            Call apply_filter()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.BackColor = Color.CadetBlue
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.BackColor = Color.DarkSlateGray
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Me.Visible = False
        Reports_menu.Visible = True
    End Sub

    Sub load_all()
        '-- load all jobs
        Try

            open_grid.Rows.Clear()

            Dim job_list = New List(Of String)()
            counter_i = 0


            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr where released = 'Y' order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    If job_list.Contains(reader2(0).ToString) = False Then
                        job_list.Add(reader2(0).ToString)
                    End If
                End While
            End If

            reader2.Close()
            '-------------------------------
            For i = 0 To job_list.Count - 1

                Dim bom_list = New List(Of String)()

                '------- get BOM ------
                '------ get all BOM of the job 
                Dim n_bom As Double : n_bom = 0
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@job", job_list.Item(i))
                check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        n_bom = reader(0)
                    End While
                End If

                reader.Close()

                For j = 1 To n_bom

                    Dim check_cmd1 As New MySqlCommand
                    check_cmd1.Parameters.Clear()
                    check_cmd1.Parameters.AddWithValue("@job", job_list.Item(i))
                    check_cmd1.Parameters.AddWithValue("@id_bom", j)
                    check_cmd1.CommandText = "select p1.job, management.projects.Job_description, p1.mr_name, p1.need_date , p1.released_by from Material_Request.mr as p1 INNER JOIN management.projects ON p1.job = management.projects.Job_number where p1.job = @job and p1.id_bom = @id_bom and (p1.BOM_type = 'old_BOM' or p1.BOM_Type = 'MB' or p1.BOM_Type = 'ASM') order by p1.release_date desc limit 1"

                    check_cmd1.Connection = Login.Connection
                    check_cmd1.ExecuteNonQuery()

                    Dim reader1 As MySqlDataReader
                    reader1 = check_cmd1.ExecuteReader

                    If reader1.HasRows Then

                        While reader1.Read
                            open_grid.Rows.Add(New String() {})
                            open_grid.Rows(counter_i).Cells(0).Value = reader1(0).ToString
                            open_grid.Rows(counter_i).Cells(1).Value = reader1(1).ToString
                            open_grid.Rows(counter_i).Cells(2).Value = reader1(2).ToString
                            open_grid.Rows(counter_i).Cells(3).Value = Convert.ToDateTime(reader1(3))  'need date delete to string
                            open_grid.Rows(counter_i).Cells(4).Value = reader1(4).ToString
                            open_grid.Rows(counter_i).Cells(5).Value = "0"

                            counter_i = counter_i + 1
                        End While
                    End If

                    reader1.Close()
                Next
            Next

            counter_i = 0

            Call fill_percentage()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label6.Visible = True
        Application.DoEvents()

        CheckBox1.Checked = False
        CheckBox3.Checked = False

        Call load_all()
        Label6.Visible = False
    End Sub

    Private Sub job_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles job_box.SelectedValueChanged
        '-- select job --
        open_grid.Rows.Clear()

        Try

            '------ get all BOM of the job Note: unelegant way of do this. Try to find a better Query
            Dim n_bom As Double : n_bom = 0
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", job_box.SelectedItem.ToString)
            check_cmd.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    n_bom = reader(0)
                End While
            End If

            reader.Close()

            For i = 1 To n_bom

                Dim check_cmd1 As New MySqlCommand
                check_cmd1.Parameters.Clear()
                check_cmd1.Parameters.AddWithValue("@job", job_box.SelectedItem.ToString)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select p1.job, management.projects.Job_description, p1.mr_name, p1.need_date , p1.released_by from Material_Request.mr as p1 INNER JOIN management.projects ON p1.job = management.projects.Job_number where p1.job = @job and p1.id_bom = @id_bom and (p1.BOM_type = 'old_BOM' or p1.BOM_Type = 'MB' or p1.BOM_Type = 'ASM') order by p1.release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    ' Dim j As Integer : j = 0


                    While reader1.Read
                        open_grid.Rows.Add(New String() {})
                        open_grid.Rows(i - 1).Cells(0).Value = reader1(0).ToString
                        open_grid.Rows(i - 1).Cells(1).Value = reader1(1).ToString
                        open_grid.Rows(i - 1).Cells(2).Value = reader1(2).ToString
                        open_grid.Rows(i - 1).Cells(3).Value = Convert.ToDateTime(reader1(3))  'need date delete to string
                        open_grid.Rows(i - 1).Cells(4).Value = reader1(4).ToString
                        open_grid.Rows(i - 1).Cells(5).Value = "0"


                    End While
                End If

                reader1.Close()
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter
        Label3.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave
        Label3.ForeColor = Color.DarkTurquoise
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.ForeColor = Color.DarkTurquoise
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles Label5.MouseEnter
        Label5.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles Label5.MouseLeave
        Label5.ForeColor = Color.DarkTurquoise
    End Sub



    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        '--go back to projects tab
        previous_f = 1

        TabControl1.TabPages.Insert(1, TabPage1)
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.SelectedTab = TabPage1

    End Sub

    Private Sub PR_grid_DoubleClick(sender As Object, e As EventArgs) Handles PR_grid.DoubleClick
        '-- open panel bom or assembly ---

        If PR_grid.Rows.Count > 0 Then

            Dim index_k = PR_grid.CurrentCell.RowIndex
            Dim panel As String = ""


            If PR_grid.Rows.Count > 0 Then
                panel = If(String.IsNullOrEmpty(PR_grid.Rows(index_k).Cells(0).Value) = True, "", PR_grid.Rows(index_k).Cells(0).Value)
            End If

            If String.Equals(panel, "") = False Then

                If my_assemblies.Contains(panel) = True Then  'if its an assembly

                    assem_name.Text = panel

                    Call open_assem(panel, PR_grid.Rows(index_k).Cells(2).Value)  'open assembly


                    previous_f = 4 'for assembly go back to build request
                    TabControl1.TabPages.Insert(1, TabPage5)
                    TabControl1.TabPages.Remove(TabPage3)
                    TabControl1.SelectedTab = TabPage5
                    job_n2.Text = job_l.Text
                    Label7.Text = "(" & PR_grid.Rows(index_k).Cells(2).Value & ")"



                Else
                    '---- fil panel data -----
                    Dim mr_name As String : mr_name = ""
                    Dim temp_panel_Qty As Double : temp_panel_Qty = 1

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@job", job_l.Text)
                    cmd5.Parameters.AddWithValue("@Panel_name", panel)
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

                    '--- get qty panel ---
                    Dim cmd419 As New MySqlCommand
                    cmd419.Parameters.Clear()
                    cmd419.Parameters.AddWithValue("@mr_name", mr_name)
                    cmd419.CommandText = "SELECT Panel_qty from Material_Request.mr where mr_name = @mr_name"
                    cmd419.Connection = Login.Connection
                    Dim reader419 As MySqlDataReader
                    reader419 = cmd419.ExecuteReader

                    If reader419.HasRows Then
                        While reader419.Read

                            n_panels.Text = "(" & If(reader419(0) Is DBNull.Value, 1, reader419(0)) & ")"
                            temp_panel_Qty = If(reader419(0) Is DBNull.Value, 1, reader419(0))

                        End While
                    End If

                    reader419.Close()
                    '----------------------------------


                    If String.Equals(mr_name, "not found") = False Then
                        Call Open_panel(mr_name, job_l.Text, temp_panel_Qty)
                    End If

                    '-------------------------

                    Call color_needed(total_grid)
                    job_p.Text = job_l.Text
                    panel_n.Text = panel
                    previous_f = 3 'go back to build request
                    TabControl1.TabPages.Insert(1, TabPage4)
                    TabControl1.TabPages.Remove(TabPage3)
                    TabControl1.SelectedTab = TabPage4

                End If

            End If

        End If


    End Sub

    Sub color_needed(mygrid As DataGridView)

        For Each row As DataGridViewRow In mygrid.Rows
            If row.IsNewRow Then Continue For

            If (IsNumeric(row.Cells(2).Value) = True And IsNumeric(row.Cells(6).Value)) Then
                row.Cells(7).Value = CType(row.Cells(2).Value, Double) - CType(row.Cells(6).Value, Double)
            Else
                row.Cells(7).Value = 0
            End If

            If CType(row.Cells(7).Value, Double) <> 0 Then
                row.Cells(7).Style.BackColor = Color.Firebrick
            Else
                row.Cells(7).Style.BackColor = Color.Gray
            End If
        Next
    End Sub


    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick
        '--- click a job
        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex
            Dim job As String = ""


            If open_grid.Rows.Count > 0 Then
                job = If(String.IsNullOrEmpty(open_grid.Rows(index_k).Cells(0).Value) = True, "", open_grid.Rows(index_k).Cells(0).Value)
            End If

            If String.Equals(job, "") = False Then

                Try
                    '---- fill br ---
                    '--get latest revision --
                    PR_grid.Rows.Clear()
                    job_l.Text = job

                    Dim n_r As String : n_r = 0
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.AddWithValue("@job", job)
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

                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@job", job)
                    cmd4.Parameters.AddWithValue("@n_r", n_r)
                    cmd4.CommandText = "SELECT panel, panel_desc, panel_qty, need_date, br_name, ready_t from Build_request.build_r where job = @job and n_r = @n_r"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        Dim i As Integer : i = 0
                        While reader4.Read
                            PR_grid.Rows.Add(New String() {})
                            PR_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                            PR_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                            PR_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                            PR_grid.Rows(i).Cells(3).Value = reader4(3).ToString
                            PR_grid.Rows(i).Cells(4).Value = If(String.Equals(reader4(5).ToString, "Y") = True, True, False)

                            i = i + 1
                        End While

                    End If

                    reader4.Close()
                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


                TabControl1.TabPages.Insert(1, TabPage3)
                TabControl1.TabPages.Remove(TabPage1)
                TabControl1.SelectedTab = TabPage3

                previous_f = 1

            End If

        End If

    End Sub

    Sub fill_percentage()

        For i = 0 To open_grid.Rows.Count - 1

            '---------- calculate percentage ---------
            Dim complete_mr As Double : complete_mr = 0
            Dim total_qty As Double : total_qty = 0
            Dim fullf As Double : fullf = 0


            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@mr_name", open_grid.Rows(i).Cells(2).Value)
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
            cmdx.Parameters.AddWithValue("@mr_name", open_grid.Rows(i).Cells(2).Value)
            cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null;"
            cmdx.Connection = Login.Connection
            Dim readerx As MySqlDataReader
            readerx = cmdx.ExecuteReader

            If readerx.HasRows Then
                While readerx.Read
                    If IsDBNull(readerx(0)) Then
                        total_qty = 0
                    Else
                        total_qty = CType(readerx(0), Double)
                    End If
                End While
            End If

            readerx.Close()

            Dim complete_noround As Double : complete_noround = 0

            If total_qty > 0 Then
                complete_mr = Math.Floor((fullf / total_qty) * 100)
                complete_noround = (fullf / total_qty) * 100
            End If

            open_grid.Rows(i).Cells(5).Value = complete_mr

        Next

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        '--go back tp build request
        previous_f = 2

        TabControl1.TabPages.Insert(1, TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        TabControl1.SelectedTab = TabPage3
    End Sub

    Sub load_inv_as()
        '------- load inventory table tab ----------
        Try
            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT Legacy_ADA_Number, Description from devices"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    MO_grid.Rows.Add(New String() {reader4(0), reader4(1), 0, 0.0})
                End While
            End If

            reader4.Close()

            For i = 0 To MO_grid.Rows.Count - 1
                MO_grid.Rows(i).Cells(2).Value = myQuote.Cost_of_Assem(MO_grid.Rows(i).Cells(0).Value)

                Dim cmd44 As New MySqlCommand
                cmd44.Parameters.Clear()
                cmd44.Parameters.AddWithValue("@part", MO_grid.Rows(i).Cells(0).Value)
                cmd44.CommandText = "select Qty_needed from inventory.Material_orders where Part_No = @part "
                cmd44.Connection = Login.Connection
                Dim reader44 As MySqlDataReader
                reader44 = cmd44.ExecuteReader

                If reader44.HasRows Then
                    While reader44.Read
                        MO_grid.Rows(i).Cells(3).Value = CType(reader44(0), Double)
                    End While
                End If

                reader44.Close()

                '---------------inv ----------
                Dim cmd46 As New MySqlCommand
                cmd46.Parameters.Clear()
                cmd46.Parameters.AddWithValue("@part", MO_grid.Rows(i).Cells(0).Value)
                cmd46.CommandText = "select current_qty, min_qty, max_qty, Qty_in_order, es_date_of_arrival from inventory.inventory_qty where part_name = @part "
                cmd46.Connection = Login.Connection
                Dim reader46 As MySqlDataReader
                reader46 = cmd46.ExecuteReader

                If reader46.HasRows Then
                    While reader46.Read
                        MO_grid.Rows(i).Cells(4).Value = If(IsDBNull(reader46(0)) = False, CType(reader46(0), Double), 0)
                        MO_grid.Rows(i).Cells(5).Value = If(IsDBNull(reader46(1)) = False, CType(reader46(1), Double), 0)
                        MO_grid.Rows(i).Cells(6).Value = If(IsDBNull(reader46(2)) = False, CType(reader46(2), Double), 0)
                        MO_grid.Rows(i).Cells(7).Value = If(IsDBNull(reader46(3)) = False, CType(reader46(3), Double), 0)
                        MO_grid.Rows(i).Cells(8).Value = If(IsDBNull(reader46(4)) = False, reader46(4).ToString, "")
                    End While
                End If

                reader46.Close()


            Next

            '--------- need to order from material order table and inventoey table ----




        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged


        If CheckBox2.Checked = False Then
            MO_grid.Rows.Clear()
            Call load_inv_as()
        Else
            Call remove_no_order()
        End If
    End Sub

    Sub remove_no_order()
        For i = MO_grid.Rows.Count - 1 To 0 Step -1
            If MO_grid.Rows(i).Cells(3).Value = 0 Then
                MO_grid.Rows.Remove(MO_grid.Rows(i))
            End If
        Next
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Call apply_filter()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Call apply_filter()
    End Sub

    Sub apply_filter()
        '--- apply filters ---


        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Project", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Mr", GetType(String))
        dimen_table.Columns.Add("need_date", GetType(DateTime))
        dimen_table.Columns.Add("release_by", GetType(String))
        dimen_table.Columns.Add("fullfille", GetType(Double))

        '--result table --
        Dim dimen_table2 = New DataTable
        dimen_table2.Columns.Add("Project", GetType(String))
        dimen_table2.Columns.Add("Description", GetType(String))
        dimen_table2.Columns.Add("Mr", GetType(String))
        dimen_table2.Columns.Add("need_date", GetType(DateTime))
        dimen_table2.Columns.Add("release_by", GetType(String))
        dimen_table2.Columns.Add("fullfille", GetType(Double))

        For i = 0 To open_grid.Rows.Count - 1
            dimen_table.Rows.Add(open_grid.Rows(i).Cells(0).Value, open_grid.Rows(i).Cells(1).Value, open_grid.Rows(i).Cells(2).Value, open_grid.Rows(i).Cells(3).Value, open_grid.Rows(i).Cells(4).Value, open_grid.Rows(i).Cells(5).Value)
        Next


        '------ hide no dates --
        If CheckBox1.Checked = True Then
            For i = dimen_table.Rows.Count - 1 To 0 Step -1
                If DateTime.Compare(dimen_table.Rows(i).Item(3), Convert.ToDateTime("1/1/1970 10:54:07")) = 0 Then
                    dimen_table.Rows.Remove(dimen_table.Rows(i))
                End If
            Next
        End If

        '---- show only fullfilled ----

        If CheckBox3.Checked = True Then
            For i = dimen_table.Rows.Count - 1 To 0 Step -1
                If dimen_table.Rows(i).Item(5) >= 100 Then
                    dimen_table.Rows.Remove(dimen_table.Rows(i))
                End If
            Next
        End If


        For i = 0 To dimen_table.Rows.Count - 1
            dimen_table2.Rows.Add(dimen_table.Rows(i).Item(0), dimen_table.Rows(i).Item(1), dimen_table.Rows(i).Item(2), dimen_table.Rows(i).Item(3), dimen_table.Rows(i).Item(4), dimen_table.Rows(i).Item(5))
        Next

        open_grid.Rows.Clear()

        For i = 0 To dimen_table2.Rows.Count - 1
            open_grid.Rows.Add(New String() {})
            open_grid.Rows(i).Cells(0).Value = dimen_table2.Rows(i).Item(0)
            open_grid.Rows(i).Cells(1).Value = dimen_table2.Rows(i).Item(1)
            open_grid.Rows(i).Cells(2).Value = dimen_table2.Rows(i).Item(2)
            open_grid.Rows(i).Cells(3).Value = Convert.ToDateTime(dimen_table2.Rows(i).Item(3))
            open_grid.Rows(i).Cells(4).Value = dimen_table2.Rows(i).Item(4)
            open_grid.Rows(i).Cells(5).Value = dimen_table2.Rows(i).Item(5)
        Next

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '=== export table =======
        If TabControl1.SelectedTab Is TabPage1 Then

            Call export_table(open_grid)

        ElseIf TabControl1.SelectedTab Is TabPage2 Then

            Call export_table(MO_grid)

        ElseIf TabControl1.SelectedTab Is TabPage3 Then

            Call export_table(PR_grid)

        ElseIf TabControl1.SelectedTab Is TabPage4 Then

            Call export_table(total_grid)

        ElseIf TabControl1.SelectedTab Is TabPage5 Then

            Call export_table(asm_grid)

        ElseIf TabControl1.SelectedTab Is TabPage6 Then

            Call export_table(asm2_grid)

        End If
    End Sub

    Sub export_table(mygrid As DataGridView)

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            Try
                Dim xlWorkBook As Excel.Workbook
                Dim xlWorkSheet As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")
                xlWorkSheet.Range("A:B").ColumnWidth = 40
                xlWorkSheet.Range("C:C").ColumnWidth = 30
                xlWorkSheet.Range("D:E").ColumnWidth = 20
                xlWorkSheet.Range("A:E").HorizontalAlignment = Excel.Constants.xlCenter


                'copy data to excel
                For i As Integer = 0 To mygrid.ColumnCount - 1
                    xlWorkSheet.Cells(1, i + 1) = mygrid.Columns(i).HeaderText
                    For j As Integer = 0 To mygrid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = mygrid.Rows(j).Cells(i).Value
                    Next j
                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Table exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Sub Open_panel(name As String, job As String, qty_p As Double)

        'this sub open  the data in the inventory request datagridview
        total_grid.Rows.Clear()

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("qty_needed", GetType(String))


        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", name)
            cmd4.CommandText = "SELECT Part_No, Description, Qty, Manufacturer, Vendor, Price, qty_fullfilled, qty_needed from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString * qty_p, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString)
                End While
            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                total_grid.Rows.Add(New String() {})
                total_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                total_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                total_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                total_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                total_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                total_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                total_grid.Rows(i).Cells(6).Value = If(IsNumeric(dimen_table.Rows(i).Item(6).ToString) = True, dimen_table.Rows(i).Item(6).ToString, 0)
                total_grid.Rows(i).Cells(7).Value = If(IsNumeric(dimen_table.Rows(i).Item(7).ToString) = True, dimen_table.Rows(i).Item(7).ToString, 0)
            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To total_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        total_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                        total_grid.Rows(i).Cells(9).Value = reader5(1).ToString
                        total_grid.Rows(i).Cells(10).Value = reader5(2).ToString
                        total_grid.Rows(i).Cells(11).Value = reader5(3).ToString
                    End While
                End If

                reader5.Close()

                '---- get project specific qty-----

                Dim cmd51 As New MySqlCommand
                cmd51.Parameters.Clear()
                cmd51.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(0).Value)
                cmd51.Parameters.AddWithValue("@job", job)
                cmd51.CommandText = "SELECT qty_needed, PO, es_date_of_arrival from Tracking_Reports.my_tracking_reports where Part_No = @Part_No and job = @job"
                cmd51.Connection = Login.Connection
                Dim reader51 As MySqlDataReader
                reader51 = cmd51.ExecuteReader

                If reader51.HasRows Then
                    While reader51.Read

                        If String.IsNullOrEmpty(reader51(1).ToString) = False Or String.IsNullOrEmpty(reader51(2).ToString) = False Then

                            total_grid.Rows(i).Cells(9).Value = reader51(0).ToString

                        End If

                    End While
                End If

                reader51.Close()



            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        '-- go back from assemblie bom
        If previous_f = 4 Then

            TabControl1.TabPages.Insert(1, TabPage3)
            TabControl1.TabPages.Remove(TabPage5)
            TabControl1.SelectedTab = TabPage3

            previous_f = 2

        ElseIf previous_f = 7 Then

            TabControl1.TabPages.Insert(1, TabPage4)
            TabControl1.TabPages.Remove(TabPage5)
            TabControl1.SelectedTab = TabPage4

            previous_f = 2
        End If




    End Sub

    Sub open_assem(assem As String, qty As Double)
        '-- display an assembly

        asm_grid.Rows.Clear()

        '=================================== If the device exist in the Devices table then ===================================
        Dim atronix_n As String : atronix_n = ""
        '--------------------------- Get ADV number --------------------
        Try


            Dim cmd_a As New MySqlCommand
            cmd_a.Parameters.AddWithValue("@Legacy_ADA_Number", assem)
            cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @Legacy_ADA_Number"
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
            cmd_pd.Parameters.AddWithValue("@adv", atronix_N)
            cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, adv.Qty,  p1.Manufacturer,  p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = @adv"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    asm_grid.Rows.Add()  'add a new row
                    asm_grid.Rows(i).Cells(0).Value = readerv(0).ToString
                    asm_grid.Rows(i).Cells(1).Value = readerv(1).ToString
                    asm_grid.Rows(i).Cells(2).Value = readerv(2).ToString * qty
                    asm_grid.Rows(i).Cells(3).Value = readerv(3).ToString
                    asm_grid.Rows(i).Cells(4).Value = readerv(4).ToString
                    i = i + 1
                End While
            End If

            readerv.Close()


            For j = 0 To asm_grid.Rows.Count - 1
                asm_grid.Rows(j).Cells(5).Value = Form1.Get_Latest_Cost(Login.Connection, asm_grid.Rows(j).Cells(0).Value, asm_grid.Rows(j).Cells(4).Value) * asm_grid.Rows(j).Cells(2).Value
            Next

            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To asm_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        asm_grid.Rows(i).Cells(6).Value = reader5(0).ToString
                        asm_grid.Rows(i).Cells(7).Value = reader5(1).ToString
                        asm_grid.Rows(i).Cells(8).Value = reader5(2).ToString
                        asm_grid.Rows(i).Cells(9).Value = reader5(3).ToString
                    End While
                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub open_assem2(assem As String, qty As Double)
        '-- display an assembly

        asm2_grid.Rows.Clear()

        '=================================== If the device exist in the Devices table then ===================================
        Dim atronix_n As String : atronix_n = ""
        '--------------------------- Get ADV number --------------------
        Try


            Dim cmd_a As New MySqlCommand
            cmd_a.Parameters.AddWithValue("@Legacy_ADA_Number", assem)
            cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @Legacy_ADA_Number"
            cmd_a.Connection = Login.Connection

            Dim reader_k As MySqlDataReader
            reader_k = cmd_a.ExecuteReader


            If reader_k.HasRows Then
                While reader_k.Read
                    atronix_n = reader_k(0)
                End While
            End If

            reader_k.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '-------------------------------------------------------------------------

        Try
            Dim cmd_pd As New MySqlCommand
            cmd_pd.Parameters.AddWithValue("@adv", atronix_n)
            cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, adv.Qty,  p1.Manufacturer,  p1.Primary_Vendor from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = @adv"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    asm2_grid.Rows.Add()  'add a new row
                    asm2_grid.Rows(i).Cells(0).Value = readerv(0).ToString
                    asm2_grid.Rows(i).Cells(1).Value = readerv(1).ToString
                    asm2_grid.Rows(i).Cells(2).Value = readerv(2).ToString * qty
                    asm2_grid.Rows(i).Cells(3).Value = readerv(3).ToString
                    asm2_grid.Rows(i).Cells(4).Value = readerv(4).ToString
                    i = i + 1
                End While
            End If

            readerv.Close()


            For j = 0 To asm2_grid.Rows.Count - 1
                asm2_grid.Rows(j).Cells(5).Value = Form1.Get_Latest_Cost(Login.Connection, asm2_grid.Rows(j).Cells(0).Value, asm2_grid.Rows(j).Cells(4).Value) * asm2_grid.Rows(j).Cells(2).Value
            Next

            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To asm2_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", asm2_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        asm2_grid.Rows(i).Cells(6).Value = reader5(0).ToString
                        asm2_grid.Rows(i).Cells(7).Value = reader5(1).ToString
                        asm2_grid.Rows(i).Cells(8).Value = reader5(2).ToString
                        asm2_grid.Rows(i).Cells(9).Value = reader5(3).ToString
                    End While
                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub MO_grid_DoubleClick(sender As Object, e As EventArgs) Handles MO_grid.DoubleClick

        Dim index_k = MO_grid.CurrentCell.RowIndex
        Dim assem As String = ""


        If MO_grid.Rows.Count > 0 Then
            assem = If(String.IsNullOrEmpty(MO_grid.Rows(index_k).Cells(0).Value) = True, "", MO_grid.Rows(index_k).Cells(0).Value)


            Call open_assem2(MO_grid.Rows(index_k).Cells(0).Value, 1)  'open assembly

            assem_name.Text = assem
            TabControl1.TabPages.Insert(2, TabPage6)
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.SelectedTab = TabPage6

        End If

    End Sub

    Private Sub total_grid_DoubleClick(sender As Object, e As EventArgs) Handles total_grid.DoubleClick

        Dim index_k = total_grid.CurrentCell.RowIndex
        Dim assem As String = ""


        '--- may open an assembly from here
        If my_assemblies.Contains(total_grid.Rows(index_k).Cells(0).Value) = True Then  'if its an assembly

            assem_name.Text = total_grid.Rows(index_k).Cells(0).Value

            Call open_assem(total_grid.Rows(index_k).Cells(0).Value, total_grid.Rows(index_k).Cells(2).Value)  'open assembly


            previous_f = 7 'for assembly go back to build request
            TabControl1.TabPages.Insert(1, TabPage5)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.SelectedTab = TabPage5
            job_n2.Text = job_l.Text
            Label7.Text = "(" & total_grid.Rows(index_k).Cells(2).Value & ")"

        End If
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click

        TabControl1.TabPages.Insert(2, TabPage2)
        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.SelectedTab = TabPage2
    End Sub
End Class