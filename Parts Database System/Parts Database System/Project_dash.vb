Imports System.Windows.Forms.DataVisualization.Charting
Imports MySql.Data.MySqlClient

Public Class Project_dash

    Public missing_p As Boolean

    Private Sub Project_dash_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''--------////////// Load MR PIE CHART //////////-----------------
        'Dim yValues As Double() = {70, 30} ' Getting values from Textboxes 
        'Dim xValues As String() = {"Parts Fullfilled", "Remainder"} ' Headings
        'Dim seriesName As String = Nothing

        'Chart1.Series.Clear()
        'Chart1.Titles.Clear()

        'seriesName = "MR Status"
        'Chart1.Series.Add(seriesName)


        '' Bind X and Y values
        'Chart1.Series(seriesName).Points.DataBindXY(xValues, yValues)
        'Chart1.Series(seriesName).Points(0).Color = Color.FromArgb(0, 64, 64)
        'Chart1.Series(seriesName).Points(1).Color = Color.FromArgb(52, 52, 52)

        'Chart1.Series(seriesName).ChartType = SeriesChartType.Pie

        'Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = True

        'Dim T As Title = Chart1.Titles.Add("Material Request Fullfillment Status")
        'With T
        '    .ForeColor = Color.WhiteSmoke
        '    .BackColor = Color.FromArgb(38, 50, 76)
        '    .Font = New System.Drawing.Font("Segoe UI", 13.0F)

        'End With

        'Chart1.Series(seriesName).Points(0).LabelForeColor = Color.WhiteSmoke
        'Chart1.Series(seriesName).Points(1).LabelForeColor = Color.WhiteSmoke
        'Chart1.Series(seriesName).Font = New Font("Segoe UI", 15.0F)
        'Chart1.Series(seriesName).IsValueShownAsLabel = True

        'Chart1.Legends(0).Enabled = True
        'Chart1.Legends(0).Font = New Font("Segoe UI", 12.0F)

        ''--------//////////// Load MPL PIE CHART ///////////-----------------
        'Dim yValues2 As Double() = {"80", "20"} ' Getting values from Textboxes 
        'Dim xValues2 As String() = {"Parts Shipped", "Remainder"} ' Headings
        'Dim seriesName2 As String = Nothing

        'Chart2.Series.Clear()
        'Chart2.Titles.Clear()

        'seriesName2 = "MPL Status"
        'Chart2.Series.Add(seriesName2)


        '' Bind X and Y values
        'Chart2.Series(seriesName2).Points.DataBindXY(xValues2, yValues2)
        'Chart2.Series(seriesName2).Points(0).Color = Color.FromArgb(0, 64, 64)
        'Chart2.Series(seriesName2).Points(1).Color = Color.FromArgb(52, 52, 52)

        'Chart2.Series(seriesName2).ChartType = SeriesChartType.Pie

        'Chart2.ChartAreas("ChartArea1").Area3DStyle.Enable3D = True
        'Chart2.ChartAreas("ChartArea1").Area3DStyle.Inclination = 30

        'Dim T2 As Title = Chart2.Titles.Add("Master Packing List Status")
        'With T2
        '    .ForeColor = Color.WhiteSmoke
        '    .BackColor = Color.FromArgb(38, 50, 76)
        '    .Font = New System.Drawing.Font("Segoe UI", 13.0F)

        'End With

        'Chart2.Series(seriesName2).Points(0).LabelForeColor = Color.WhiteSmoke
        'Chart2.Series(seriesName2).Points(1).LabelForeColor = Color.WhiteSmoke

        'Chart2.Series(seriesName2).Font = New Font("Segoe UI", 15.0F)
        'Chart2.Series(seriesName2).IsValueShownAsLabel = True

        'Chart2.Legends(0).Enabled = True
        'Chart2.Legends(0).Font = New Font("Segoe UI", 12.0F)

        ''--------//////////// Load date line CHART ///////////-----------------
        'Dim yValues3 As Double() = {"80", "20", "23", "111", "34", "12", "11", "112"} ' Getting values from Textboxes 
        'Dim xValues3 As String() = {"11/10/2018", "11/12/2018", "11/14/2018", "11/18/2018", "11/22/2018", "12/14/2018", "12/18/2018", "12/22/2018"} ' Headings
        'Dim seriesName3 As String = Nothing

        'Chart3.Series.Clear()
        'Chart3.Titles.Clear()

        'seriesName3 = "Material Shipping Graph"
        'Chart3.Series.Add(seriesName3)


        '' Bind X and Y values
        'Chart3.Series(seriesName3).Points.DataBindXY(xValues3, yValues3)
        'Chart3.Series(seriesName3).ChartType = SeriesChartType.Line


        'Dim T3 As Title = Chart3.Titles.Add("Material Shipping Graph")
        'With T3
        '    .ForeColor = Color.WhiteSmoke
        '    .BackColor = Color.FromArgb(38, 50, 76)
        '    .Font = New System.Drawing.Font("Segoe UI", 13.0F)

        'End With


        'Chart3.Series(seriesName3).Font = New Font("Segoe UI", 15.0F)
        'Chart3.Series(seriesName3).LabelForeColor = Color.WhiteSmoke
        'Chart3.Series(seriesName3).IsValueShownAsLabel = True

        'Chart3.Legends(0).Enabled = False

        'Chart3.ChartAreas("ChartArea1").AxisX.LabelStyle.ForeColor = Color.Gainsboro
        'Chart3.ChartAreas("ChartArea1").AxisY.LabelStyle.ForeColor = Color.Gainsboro

        missing_p = False

        Call Load_project("Project:  Not Specified", 100, "none")

        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null order by job"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    ComboBox1.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub


    Sub Load_project(project As String, mr_completed As Integer, mr_name As String)
        '========== Load Project data in dashboard ===============

        project_label.Text = project

        Try

            '----------- status ------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT Part_No, Description, Qty, qty_fullfilled, qty_needed, status_m  from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read

                    If (If(IsNumeric(reader4(2)), reader4(2), 0) - If(IsNumeric(reader4(3)), reader4(3), 0)) > 0 = True Or missing_p = False Then

                        total_grid.Rows.Add(New String() {})
                        total_grid.Rows(i).Cells(0).Value = reader4(0).ToString 'part
                        total_grid.Rows(i).Cells(1).Value = reader4(1).ToString  'desc
                        total_grid.Rows(i).Cells(2).Value = reader4(5).ToString  'status
                        total_grid.Rows(i).Cells(3).Value = "" 'apl stats
                        total_grid.Rows(i).Cells(4).Value = reader4(2).ToString  'qty d
                        total_grid.Rows(i).Cells(5).Value = reader4(3).ToString  'qty ful
                        total_grid.Rows(i).Cells(6).Value = If(IsNumeric(reader4(2)), reader4(2), 0) - If(IsNumeric(reader4(3)), reader4(3), 0) 'qty nee

                        If total_grid.Rows(i).Cells(6).Value > 0 Then
                            total_grid.Rows(i).DefaultCellStyle.BackColor = Color.Tan
                        End If

                        i = i + 1

                    End If

                End While

            End If

            reader4.Close()

            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To total_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        total_grid.Rows(i).Cells(7).Value = reader5(0).ToString
                        total_grid.Rows(i).Cells(9).Value = "Yes"
                    End While
                Else
                    total_grid.Rows(i).Cells(7).Value = 0
                    total_grid.Rows(i).Cells(9).Value = "No"
                End If

                reader5.Close()

                total_grid.Rows(i).Cells(8).Value = Real_Qty_transit(total_grid.Rows(i).Cells(0).Value, ComboBox1.Text)
                total_grid.Rows(i).Cells(3).Value = APL_analyzer(total_grid.Rows(i).Cells(0).Value, total_grid.Rows(i).Cells(6).Value)

            Next



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try




        '--------////////// Load MR PIE CHART //////////-----------------
        Dim yValues As Double() = {mr_completed, 100 - mr_completed} ' Getting values from Textboxes 
        Dim xValues As String() = {"Parts Fullfilled %", "Remainder %"} ' Headings
        Dim seriesName As String = Nothing

        Chart1.Series.Clear()
        Chart1.Titles.Clear()

        seriesName = "MR Status"
        Chart1.Series.Add(seriesName)


        ' Bind X and Y values
        Chart1.Series(seriesName).Points.DataBindXY(xValues, yValues)
        Chart1.Series(seriesName).Points(0).Color = Color.FromArgb(0, 64, 64)
        Chart1.Series(seriesName).Points(1).Color = Color.FromArgb(52, 52, 52)

        Chart1.Series(seriesName).ChartType = SeriesChartType.Pie

        Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = True

        Dim T As Title = Chart1.Titles.Add("Material Request Fullfillment Status")
        With T
            .ForeColor = Color.WhiteSmoke
            .BackColor = Color.FromArgb(61, 60, 88)
            .Font = New System.Drawing.Font("Segoe UI", 13.0F)

        End With

        Chart1.Series(seriesName).Points(0).LabelForeColor = Color.WhiteSmoke
        Chart1.Series(seriesName).Points(1).LabelForeColor = Color.WhiteSmoke
        Chart1.Series(seriesName).Font = New Font("Segoe UI", 15.0F)
        Chart1.Series(seriesName).IsValueShownAsLabel = True

        Chart1.Legends(0).Enabled = True
        Chart1.Legends(0).Font = New Font("Segoe UI", 12.0F)

        ''--------//////////// Load MFG PIE CHART ///////////-----------------
        'Dim yValues2 As Double() = {25, 25, 25, 25}
        'Dim xValues2 As String() = {"Panel", "Field", "Assembly", "Bulk"} ' Headings
        'Dim seriesName2 As String = Nothing

        'Chart2.Series.Clear()
        'Chart2.Titles.Clear()

        'seriesName2 = "MFG Type Status"
        'Chart2.Series.Add(seriesName2)


        ' Bind X and Y values
        'Chart2.Series(seriesName2).Points.DataBindXY(xValues2, yValues2)
        'Chart2.Series(seriesName2).Points(0).Color = Color.FromArgb(0, 64, 64)
        'Chart2.Series(seriesName2).Points(1).Color = Color.FromArgb(52, 52, 52)
        'Chart2.Series(seriesName2).Points(2).Color = Color.CadetBlue
        'Chart2.Series(seriesName2).Points(3).Color = Color.Maroon

        'Chart2.Series(seriesName2).ChartType = SeriesChartType.Pie

        'Chart2.ChartAreas("ChartArea1").Area3DStyle.Enable3D = True
        'Chart2.ChartAreas("ChartArea1").Area3DStyle.Inclination = 30

        'Dim T2 As Title = Chart2.Titles.Add("MFG Type Status")
        'With T2
        '    .ForeColor = Color.WhiteSmoke
        '    .BackColor = Color.FromArgb(38, 50, 76)
        '    .Font = New System.Drawing.Font("Segoe UI", 13.0F)

        'End With

        'Chart2.Series(seriesName2).Points(0).LabelForeColor = Color.WhiteSmoke
        'Chart2.Series(seriesName2).Points(1).LabelForeColor = Color.WhiteSmoke
        'Chart2.Series(seriesName2).Points(2).LabelForeColor = Color.WhiteSmoke
        'Chart2.Series(seriesName2).Points(3).LabelForeColor = Color.WhiteSmoke

        'Chart2.Series(seriesName2).Font = New Font("Segoe UI", 15.0F)
        'Chart2.Series(seriesName2).IsValueShownAsLabel = True

        'Chart2.Legends(0).Enabled = True
        'Chart2.Legends(0).Font = New Font("Segoe UI", 12.0F)



    End Sub



    Private Sub Chart1_DoubleClick(sender As Object, e As EventArgs) Handles Chart1.DoubleClick
        Monitor_MR.ShowDialog()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        '----- fill mr_box ----
        mr_box.Items.Clear()

        Try

            'Dim cmd_v As New MySqlCommand
            'cmd_v.Parameters.AddWithValue("@job", ComboBox1.Text)
            'cmd_v.CommandText = "SELECT distinct mr_name from Material_Request.mr where  job = @job order by mr_name desc"

            'cmd_v.Connection = Login.Connection
            'Dim readerv As MySqlDataReader
            'readerv = cmd_v.ExecuteReader

            'If readerv.HasRows Then
            '    While readerv.Read
            '        mr_box.Items.Add(readerv(0))
            '    End While
            'End If

            'readerv.Close()
            '------------------

            '------ get all BOM of the job Note: unelegant way of do this. Try to find a better Query
            Dim n_bom As Double : n_bom = 0
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
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
                check_cmd1.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
                check_cmd1.Parameters.AddWithValue("@id_bom", i)
                check_cmd1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                check_cmd1.Connection = Login.Connection
                check_cmd1.ExecuteNonQuery()

                Dim reader1 As MySqlDataReader
                reader1 = check_cmd1.ExecuteReader

                If reader1.HasRows Then
                    ' Dim j As Integer : j = 0
                    While reader1.Read
                        mr_box.Items.Add(reader1(0))
                    End While
                End If

                reader1.Close()
            Next


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            Try
                ComboBox1.Items.Clear()
                Dim cmd_v As New MySqlCommand
                cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null  and finished is null order by job"

                cmd_v.Connection = Login.Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_v.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        ComboBox1.Items.Add(readerv(0))
                    End While
                End If

                readerv.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            Try
                ComboBox1.Items.Clear()
                Dim cmd_v As New MySqlCommand
                cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null order by job"

                cmd_v.Connection = Login.Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_v.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        ComboBox1.Items.Add(readerv(0))
                    End While
                End If

                readerv.Close()
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Function cal_per(mr_name As String, mfg As String) As Double

        cal_per = 0
        Dim total_qty As Double : total_qty = 0
        Dim fullf As Double : fullf = 0


        Dim cmd3 As New MySqlCommand
        cmd3.Parameters.AddWithValue("@mr_name", mr_name)
        cmd3.Parameters.AddWithValue("@mfg_type", mfg)
        cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null and mfg_type = @mfg_type;"
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
        cmdx.Parameters.AddWithValue("@mfg_type", mfg)
        cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null and mfg_type = @mfg_type"
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

        If total_qty > 0 Then
            cal_per = Math.Round((fullf / total_qty) * 100)
        End If

    End Function

    Function Real_Qty_transit(part_name As String, job As String) As Double

        '====== qty transit against inventory and special order ======

        Real_Qty_transit = 0
        Dim not_inv As Boolean : not_inv = False

        Try
            Dim cmd5 As New MySqlCommand
            cmd5.Parameters.Clear()
            cmd5.Parameters.AddWithValue("@part_name", part_name)
            cmd5.CommandText = "SELECT Qty_in_order from inventory.inventory_qty where part_name = @part_name"
            cmd5.Connection = Login.Connection
            Dim reader5 As MySqlDataReader
            reader5 = cmd5.ExecuteReader

            If reader5.HasRows Then
                While reader5.Read
                    not_inv = False
                    Real_Qty_transit = reader5(0).ToString
                End While
            Else
                not_inv = True
            End If

            reader5.Close()

            If not_inv = True Then

                Dim cmd6 As New MySqlCommand
                cmd6.Parameters.Clear()
                cmd6.Parameters.AddWithValue("@part_name", part_name)
                cmd6.Parameters.AddWithValue("@job", job)
                cmd6.CommandText = "SELECT qty_needed from Tracking_Reports.my_tracking_reports where Part_No = @part_name and job = @job"
                cmd6.Connection = Login.Connection
                Dim reader6 As MySqlDataReader
                reader6 = cmd6.ExecuteReader

                If reader6.HasRows Then
                    While reader6.Read
                        Real_Qty_transit = reader6(0).ToString
                    End While
                End If

                reader6.Close()
            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function

    Function APL_analyzer(part_name As String, qty_needed As Double) As String

        '============ APL description ==============

        APL_analyzer = ""
        Dim lives_in As Boolean : lives_in = False
        Dim current_i As Double : current_i = 0

        Try
            If qty_needed > 0 Then

                '--check if lives in inventory
                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", part_name)
                cmd5.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        lives_in = True
                        current_i = reader5(0).ToString
                    End While
                End If

                reader5.Close()

                If lives_in = True Then  'if lives in inventory

                    If qty_needed - current_i > 0 Then  'if there is not enough in inventory

                        qty_needed = qty_needed - current_i

                        '-- check if there is some parts coming

                        Dim qty_transit As Double : qty_transit = 0
                        Dim es_date As String : es_date = "soon"

                        Dim cmd51 As New MySqlCommand
                        cmd51.Parameters.Clear()
                        cmd51.Parameters.AddWithValue("@part_name", part_name)
                        cmd51.CommandText = "SELECT Qty_in_order, es_date_of_arrival from inventory.inventory_qty where part_name = @part_name"
                        cmd51.Connection = Login.Connection
                        Dim reader51 As MySqlDataReader
                        reader51 = cmd5.ExecuteReader

                        If reader51.HasRows Then
                            While reader51.Read
                                qty_transit = reader51(0).ToString
                            End While

                        End If

                        reader51.Close()

                        If qty_needed - qty_transit > 0 Then
                            APL_analyzer = "(" & qty_needed - qty_transit & ") " & part_name & "  need to be ordered against inventory immediately"
                        Else
                            APL_analyzer = "(" & qty_transit & ") " & part_name & "  coming in " & es_date
                        End If


                    Else  'if there is enough in inventory
                        APL_analyzer = "(" & qty_needed & ") " & part_name & "  waiting to be taken from inventory"
                    End If

                Else  'if it doesnt live in inventory vheck tracking report

                    Dim qty_p As Double : qty_p = 0
                    Dim est_ar As String : est_ar = "soon" 'date of arrival
                    Dim PO_n As String : PO_n = "Not Specified"  'PO

                    Dim cmd51 As New MySqlCommand
                    cmd51.Parameters.Clear()
                    cmd51.Parameters.AddWithValue("@part_name", part_name)
                    cmd51.Parameters.AddWithValue("@job", ComboBox1.Text)
                    cmd51.CommandText = "SELECT qty_needed, es_date_of_arrival, PO from Tracking_Reports.my_tracking_reports  where Part_No = @part_name and job = @job"
                    cmd51.Connection = Login.Connection
                    Dim reader51 As MySqlDataReader
                    reader51 = cmd51.ExecuteReader

                    If reader51.HasRows Then
                        While reader51.Read
                            qty_p = reader51(0).ToString
                            est_ar = reader51(1).ToString
                            PO_n = reader51(2).ToString

                        End While

                    End If

                    reader51.Close()

                    If qty_needed - qty_p > 0 Then  'if what has been ordered is enough
                        APL_analyzer = "(" & qty_needed - qty_p & ") " & part_name & " need to be ordered immediately. This is a Special Order"
                    Else 'if what has been ordered satisfy the demand
                        APL_analyzer = "(" & qty_needed & ") " & part_name & " are on its way. Status: " & est_ar & "  PO: " & PO_n
                    End If

                End If

            Else
                APL_analyzer = "Quantity fullfilled, no action needed"
            End If



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Function

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles missing_box.CheckedChanged

        Try

            If missing_box.Checked = True Then

                missing_p = True
                Call reload_analyzer()

            Else
                missing_p = False
                Call reload_analyzer()

            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub reload_analyzer()

        'mr_box.Items.Clear()
        'Dim cmd_v As New MySqlCommand
        'cmd_v.Parameters.AddWithValue("@job", ComboBox1.Text)
        'cmd_v.CommandText = "SELECT distinct mr_name from Material_Request.mr where  job = @job order by mr_name desc"

        'cmd_v.Connection = Login.Connection
        'Dim readerv As MySqlDataReader
        'readerv = cmd_v.ExecuteReader

        'If readerv.HasRows Then
        '    While readerv.Read
        '        mr_box.Items.Add(readerv(0))
        '    End While
        'End If

        'readerv.Close()
        ''------------------


        'Dim project_l As String : project_l = ""
        'total_grid.Rows.Clear()


        'Dim cmd4 As New MySqlCommand
        'cmd4.Parameters.AddWithValue("@job", ComboBox1.Text)
        'cmd4.CommandText = "Select Job_description from management.projects where Job_number = @job"
        'cmd4.Connection = Login.Connection
        'Dim reader4 As MySqlDataReader
        'reader4 = cmd4.ExecuteReader

        'If reader4.HasRows Then
        '    While reader4.Read
        '        project_l = reader4(0).ToString
        '    End While
        'End If

        'reader4.Close()

        ''-------- get mr_name -------
        'Dim name2 As String : name2 = "xxxx5xxxzzz"

        'Dim cmd41 As New MySqlCommand
        'cmd41.Parameters.AddWithValue("@job", ComboBox1.Text)
        'cmd41.CommandText = "select mr_name from Material_Request.mr where job = @job order by release_date desc limit 1"
        'cmd41.Connection = Login.Connection
        'Dim reader41 As MySqlDataReader
        'reader41 = cmd41.ExecuteReader

        'If reader41.HasRows Then
        '    While reader41.Read
        '        name2 = reader41(0).ToString
        '    End While
        'End If


        'reader41.Close()

        ''---------- calculate percentage ---------
        'Dim complete_mr As Double : complete_mr = 0
        'Dim total_qty As Double : total_qty = 0
        'Dim fullf As Double : fullf = 0


        'Dim cmd3 As New MySqlCommand
        'cmd3.Parameters.AddWithValue("@mr_name", name2)
        'cmd3.CommandText = "select sum(qty_fullfilled) from Material_Request.mr_data where mr_name = @mr_name and qty_fullfilled is not null;"
        'cmd3.Connection = Login.Connection
        'Dim reader3 As MySqlDataReader
        'reader3 = cmd3.ExecuteReader

        'If reader3.HasRows Then
        '    While reader3.Read
        '        If IsDBNull(reader3(0)) Then
        '            fullf = 0
        '        Else
        '            fullf = CType(reader3(0), Double)
        '        End If
        '    End While
        'End If

        'reader3.Close()

        'Dim cmdx As New MySqlCommand
        'cmdx.Parameters.AddWithValue("@mr_name", name2)
        'cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null"
        'cmdx.Connection = Login.Connection
        'Dim readerx As MySqlDataReader
        'readerx = cmdx.ExecuteReader

        'If readerx.HasRows Then
        '    While readerx.Read
        '        If IsDBNull(readerx(0)) Then
        '            total_qty = 0
        '        Else
        '            total_qty = CType(readerx(0), Double)
        '        End If
        '    End While
        'End If

        'readerx.Close()

        'If total_qty > 0 Then
        '    complete_mr = Math.Round((fullf / total_qty) * 100)
        'End If

        'Call Load_project(ComboBox1.Text & "   " & project_l, complete_mr, name2)

        Call Dash_info(mr_box.Text)

    End Sub


    '--- This displays all info of a BOM
    Sub Dash_info(mr_name As String)

        '----- fill mr_box ----
        Try

            Dim project_l As String : project_l = ""
            total_grid.Rows.Clear()
            '  close_grid.Rows.Clear()


            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd4.CommandText = "Select Job_description from management.projects where Job_number = @job"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    project_l = reader4(0).ToString
                End While
            End If

            reader4.Close()


            '------------ date mr was released ------------

            Dim date1 As DateTime
            Dim id_temp As String : id_temp = 1

            '---------- id_bom --
            Dim cmd421 As New MySqlCommand
            cmd421.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd421.Parameters.AddWithValue("@mr_name", mr_name)
            cmd421.CommandText = "select id_bom from Material_Request.mr where job = @job and mr_name = @mr_name"
            cmd421.Connection = Login.Connection
            Dim reader421 As MySqlDataReader
            reader421 = cmd421.ExecuteReader

            If reader421.HasRows Then
                While reader421.Read
                    id_temp = reader421(0)
                End While
            End If


            reader421.Close()


            Dim cmd42 As New MySqlCommand
            cmd42.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd42.Parameters.AddWithValue("@id_bom", id_temp)
            cmd42.CommandText = "select release_date from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date asc limit 1"
            cmd42.Connection = Login.Connection
            Dim reader42 As MySqlDataReader
            reader42 = cmd42.ExecuteReader

            If reader42.HasRows Then
                While reader42.Read
                    date1 = reader42(0)
                End While
            End If


            reader42.Close()

            Label1.Text = "Days Passed Since MR was released: " & DateDiff(DateInterval.Day, date1.Date, Now.Date)
            Label2.Text = "Material Request Released Date: " & date1

            '---------- calculate percentage ---------
            Dim complete_mr As Double : complete_mr = 0
            Dim total_qty As Double : total_qty = 0
            Dim fullf As Double : fullf = 0


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
            cmdx.CommandText = "select sum(QTY) from Material_Request.mr_data where mr_name = @mr_name and  QTY is not null"
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

            If total_qty > 0 Then
                complete_mr = Math.Round((fullf / total_qty) * 100)
            End If

            '/////////// --------- Calculate Percentage MFG Type -------////////////

            Label3.Text = "Panel Parts Fullfilled: " & cal_per(mr_name, "Panel") & "%"
            Label4.Text = "Field Parts Fullfilled: : " & cal_per(mr_name, "Field") & "%"
            Label5.Text = "Assembly Parts Fullfilled: " & cal_per(mr_name, "Assembly") & "%"

            '-------------------------------------------------------------

            Call Load_project(ComboBox1.Text & "   " & project_l, complete_mr, mr_name)

            '----- history table -----

            history_grid.Rows.Clear()
            Dim cmd5 As New MySqlCommand
            cmd5.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd5.CommandText = "SELECT username, action_m, date_m , role from management.history where job = @job"
            cmd5.Connection = Login.Connection
            Dim reader5 As MySqlDataReader
            reader5 = cmd5.ExecuteReader

            If reader5.HasRows Then
                Dim i As Integer : i = 0
                While reader5.Read
                    history_grid.Rows.Add(New String() {})
                    history_grid.Rows(i).Cells(0).Value = reader5(0).ToString
                    history_grid.Rows(i).Cells(1).Value = reader5(1).ToString
                    history_grid.Rows(i).Cells(2).Value = reader5(2).ToString
                    history_grid.Rows(i).Cells(3).Value = reader5(3).ToString
                    i = i + 1
                End While

            End If

            reader5.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub mr_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles mr_box.SelectedValueChanged
        Call Dash_info(mr_box.Text)
    End Sub


End Class