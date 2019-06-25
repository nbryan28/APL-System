Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class Procurement_Overview

    Public open_job As String
    Public toggle As Boolean
    Public go_ahead As Boolean

    Public plabels() As Label
    Public alabels() As Label

    Public ASM_mode As Boolean

    Public temp_qty_panel As Integer

    Private Sub Procurement_Overview_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        temp_qty_panel = 1
        ASM_mode = False
        ToolTip1.IsBalloon = True
        '  Call Load_BOM_map("113185_Master_BOM", "113185")

        Call EnableDoubleBuffered(total_grid)
        Call EnableDoubleBuffered(fullfill_grid)
        Call EnableDoubleBuffered(orders_grid)
        Call EnableDoubleBuffered(asm_grid)

        Me.DoubleBuffered = True


        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.TabPages.Remove(TabPage10)
        total_grid.Columns(11).Visible = False 'add return item column hide
        open_job = ""
        Timer1.Interval = 13000
        Timer1.Start()
        go_ahead = True



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
            Dim tronco(list.Count - 1) As Label

            '   Dim alabels(list_as.Count - 1) As Label
            ReDim alabels(list_as.Count - 1)
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
                tronco(i).Location = New Point(ini_y, y) '286
                tronco(i).BackColor = Color.DimGray
                tronco(i).Size = New System.Drawing.Size(10, 22)
                tronco(i).Anchor = AnchorStyles.Top

                TabPage5.Controls.Add(tronco(i))
                y = y + 96

                'panel blocks

                plabels(i) = New Label()
                plabels(i).Size = New System.Drawing.Size(267, 74)   '267,  74
                plabels(i).BackColor = Color.Teal
                plabels(i).Text = list.Item(i).ToString
                plabels(i).ForeColor = Color.WhiteSmoke
                plabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
                plabels(i).Anchor = AnchorStyles.Top
                plabels(i).Location = New Point(ini_y - CType((w_i) / 2, Integer), z)  '161
                plabels(i).TextAlign = ContentAlignment.MiddleCenter


                TabPage5.Controls.Add(plabels(i))
                z = z + 96

                Dim temp_i As Integer : temp_i = i
                AddHandler plabels(i).Click, Sub(sender, e) Module_show(plabels(temp_i).Text, job)


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

                TabPage5.Controls.Add(tronco2(i))
                y2 = y2 + 96

                'panel blocks

                alabels(i) = New Label()
                alabels(i).Size = New System.Drawing.Size(267, 74)   '267, 74
                alabels(i).BackColor = Color.Teal
                alabels(i).Text = list_as.Item(i).ToString
                alabels(i).ForeColor = Color.WhiteSmoke
                alabels(i).Font = New Font("Segoe UI", 12, FontStyle.Regular)
                alabels(i).Anchor = AnchorStyles.Top
                alabels(i).Location = New Point(x_ia, z2)  '817
                alabels(i).TextAlign = ContentAlignment.MiddleCenter

                TabPage5.Controls.Add(alabels(i))
                z2 = z2 + 96

                Dim temp_i As Integer : temp_i = i
                AddHandler alabels(i).Click, Sub(sender, e) Module_show(alabels(temp_i).Text, job)


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

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        'home
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Material_order.Visible = True
    End Sub

    Private Sub OpenJobToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenJobToolStripMenuItem.Click

        If Log_out = True Then

            MR_init.Visible = True
            Me.Visible = False
        Else
            Me.Visible = False
        End If

    End Sub

    Private Sub ProjectTotalsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProjectTotalsToolStripMenuItem.Click
        Call Export_totals()
    End Sub

    Sub Export_special_order()
        '-- exports special order table
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
                For i As Integer = 0 To orders_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = orders_grid.Columns(i).HeaderText
                    For j As Integer = 0 To orders_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = orders_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)

                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Special Order table exported successfully!")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Export_comparison()
        '-- exports comparison table
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
                For i As Integer = 0 To compare_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = compare_grid.Columns(i).HeaderText
                    For j As Integer = 0 To compare_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = compare_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)

                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Comparison Table exported successfully!")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Export_totals()
        '-- exports project totals table
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

                Cursor.Current = Cursors.WaitCursor

                'copy data to excel
                For i As Integer = 0 To total_grid.ColumnCount - 1
                    xlWorkSheet.Cells(1, i + 1) = total_grid.Columns(i).HeaderText
                    For j As Integer = 0 To total_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = total_grid.Rows(j).Cells(i).Value
                    Next j
                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                Cursor.Current = Cursors.Default

                MessageBox.Show("Projects Totals table exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Sub Export_fullfill()
        '-- exports project totals table
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
                For i As Integer = 0 To fullfill_grid.ColumnCount - 1
                    xlWorkSheet.Cells(1, i + 1) = fullfill_grid.Columns(i).HeaderText
                    For j As Integer = 0 To fullfill_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = fullfill_grid.Rows(j).Cells(i).Value
                    Next j
                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Fullfillment table exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Sub Export_asm()
        '-- exports project totals table
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
                For i As Integer = 0 To asm_grid.ColumnCount - 1
                    xlWorkSheet.Cells(1, i + 1) = asm_grid.Columns(i).HeaderText
                    For j As Integer = 0 To asm_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = asm_grid.Rows(j).Cells(i).Value
                    Next j
                Next i

                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                xlWorkBook.Close(False)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Assembly table exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub SpecialOrdersTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecialOrdersTableToolStripMenuItem.Click
        Call Export_special_order()
    End Sub

    Private Sub ComparisonTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ComparisonTableToolStripMenuItem.Click
        Call Export_comparison()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Export_totals()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Export_special_order()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call Export_comparison()
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        '-- refresn on/off
        If String.Equals(Label15.Text, "Refresh off") = True Then

            Label15.Text = "Refresh on"
            toggle = True
        Else
            Label15.Text = "Refresh off"
            toggle = False
        End If
    End Sub

    Sub open_my_job(mr_name As String)
        '---/////////////////  open a job temp sub for test ///////////////////////

        '----------- Open a release MR ----------
        total_grid.Rows.Clear()
        orders_grid.Rows.Clear()
        compare_grid.Rows.Clear()
        MO_grid.Rows.Clear()
        fullfill_grid.Rows.Clear()

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Project Qty", GetType(String))
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("MFG Type", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("subtotal", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("need_by_date", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))
        dimen_table.Columns.Add("Assembly_name", GetType(String))

        Dim job As String : job = ""

        Try

            '--------- get job number --------

            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@mr_name", mr_name)
            cmd41.CommandText = "SELECT job from Material_Request.mr where mr_name = @mr_name"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    job = reader41(0).ToString
                End While
            End If

            reader41.Close()


            '--------- get shipping address and need by date 
            Dim cmd51 As New MySqlCommand
            cmd51.Parameters.Clear()
            cmd51.Parameters.AddWithValue("@mr_name", mr_name)
            cmd51.CommandText = "SELECT shipping_ad, need_date from Material_Request.mr where mr_name = @mr_name"
            cmd51.Connection = Login.Connection
            Dim reader51 As MySqlDataReader
            reader51 = cmd51.ExecuteReader

            If reader51.HasRows Then
                While reader51.Read
                    shipping_b.Text = reader51(0).ToString & "   Need by Date: "

                    If reader51.IsDBNull(1) = False Then
                        shipping_b.Text = shipping_b.Text & reader51(1).ToString
                    End If
                End While
            End If

            reader51.Close()

            '-------------------------------------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT Qty, Part_No, Description, Manufacturer, mfg_type, Vendor, Price, subtotal,  qty_fullfilled, need_by_date, Notes, Assembly_name from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, (If(IsNumeric(reader4(6)), reader4(6), 0) * If(IsNumeric(reader4(0)), reader4(0), 0)), reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString)
                End While

            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                total_grid.Rows.Add(New String() {})
                total_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString 'project qty
                total_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString 'Part_No
                total_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString  'Description
                total_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString 'Manufacturer
                total_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString   'MFG Type
                total_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString  'vendor
                total_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString  'Price
                total_grid.Rows(i).Cells(7).Value = If(IsNumeric(dimen_table.Rows(i).Item(7).ToString), CType(dimen_table.Rows(i).Item(7).ToString, Double).ToString("N"), 0) 'subtotal
                total_grid.Rows(i).Cells(8).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)   'fulfilled
                total_grid.Rows(i).Cells(9).Value = 0  'needed
                total_grid.Rows(i).Cells(10).Value = 0 'on hand
                total_grid.Rows(i).Cells(11).Value = 0 '+/-
                total_grid.Rows(i).Cells(12).Value = 0 'in transit
                total_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(9).ToString 'need date
                total_grid.Rows(i).Cells(14).Value = 0 'min
                total_grid.Rows(i).Cells(15).Value = 0 'max
                total_grid.Rows(i).Cells(16).Value = dimen_table.Rows(i).Item(10).ToString 'notes
                total_grid.Rows(i).Cells(17).Value = dimen_table.Rows(i).Item(11).ToString 'assembly
            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To total_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                cmd5.CommandText = "SELECT current_qty,  Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read

                        total_grid.Rows(i).Cells(10).Value = reader5(0)
                        total_grid.Rows(i).Cells(12).Value = reader5(1).ToString
                        total_grid.Rows(i).Cells(14).Value = reader5(2).ToString 'min
                        total_grid.Rows(i).Cells(15).Value = reader5(3).ToString 'max

                    End While
                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

            mrbox1.Items.Clear()
            mrbox2.Items.Clear()


            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", job)
            check_cmd.CommandText = "select distinct mr_name from Material_Request.mr where released = 'Y' and job = @job"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                While reader.Read
                    mrbox1.Items.Add(reader(0))
                    mrbox2.Items.Add(reader(0))
                End While
            End If

            reader.Close()

            '------------------------------------

            Me.Text = mr_name
            open_job = job
            job_label.Text = job

            Call Load_BOM_map(mr_name, job)  'LOAD MAP
            Call Report_track_load(mr_name, job) 'load track report (special parts)
            Call Inventory_project() 'load inventory parts request by project
            Call MPL_open(job)  'open MPL in grid

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub total_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles total_grid.CellValueChanged

        '-- recalculation for project totals. get need by date
        If go_ahead = True Then

            For Each row As DataGridViewRow In total_grid.Rows
                If row.IsNewRow Then Continue For

                If (IsNumeric(row.Cells(0).Value) = True And IsNumeric(row.Cells(11).Value)) Then

                    If CType(row.Cells(10).Value, Double) < CType(row.Cells(11).Value, Double) Then
                        row.Cells(11).Value = 0
                    End If

                    If (IsNumeric(row.Cells(0).Value) = True And IsNumeric(row.Cells(8).Value)) Then
                        row.Cells(9).Value = CType(row.Cells(0).Value, Double) - CType(row.Cells(8).Value, Double)
                    Else
                        row.Cells(9).Value = If(IsNumeric(row.Cells(0).Value) = False, 0, row.Cells(0).Value)
                    End If


                    If CType(row.Cells(9).Value, Double) <> 0 Then
                        row.Cells(9).Style.BackColor = Color.Firebrick
                    Else
                        row.Cells(9).Style.BackColor = Color.Gray
                    End If

                End If
            Next

        End If
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================

        If TabControl1.SelectedTab Is TabPage3 Then

            If total_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(total_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage1 Then
            If orders_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(orders_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage2 Then
            If compare_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(compare_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage4 Then
            If MO_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(MO_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage6 Then

            If fullfill_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(fullfill_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If

        ElseIf TabControl1.SelectedTab Is TabPage10 Then

            If asm_grid.CurrentCell.Value.ToString <> String.Empty Then
                Clipboard.SetText(asm_grid.CurrentCell.Value.ToString)
            Else
                Clipboard.Clear()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- compare select mr
        If Not mrbox1.SelectedItem Is Nothing And Not mrbox2.SelectedItem Is Nothing Then

            compare_grid.Rows.Clear()

            Dim total_mr As New Dictionary(Of String, String)
            '-- store mr1 parts
            Dim dimen_table1 = New DataTable
            dimen_table1.Columns.Add("Part_No", GetType(String))
            dimen_table1.Columns.Add("Description", GetType(String))

            '-- store mr2 parts
            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("Description", GetType(String))

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                cmd4.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        dimen_table1.Rows.Add(reader4(0).ToString, reader4(1).ToString)
                    End While
                End If

                reader4.Close()

                '--- mr2
                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                cmd41.CommandText = "SELECT Part_No, Description from Material_Request.mr_data where mr_name = @mr_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        dimen_table2.Rows.Add(reader41(0).ToString, reader41(1).ToString)
                    End While
                End If

                reader41.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            For i = 0 To dimen_table1.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table1.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table1.Rows(i).Item(0), dimen_table1.Rows(i).Item(1))
                End If
            Next

            For i = 0 To dimen_table2.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table2.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table2.Rows(i).Item(0), dimen_table2.Rows(i).Item(1))
                End If
            Next

            '--- add to comp_Grid
            For Each kvp As KeyValuePair(Of String, String) In total_mr.ToArray
                compare_grid.Rows.Add(New String() {kvp.Key, total_mr(kvp.Key)})
            Next

            For i = 0 To compare_grid.Rows.Count - 1
                Try
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.Clear()
                    cmd41.Parameters.AddWithValue("@mr_name", mrbox1.Text)
                    cmd41.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd41.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            compare_grid.Rows(i).Cells(2).Value = reader41(0).ToString
                            compare_grid.Rows(i).Cells(4).Value = reader41(1).ToString
                        End While
                    End If

                    reader41.Close()

                    Dim cmd42 As New MySqlCommand
                    cmd42.Parameters.Clear()
                    cmd42.Parameters.AddWithValue("@mr_name", mrbox2.Text)
                    cmd42.Parameters.AddWithValue("@part", compare_grid.Rows(i).Cells(0).Value)
                    cmd42.CommandText = "SELECT mfg_type, QTY from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part"
                    cmd42.Connection = Login.Connection
                    Dim reader42 As MySqlDataReader
                    reader42 = cmd42.ExecuteReader

                    If reader42.HasRows Then
                        While reader42.Read
                            compare_grid.Rows(i).Cells(3).Value = reader42(0).ToString
                            compare_grid.Rows(i).Cells(5).Value = reader42(1).ToString
                        End While
                    End If

                    reader42.Close()

                    compare_grid.Rows(i).Cells(6).Value = If(IsNumeric(compare_grid.Rows(i).Cells(5).Value), compare_grid.Rows(i).Cells(5).Value, 0) - If(IsNumeric(compare_grid.Rows(i).Cells(4).Value), compare_grid.Rows(i).Cells(4).Value, 0)

                    If compare_grid.Rows(i).Cells(6).Value <> 0 Then
                        compare_grid.Rows(i).Cells(6).Style.BackColor = Color.CadetBlue
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Next



        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TabControl1.SelectedTab Is TabPage3 Then

            '---- MBOM fullfill search

            total_grid.Rows.Clear()

            '---- datatable store MR without assemblies --------
            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Project Qty", GetType(String))
            dimen_table.Columns.Add("Part_No", GetType(String))
            dimen_table.Columns.Add("Description", GetType(String))
            dimen_table.Columns.Add("Manufacturer", GetType(String))
            dimen_table.Columns.Add("MFG Type", GetType(String))
            dimen_table.Columns.Add("Vendor", GetType(String))
            dimen_table.Columns.Add("Price", GetType(String))
            dimen_table.Columns.Add("subtotal", GetType(String))
            dimen_table.Columns.Add("qty_fullfilled", GetType(String))
            dimen_table.Columns.Add("need_by_date", GetType(String))
            dimen_table.Columns.Add("Notes", GetType(String))
            dimen_table.Columns.Add("Assembly_name", GetType(String))

            Try

                '-------------------------------------
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd4.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")
                cmd4.CommandText = "SELECT Qty, Part_No, Description, Manufacturer, mfg_type, Vendor, Price, subtotal,  qty_fullfilled, need_by_date, Notes, Assembly_name from Material_Request.mr_data where mr_name = @mr_name and Part_No LIKE @search"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader


                If reader4.HasRows Then
                    Dim i As Integer : i = 0
                    While reader4.Read
                        dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, (If(IsNumeric(reader4(6)), reader4(6), 0) * If(IsNumeric(reader4(0)), reader4(0), 0)), reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString)
                    End While

                End If

                reader4.Close()


                '------- now copy dimen_table_a to inventroy grid
                For i = 0 To dimen_table.Rows.Count - 1
                    total_grid.Rows.Add(New String() {})
                    total_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString 'project qty
                    total_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString 'Part_No
                    total_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString  'Description
                    total_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString 'Manufacturer
                    total_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString   'MFG Type
                    total_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString  'vendor
                    total_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString  'Price
                    total_grid.Rows(i).Cells(7).Value = If(IsNumeric(dimen_table.Rows(i).Item(7).ToString), CType(dimen_table.Rows(i).Item(7).ToString, Double).ToString("N"), 0)  'subtotal
                    total_grid.Rows(i).Cells(8).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)   'fulfilled
                    total_grid.Rows(i).Cells(9).Value = 0  'needed
                    total_grid.Rows(i).Cells(10).Value = 0 'on hand
                    total_grid.Rows(i).Cells(11).Value = 0 '+/-
                    total_grid.Rows(i).Cells(12).Value = 0 'in transit
                    total_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(9).ToString 'need date
                    total_grid.Rows(i).Cells(14).Value = 0 'min
                    total_grid.Rows(i).Cells(15).Value = 0 'max
                    total_grid.Rows(i).Cells(16).Value = dimen_table.Rows(i).Item(10).ToString 'notes
                    total_grid.Rows(i).Cells(17).Value = dimen_table.Rows(i).Item(11).ToString 'assembly
                Next


                '////////////////////  ---------- fill current inventory values -----------------------

                For i = 0 To total_grid.Rows.Count - 1

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                    cmd5.CommandText = "SELECT current_qty,  Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    Dim reader5 As MySqlDataReader
                    reader5 = cmd5.ExecuteReader

                    If reader5.HasRows Then
                        While reader5.Read
                            total_grid.Rows(i).Cells(10).Value = reader5(0).ToString
                            total_grid.Rows(i).Cells(12).Value = reader5(1).ToString
                            total_grid.Rows(i).Cells(14).Value = reader5(2).ToString 'min
                            total_grid.Rows(i).Cells(15).Value = reader5(3).ToString 'max
                        End While
                    End If
                    reader5.Close()

                Next
                '----------------------------------------------------------------             

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try


        ElseIf TabControl1.SelectedTab Is TabPage6 Then

            fullfill_grid.Rows.Clear()

            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Part_No", GetType(String))
            dimen_table.Columns.Add("Description", GetType(String))
            dimen_table.Columns.Add("Manufacturer", GetType(String))
            dimen_table.Columns.Add("Assembly_name", GetType(String))
            dimen_table.Columns.Add("Vendor", GetType(String))
            dimen_table.Columns.Add("Price", GetType(String))
            dimen_table.Columns.Add("mfg_type", GetType(String))
            dimen_table.Columns.Add("Qty", GetType(String))
            dimen_table.Columns.Add("qty_fullfilled", GetType(String))
            dimen_table.Columns.Add("qty_needed", GetType(String))
            dimen_table.Columns.Add("Notes", GetType(String))


            Try

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mr_name", bom_open.Text)
                cmd4.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")
                cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, Assembly_name, Vendor, Price, mfg_type, Qty, qty_fullfilled, qty_needed , Notes from Material_Request.mr_data where mr_name = @mr_name and Part_No LIKE @search"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader


                If reader4.HasRows Then
                    While reader4.Read
                        dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString * temp_qty_panel, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString)
                    End While
                End If

                reader4.Close()


                '------- now copy dimen_table_a to inventroy grid
                For i = 0 To dimen_table.Rows.Count - 1
                    fullfill_grid.Rows.Add(New String() {})
                    fullfill_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                    fullfill_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                    fullfill_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                    fullfill_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                    fullfill_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                    fullfill_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                    fullfill_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString
                    fullfill_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString
                    fullfill_grid.Rows(i).Cells(8).Value = 0
                    fullfill_grid.Rows(i).Cells(9).Value = 0
                    fullfill_grid.Rows(i).Cells(10).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)
                    fullfill_grid.Rows(i).Cells(11).Value = If(IsNumeric(dimen_table.Rows(i).Item(9).ToString) = True, dimen_table.Rows(i).Item(9).ToString, 0)
                    fullfill_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(10).ToString

                Next


                '////////////////////  ---------- fill current inventory values -----------------------

                For i = 0 To fullfill_grid.Rows.Count - 1

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.CommandText = "SELECT current_qty , Qty_in_order from inventory.inventory_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    Dim reader5 As MySqlDataReader
                    reader5 = cmd5.ExecuteReader

                    If reader5.HasRows Then
                        While reader5.Read
                            fullfill_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                            fullfill_grid.Rows(i).Cells(12).Value = reader5(1).ToString
                        End While
                    End If

                    reader5.Close()

                Next
                '----------------------------------------------------------------

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub

    Private Sub AcknowledgeBOMSenderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcknowledgeBOMSenderToolStripMenuItem.Click
        '--- This menu send a APL message to the person that released the BOM
        If String.Equals(Me.Text, "My Material Requests") = False Then
            Try
                '--- get username of person that release the BOM 
                Dim user_n As String : user_n = ""
                Dim role_m As String : role_m = ""
                Dim role_s As String : role_s = "General Management"

                Dim cmd_v2 As New MySqlCommand

                cmd_v2.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd_v2.CommandText = "SELECT released_by from Material_Request.mr where mr_name = @mr_name"
                cmd_v2.Connection = Login.Connection
                Dim readerv2 As MySqlDataReader
                readerv2 = cmd_v2.ExecuteReader

                If readerv2.HasRows Then
                    While readerv2.Read
                        user_n = readerv2(0).ToString
                    End While
                End If

                readerv2.Close()
                '------------------------------------
                '-- get role of user -----
                Dim cmd_v21 As New MySqlCommand

                cmd_v21.Parameters.AddWithValue("@user", user_n)
                cmd_v21.CommandText = "SELECT Role from users where username = @user"
                cmd_v21.Connection = Login.Connection
                Dim readerv21 As MySqlDataReader
                readerv21 = cmd_v21.ExecuteReader

                If readerv21.HasRows Then
                    While readerv21.Read
                        role_m = readerv21(0).ToString
                    End While
                End If

                readerv21.Close()
                '---------------------------

                'send APL message
                Dim Create_cmd As New MySqlCommand

                Dim mail_n As String : mail_n = "BOM has been received successfully. The Procurement team will begin allocating and purchasing the parts you requested" & vbCrLf & vbCrLf _
                         & "Project: " & job_label.Text & vbCrLf _
                         & "BOM File Name: " & Me.Text & vbCrLf

                Create_cmd.Parameters.AddWithValue("@Sender", current_user)
                Create_cmd.Parameters.AddWithValue("@role_s", role_s)
                Create_cmd.Parameters.AddWithValue("@Receiver", user_n)
                Create_cmd.Parameters.AddWithValue("@role_r", role_m)
                Create_cmd.Parameters.AddWithValue("@priority", "n")
                Create_cmd.Parameters.AddWithValue("@date_s", Now)
                Create_cmd.Parameters.AddWithValue("@read_m", "n")
                Create_cmd.Parameters.AddWithValue("@Mail", mail_n)
                Create_cmd.Parameters.AddWithValue("@title", "BOM Confirmation Message")

                Create_cmd.CommandText = "INSERT INTO management.Dropbox(Sender, role_s, Receiver, role_r, priority, date_s, read_m, Mail, title) VALUES (@Sender,@role_s,@Receiver,@role_r,@priority,@date_s,@read_m,@Mail,@title)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()


                MessageBox.Show("Acknowldge Message sent!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Report_track_load(mr_name As String, job As String)

        '--load the report tracking in the special order table
        Dim my_assemblies = New List(Of String)()

        Try
            '--------------  add to device -----------------
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

            Dim id_bom As String : id_bom = 1
            Dim cmd21 As New MySqlCommand
            cmd21.Parameters.AddWithValue("@mr_name", mr_name)
            cmd21.CommandText = "SELECT id_bom from Material_Request.mr where mr_name = @mr_name"
            cmd21.Connection = Login.Connection
            Dim reader21 As MySqlDataReader
            reader21 = cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    id_bom = reader21(0).ToString
                End While
            End If

            reader21.Close()

            '-----------------------------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@job", job)
            cmd4.Parameters.AddWithValue("@id_bom", id_bom)
            cmd4.CommandText = "SELECT Part_No, qty_needed, PO, es_date_of_arrival, mfg, primary_vendor, cost from Tracking_Reports.my_tracking_reports where job = @job and id_bom = @id_bom"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    orders_grid.Rows.Add(New String() {})
                    orders_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    orders_grid.Rows(i).Cells(4).Value = reader4(1).ToString
                    orders_grid.Rows(i).Cells(5).Value = reader4(2).ToString
                    orders_grid.Rows(i).Cells(6).Value = reader4(3).ToString
                    orders_grid.Rows(i).Cells(1).Value = reader4(4).ToString
                    orders_grid.Rows(i).Cells(2).Value = reader4(5).ToString
                    orders_grid.Rows(i).Cells(3).Value = reader4(6).ToString
                    i = i + 1
                End While

            End If

            reader4.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        For i = 0 To total_grid.Rows.Count - 1

            Try

                '--check if part exist in inventory
                Dim exist_c As Boolean : exist_c = False
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                check_cmd.CommandText = "select * from inventory.inventory_qty where part_name = @part_name and min_qty > 0"  'and min_qty > 0
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then
                    exist_c = True
                End If

                reader.Close()

                If exist_c = False Then

                    '------------ if it's an assembly not kept in inventory then
                    If my_assemblies.Contains(total_grid.Rows(i).Cells(1).Value) = True Then
                        Call Get_need_buy(total_grid.Rows(i).Cells(1).Value, total_grid.Rows(i).Cells(0).Value)
                    Else

                        If orders_grid.Rows.Count > 0 Then

                            Dim new_qty As Double = Isintable(total_grid.Rows(i).Cells(1).Value, total_grid.Rows(i).Cells(0).Value)

                            If new_qty > 0 Then
                                orders_grid.Rows.Add(New String() {total_grid.Rows(i).Cells(1).Value, total_grid.Rows(i).Cells(3).Value, total_grid.Rows(i).Cells(5).Value, total_grid.Rows(i).Cells(6).Value, new_qty})
                            End If

                        Else
                            orders_grid.Rows.Add(New String() {})
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(0).Value = total_grid.Rows(i).Cells(1).Value 'part name
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(1).Value = total_grid.Rows(i).Cells(3).Value 'manu
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(2).Value = total_grid.Rows(i).Cells(5).Value  'vendor
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(3).Value = total_grid.Rows(i).Cells(6).Value  'cost
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(4).Value = total_grid.Rows(i).Cells(0).Value  'qty
                        End If
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        Next

    End Sub

    Function Isintable(part_name As String, qty As Double) As Double

        Isintable = 0

        For i = 0 To orders_grid.Rows.Count - 1
            If String.Equals(part_name, orders_grid.Rows(i).Cells(0).Value.ToString) = True Then
                qty = qty - If(IsNumeric(orders_grid.Rows(i).Cells(4).Value) = True, CType(orders_grid.Rows(i).Cells(4).Value, Double), 0)
            End If
        Next

        Isintable = qty
    End Function

    '---- This sub takes an assembly break it apart and pick the parts not in inventory and add them to the tracking report
    Sub Get_need_buy(assembly As String, qty As Double)

        Dim datatable = New DataTable
        datatable.Columns.Add("part_name", GetType(String))
        datatable.Columns.Add("manufacturer", GetType(String))
        datatable.Columns.Add("primary_vendor", GetType(String))
        datatable.Columns.Add("cost", GetType(Double))
        datatable.Columns.Add("qty", GetType(Double))

        Try
            Dim cmd3 As New MySqlCommand
            cmd3.Parameters.AddWithValue("@ADA", assembly)
            cmd3.CommandText = "SELECT p1.Part_Name, p1.Manufacturer, p1.Primary_Vendor, adv.Qty from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.Legacy_ADA  = @ADA"
            cmd3.Connection = Login.Connection
            Dim reader3 As MySqlDataReader
            reader3 = cmd3.ExecuteReader

            If reader3.HasRows Then
                While reader3.Read
                    datatable.Rows.Add(reader3(0).ToString, reader3(1).ToString, reader3(2).ToString, 0, reader3(3) * qty)
                End While
            End If

            reader3.Close()

            For i = 0 To datatable.Rows.Count - 1
                datatable.Rows(i).Item(3) = Form1.Get_Latest_Cost(Login.Connection, datatable.Rows(i).Item(0), datatable.Rows(i).Item(2))
            Next


            '------//////// determine which of these list of parts need to be buy
            For i = 0 To datatable.Rows.Count - 1

                Try
                    '--check if part exist in inventory
                    Dim exist_c As Boolean : exist_c = False
                    Dim check_cmd As New MySqlCommand
                    check_cmd.Parameters.AddWithValue("@part_name", datatable.Rows(i).Item(0))
                    check_cmd.CommandText = "select * from inventory.inventory_qty where part_name = @part_name And min_qty > 0" 'And min_qty > 0
                    check_cmd.Connection = Login.Connection
                    check_cmd.ExecuteNonQuery()

                    Dim reader As MySqlDataReader
                    reader = check_cmd.ExecuteReader

                    If reader.HasRows Then
                        exist_c = True
                    End If

                    reader.Close()

                    If exist_c = False Then

                        orders_grid.Rows.Add(New String() {})
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(0).Value = datatable.Rows(i).Item(0) 'part name
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(1).Value = datatable.Rows(i).Item(1) 'manu
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(2).Value = datatable.Rows(i).Item(2)  'vendor
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(3).Value = datatable.Rows(i).Item(3)  'cost
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(4).Value = datatable.Rows(i).Item(4)  'qty

                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try

            Next
            '////////////////////////////////////////////

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '------ tracking report update

        If String.Equals(open_job, "Open Project:") = False Then

            '---------- Lets create a tracking report
            Try
                '-- get id of BOM
                Dim id_bom As String : id_bom = 1
                Dim cmd21 As New MySqlCommand
                cmd21.Parameters.AddWithValue("@mr_name", Me.Text)
                cmd21.CommandText = "SELECT id_bom from Material_Request.mr where mr_name = @mr_name"
                cmd21.Connection = Login.Connection
                Dim reader21 As MySqlDataReader
                reader21 = cmd21.ExecuteReader

                If reader21.HasRows Then
                    While reader21.Read
                        id_bom = reader21(0).ToString
                    End While
                End If

                reader21.Close()



                Dim check_cmd2 As New MySqlCommand
                check_cmd2.Parameters.AddWithValue("@job", open_job)
                check_cmd2.Parameters.AddWithValue("@id_bom", id_bom)
                check_cmd2.CommandText = "delete from Tracking_Reports.my_tracking_reports where job = @job and id_bom = @id_bom"
                check_cmd2.Connection = Login.Connection
                check_cmd2.ExecuteNonQuery()

                '---insert data

                For i = 0 To orders_grid.Rows.Count - 1

                    If IsNumeric(orders_grid.Rows(i).Cells(4).Value.ToString) = True And String.Equals(orders_grid.Rows(i).Cells(4).Value.ToString, "0") = False Then
                        Dim Create_cmd6 As New MySqlCommand
                        Create_cmd6.Parameters.Clear()
                        Create_cmd6.Parameters.AddWithValue("@job", open_job)
                        Create_cmd6.Parameters.AddWithValue("@Part_No", orders_grid.Rows(i).Cells(0).Value.ToString)
                        Create_cmd6.Parameters.AddWithValue("@primary_vendor", orders_grid.Rows(i).Cells(1).Value.ToString)
                        Create_cmd6.Parameters.AddWithValue("@mfg", orders_grid.Rows(i).Cells(2).Value.ToString)
                        Create_cmd6.Parameters.AddWithValue("@cost", orders_grid.Rows(i).Cells(3).Value.ToString)
                        Create_cmd6.Parameters.AddWithValue("@qty_needed", orders_grid.Rows(i).Cells(4).Value.ToString)
                        Create_cmd6.Parameters.AddWithValue("@PO", If(orders_grid.Rows(i).Cells(5).Value Is Nothing, "", orders_grid.Rows(i).Cells(5).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@es_date_of_arrival", If(orders_grid.Rows(i).Cells(6).Value Is Nothing, "", orders_grid.Rows(i).Cells(6).Value.ToString))
                        Create_cmd6.Parameters.AddWithValue("@id_bom", id_bom)

                        Create_cmd6.CommandText = "INSERT INTO Tracking_Reports.my_tracking_reports(job, Part_No, qty_needed, PO, es_date_of_arrival, primary_vendor, mfg, cost, id_bom) VALUES (@job, @Part_No, @qty_needed, @PO, @es_date_of_arrival, @primary_vendor, @mfg, @cost, @id_bom)"
                        Create_cmd6.Connection = Login.Connection
                        Create_cmd6.ExecuteNonQuery()

                    End If

                Next

                MessageBox.Show("Report updated")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Sub Inventory_project()
        '--- fill Inventory parts demand per project
        MO_grid.Rows.Clear()

        For i = 0 To total_grid.Rows.Count - 1
            If total_grid.Rows(i).Cells(14).Value > 0 And (total_grid.Rows(i).Cells(9).Value - total_grid.Rows(i).Cells(10).Value - total_grid.Rows(i).Cells(12).Value) > 0 Then    'if qty - onhand - transit > 0 and min > 0
                MO_grid.Rows.Add(New String() {total_grid.Rows(i).Cells(1).Value, total_grid.Rows(i).Cells(2).Value})
            End If
        Next
    End Sub

    Private Sub UpdatePricesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdatePricesToolStripMenuItem.Click
        '---- update prices of the current MR and latest mr
        Dim latest_mr As String : latest_mr = "xxxx"
        latest_mr = get_last_revision(Me.Text)

        Call Update_prices(Me.Text)
        Call Update_prices(latest_mr)

        MessageBox.Show("Prices updated successfully")
    End Sub

    Function get_last_revision(mr_name As String) As String

        get_last_revision = "not found"
        Dim id_bom As String : id_bom = "9999999999"
        Dim job As String : job = ""

        Try
            '---------- get id_bom --------------
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT id_bom, job from Material_Request.mr where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    id_bom = reader4(0).ToString
                    job = reader4(1).ToString
                End While
            End If

            reader4.Close()

            '----------------------------------
            '-- get mr name ---
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@id_bom", id_bom)
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                While reader41.Read
                    get_last_revision = reader41(0).ToString

                End While
            End If

            reader41.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Function

    Sub Update_prices(mr_name As String)
        '-- update the prices of a mr
        Try
            For i = 0 To total_grid.Rows.Count - 1

                If String.IsNullOrEmpty(total_grid.Rows(i).Cells(6).Value) = False Then
                    If IsNumeric(total_grid.Rows(i).Cells(6).Value) = True And IsNumeric(total_grid.Rows(i).Cells(0).Value) = True Then

                        Dim new_sub As Double = total_grid.Rows(i).Cells(6).Value * total_grid.Rows(i).Cells(0).Value
                        Dim Create_cmd As New MySqlCommand
                        Create_cmd.Parameters.Clear()
                        Create_cmd.Parameters.AddWithValue("@mr_name", mr_name)
                        Create_cmd.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(1).Value)
                        Create_cmd.Parameters.AddWithValue("@Price", total_grid.Rows(i).Cells(6).Value)
                        Create_cmd.Parameters.AddWithValue("@subtotal", new_sub.ToString)



                        Create_cmd.CommandText = "UPDATE Material_Request.mr_data  SET Price = @Price, subtotal = @subtotal where Part_No = @Part_No and mr_name = @mr_name"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                        '--- latest
                        Dim Create_cmd2 As New MySqlCommand
                        Create_cmd2.Parameters.Clear()
                        Create_cmd2.Parameters.AddWithValue("@mr_name", get_last_revision(mr_name))
                        Create_cmd2.Parameters.AddWithValue("@Part_No", total_grid.Rows(i).Cells(1).Value)
                        Create_cmd2.Parameters.AddWithValue("@Price", total_grid.Rows(i).Cells(6).Value)
                        Create_cmd2.Parameters.AddWithValue("@subtotal", new_sub.ToString)



                        Create_cmd2.CommandText = "UPDATE Material_Request.mr_data  SET Price = @Price, subtotal = @subtotal where Part_No = @Part_No and mr_name = @mr_name"
                        Create_cmd2.Connection = Login.Connection
                        Create_cmd2.ExecuteNonQuery()


                    End If
                End If

            Next

            MessageBox.Show("Price Update Done!")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If String.Equals("My Material Requests", Me.Text) = False And Me.Visible = True And toggle = True Then

            Dim mr_name As String : mr_name = Me.Text

            If TabControl1.SelectedTab Is TabPage3 Then

                '---- MBOM fullfill search

                total_grid.Rows.Clear()

                '---- datatable store MR without assemblies --------
                Dim dimen_table = New DataTable
                dimen_table.Columns.Add("Project Qty", GetType(String))
                dimen_table.Columns.Add("Part_No", GetType(String))
                dimen_table.Columns.Add("Description", GetType(String))
                dimen_table.Columns.Add("Manufacturer", GetType(String))
                dimen_table.Columns.Add("MFG Type", GetType(String))
                dimen_table.Columns.Add("Vendor", GetType(String))
                dimen_table.Columns.Add("Price", GetType(String))
                dimen_table.Columns.Add("subtotal", GetType(String))
                dimen_table.Columns.Add("qty_fullfilled", GetType(String))
                dimen_table.Columns.Add("need_by_date", GetType(String))
                dimen_table.Columns.Add("Notes", GetType(String))
                dimen_table.Columns.Add("Assembly_name", GetType(String))

                Try

                    '-------------------------------------
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@mr_name", Me.Text)
                    cmd4.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")
                    cmd4.CommandText = "SELECT Qty, Part_No, Description, Manufacturer, mfg_type, Vendor, Price, subtotal,  qty_fullfilled, need_by_date, Notes, Assembly_name from Material_Request.mr_data where mr_name = @mr_name and Part_No LIKE @search"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader


                    If reader4.HasRows Then
                        Dim i As Integer : i = 0
                        While reader4.Read
                            dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, (If(IsNumeric(reader4(6)), reader4(6), 0) * If(IsNumeric(reader4(0)), reader4(0), 0)), reader4(8).ToString, reader4(9).ToString, reader4(10).ToString, reader4(11).ToString)
                        End While

                    End If

                    reader4.Close()

                    '------- now copy dimen_table_a to inventroy grid
                    For i = 0 To dimen_table.Rows.Count - 1
                        total_grid.Rows.Add(New String() {})
                        total_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString 'project qty
                        total_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString 'Part_No
                        total_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString  'Description
                        total_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString 'Manufacturer
                        total_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString   'MFG Type
                        total_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString  'vendor
                        total_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString  'Price
                        total_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString  'subtotal
                        total_grid.Rows(i).Cells(8).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)   'fulfilled
                        total_grid.Rows(i).Cells(9).Value = 0  'needed
                        total_grid.Rows(i).Cells(10).Value = 0 'on hand
                        total_grid.Rows(i).Cells(11).Value = 0 '+/-
                        total_grid.Rows(i).Cells(12).Value = 0 'in transit
                        total_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(9).ToString 'need date
                        total_grid.Rows(i).Cells(14).Value = 0 'min
                        total_grid.Rows(i).Cells(15).Value = 0 'max
                        total_grid.Rows(i).Cells(16).Value = dimen_table.Rows(i).Item(10).ToString 'notes
                        total_grid.Rows(i).Cells(17).Value = dimen_table.Rows(i).Item(11).ToString 'assembly
                    Next


                    '////////////////////  ---------- fill current inventory values -----------------------

                    For i = 0 To total_grid.Rows.Count - 1

                        Dim cmd5 As New MySqlCommand
                        cmd5.Parameters.Clear()
                        cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                        cmd5.CommandText = "SELECT current_qty,  Qty_in_order, min_qty, max_qty from inventory.inventory_qty where part_name = @part_name"
                        cmd5.Connection = Login.Connection
                        Dim reader5 As MySqlDataReader
                        reader5 = cmd5.ExecuteReader

                        If reader5.HasRows Then
                            While reader5.Read
                                total_grid.Rows(i).Cells(10).Value = reader5(0).ToString
                                total_grid.Rows(i).Cells(12).Value = reader5(1).ToString
                                total_grid.Rows(i).Cells(14).Value = reader5(2).ToString 'min
                                total_grid.Rows(i).Cells(15).Value = reader5(3).ToString 'max
                            End While
                        End If
                        reader5.Close()

                    Next
                    '----------------------------------------------------------------             

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try


            End If

        End If
    End Sub



    Sub Module_show(name As String, job As String)
        '------------ when click show data in datagridview ---------------
        Try

            TabControl1.TabPages.Insert(4, TabPage6)
            TabControl1.TabPages.Remove(TabPage5)
            TabControl1.SelectedTab = TabPage6
            Dim mr_name As String : mr_name = ""
            fullfill_grid.Rows.Clear()

            If String.Equals(name, "Field BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "Field")
                cmd5.Parameters.AddWithValue("@job", job)

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

                mr_name = get_last_revision(mr_name)

                If String.Equals(mr_name, "not found") = False Then
                    Call Open_recent(mr_name, job)
                End If

                bom_open.Text = mr_name
                qty_panel.Text = "(1)"
                temp_qty_panel = 1

            ElseIf String.Equals(name, "Spare Parts BOM") = True Then

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@BOM_type", "Spare_Parts")
                cmd5.Parameters.AddWithValue("@job", job)

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
                mr_name = get_last_revision(mr_name)


                If String.Equals(mr_name, "not found") = False Then
                    Call Open_recent(mr_name, job)
                End If

                bom_open.Text = mr_name
                qty_panel.Text = "(1)"
                temp_qty_panel = 1

            Else

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@job", job)
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
                mr_name = get_last_revision(mr_name)

                '--- get qty panel ---
                Dim cmd419 As New MySqlCommand
                cmd419.Parameters.Clear()
                cmd419.Parameters.AddWithValue("@mr_name", mr_name)
                cmd419.CommandText = "SELECT Panel_qty, BOM_type from Material_Request.mr where mr_name = @mr_name"
                cmd419.Connection = Login.Connection
                Dim reader419 As MySqlDataReader
                reader419 = cmd419.ExecuteReader

                If reader419.HasRows Then
                    While reader419.Read

                        qty_panel.Text = "(" & If(reader419(0) Is DBNull.Value, 1, reader419(0)) & ")"

                        If String.Equals(reader419(1).ToString, "Assembly") = False Then
                            temp_qty_panel = If(reader419(0) Is DBNull.Value, 1, reader419(0))
                        Else
                            temp_qty_panel = 1
                        End If


                    End While
                End If

                reader419.Close()
                '----------------------------------


                If String.Equals(mr_name, "not found") = False Then
                    Call Open_recent(mr_name, job)
                End If

                bom_open.Text = mr_name
                p_name_p.Text = name


            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Open_recent(name As String, job As String)

        'this sub open  the data in the inventory request datagridview

        Dim dimen_table = New DataTable
        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Description", GetType(String))
        dimen_table.Columns.Add("Manufacturer", GetType(String))
        dimen_table.Columns.Add("Assembly_name", GetType(String))
        dimen_table.Columns.Add("Vendor", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("mfg_type", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("qty_fullfilled", GetType(String))
        dimen_table.Columns.Add("qty_needed", GetType(String))
        dimen_table.Columns.Add("Notes", GetType(String))



        Try

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", name)
            cmd4.CommandText = "SELECT Part_No, Description, Manufacturer, Assembly_name, Vendor, Price, mfg_type, Qty, qty_fullfilled, qty_needed , Notes from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, reader4(4).ToString, reader4(5).ToString, reader4(6).ToString, reader4(7).ToString * temp_qty_panel, reader4(8).ToString, reader4(9).ToString, reader4(10).ToString)
                End While
            End If

            reader4.Close()


            '------- now copy dimen_table_a to inventroy grid
            For i = 0 To dimen_table.Rows.Count - 1
                fullfill_grid.Rows.Add(New String() {})
                fullfill_grid.Rows(i).Cells(0).Value = dimen_table.Rows(i).Item(0).ToString
                fullfill_grid.Rows(i).Cells(1).Value = dimen_table.Rows(i).Item(1).ToString
                fullfill_grid.Rows(i).Cells(2).Value = dimen_table.Rows(i).Item(2).ToString
                fullfill_grid.Rows(i).Cells(3).Value = dimen_table.Rows(i).Item(3).ToString
                fullfill_grid.Rows(i).Cells(4).Value = dimen_table.Rows(i).Item(4).ToString
                fullfill_grid.Rows(i).Cells(5).Value = dimen_table.Rows(i).Item(5).ToString
                fullfill_grid.Rows(i).Cells(6).Value = dimen_table.Rows(i).Item(6).ToString
                fullfill_grid.Rows(i).Cells(7).Value = dimen_table.Rows(i).Item(7).ToString
                fullfill_grid.Rows(i).Cells(8).Value = 0
                fullfill_grid.Rows(i).Cells(9).Value = 0
                fullfill_grid.Rows(i).Cells(10).Value = If(IsNumeric(dimen_table.Rows(i).Item(8).ToString) = True, dimen_table.Rows(i).Item(8).ToString, 0)
                fullfill_grid.Rows(i).Cells(11).Value = If(IsNumeric(dimen_table.Rows(i).Item(9).ToString) = True, dimen_table.Rows(i).Item(9).ToString, 0)
                fullfill_grid.Rows(i).Cells(13).Value = dimen_table.Rows(i).Item(10).ToString

            Next


            '////////////////////  ---------- fill current inventory values -----------------------

            For i = 0 To fullfill_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty , Qty_in_order from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        fullfill_grid.Rows(i).Cells(8).Value = reader5(0).ToString
                        fullfill_grid.Rows(i).Cells(12).Value = reader5(1).ToString
                    End While
                End If

                reader5.Close()

            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub F_BOM_Click(sender As Object, e As EventArgs) Handles F_BOM.Click

        If String.Equals(open_job, "") = False Then
            Module_show("Field BOM", open_job)
        End If
    End Sub

    Private Sub SP_BOM_Click(sender As Object, e As EventArgs) Handles SP_BOM.Click
        If String.Equals(open_job, "") = False Then
            Module_show("Spare Parts BOM", open_job)
        End If
    End Sub


    '--go back to BOM Map
    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        TabControl1.TabPages.Insert(4, TabPage5)
        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.SelectedTab = TabPage5
    End Sub

    Private Sub Label16_MouseEnter(sender As Object, e As EventArgs) Handles Label16.MouseEnter
        Label16.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label16_MouseLeave(sender As Object, e As EventArgs) Handles Label16.MouseLeave
        Label16.ForeColor = Color.DarkCyan
    End Sub

    Private Sub fullfill_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles fullfill_grid.CellValueChanged
        If go_ahead = True Then

            For Each row As DataGridViewRow In fullfill_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(7).Value) = True And IsNumeric(row.Cells(9).Value)) Then

                    If CType(row.Cells(8).Value, Double) < CType(row.Cells(9).Value, Double) Then
                        row.Cells(9).Value = 0
                    End If


                    If (IsNumeric(row.Cells(7).Value) = True And IsNumeric(row.Cells(10).Value)) Then
                        row.Cells(11).Value = CType(row.Cells(7).Value, Double) - CType(row.Cells(10).Value, Double)
                    Else
                        row.Cells(11).Value = If(IsNumeric(row.Cells(7).Value) = False, 0, row.Cells(7).Value)
                    End If


                    If CType(row.Cells(11).Value, Double) <> 0 Then
                        row.Cells(11).Style.BackColor = Color.Firebrick
                    Else
                        row.Cells(11).Style.BackColor = Color.Gray
                    End If

                End If
            Next
        End If
    End Sub

    Private Sub DisableAllocationSecurityToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisableAllocationSecurityToolStripMenuItem.Click
        If DisableAllocationSecurityToolStripMenuItem.Checked = True Then
            DisableAllocationSecurityToolStripMenuItem.Checked = False
        Else
            DisableAllocationSecurityToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub NotifyProgressToProcurementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NotifyProgressToProcurementToolStripMenuItem.Click
        If String.Equals(Me.Text, "My Material Requests") = False Then

            Try

                '---------- calculate percentage ---------
                Dim complete_mr As Double : complete_mr = 0
                Dim total_qty As Double : total_qty = 0
                Dim fullf As Double : fullf = 0


                Dim cmd3 As New MySqlCommand
                cmd3.Parameters.AddWithValue("@mr_name", Me.Text)
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
                cmdx.Parameters.AddWithValue("@mr_name", Me.Text)
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

                If total_qty > 0 Then
                    complete_mr = Math.Round((fullf / total_qty) * 100)
                End If
                '---------------------------------------

                '---------- send notification ------------

                If enable_mess = True Then

                    Dim email_m As String
                    email_m = "Material Request Allocation for Project " & open_job & "  has been updated" & vbCrLf & vbCrLf _
                                & "Sent by: " & current_user & vbCrLf _
                                & "Sent Date: " & Now & vbCrLf _
                                & "Project: " & open_job & vbCrLf _
                                & "MR Filename: " & Me.Text & vbCrLf _
                                & "Material Request Status: " & complete_mr & "% fulfilled" & vbCrLf


                    Dim message_title As String : message_title = "Material Request Allocation for Project " & open_job & " updated"

                    ' Call Sent_mail.Sent_multiple_emails("Procurement", message_title, email_m)
                    Call Sent_mail.Sent_multiple_emails("General Management", message_title, email_m)
                    'Call Sent_mail.Sent_multiple_emails("Inventory", message_title, email_m)

                End If

                MessageBox.Show("Notification sent!")


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Sub Color_Module()

        '--- This sub color the blocks of the BOM Map
        Dim MBOM_p As Double : MBOM_p = percentage_full("Master BOM")
        Dim Field_p As Double : Field_p = percentage_full("Field BOM")
        Dim SP_p As Double : SP_p = percentage_full("Spare Parts BOM")

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
        ElseIf Field_p = 0 And in_progress_bom_color("Field BOM", job_label.text) = False Then
            F_BOM.BackColor = Color.DarkRed
        Else
            F_BOM.BackColor = Color.DarkGoldenrod
        End If

        '----------
        If SP_p = 100 Then
            SP_BOM.BackColor = Color.Teal
        ElseIf SP_p = 0 And in_progress_bom_color("Spare Parts BOM", job_label.text) = False Then
            SP_BOM.BackColor = Color.DarkRed
        Else
            SP_BOM.BackColor = Color.DarkGoldenrod
        End If



        For i = 0 To plabels.Count - 1
            Dim P_p As Double : P_p = percentage_full(plabels(i).Text)

            If P_p = 100 Then
                plabels(i).BackColor = Color.Teal
            ElseIf P_p = 0 And in_progress_bom_color(plabels(i).Text, job_label.text) = False Then
                plabels(i).BackColor = Color.DarkRed
            Else
                plabels(i).BackColor = Color.DarkGoldenrod
            End If
        Next



        For i = 0 To alabels.Count - 1
            Dim AP_p As Double : AP_p = percentage_full(alabels(i).Text)

            If AP_p = 100 Then
                alabels(i).BackColor = Color.Teal
            ElseIf AP_p = 0 Then
                alabels(i).BackColor = Color.DarkRed
            Else
                alabels(i).BackColor = Color.DarkGoldenrod
            End If
        Next

    End Sub

    Function percentage_full(name As String) As Double

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

                mr_name = get_last_revision(mr_name)

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
                mr_name = get_last_revision(mr_name)


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
                mr_name = get_last_revision(mr_name)



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

                mr_name = get_last_revision(mr_name)

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

    Private Sub DisableAllocationSecurityToolStripMenuItem_CheckedChanged(sender As Object, e As EventArgs) Handles DisableAllocationSecurityToolStripMenuItem.CheckedChanged
        If DisableAllocationSecurityToolStripMenuItem.Checked = True Then
            go_ahead = False
        Else
            go_ahead = True
        End If
    End Sub

    Private Sub FullfillmentTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FullfillmentTableToolStripMenuItem.Click
        Call Export_fullfill()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        '--------- Fullfillment button --------
        '--------------- Start updating the Mr table, the current inventory qty and generate a material report -------------------

        If String.Equals(Me.Text, "My Material Requests") = False Then

            wait_la.Visible = True
            Application.DoEvents()

            '--------------- find latest version --------------
            Dim latest_v As String : latest_v = Proc_Material_R.get_last_revision(bom_open.Text)
            '---------------------------------------------------
            Call Update_mr_fullfill(latest_v)   'update latest revision mr
            '   Call Update_mr_fullfill(bom_open.Text)    'update current mr which can be the same as above
            Call update_mb_fullfill(get_last_revision(Me.Text)) 'update latest MB

            Call Update_current_inv()  'update inventory current qtys
            Call Inventory_manage.General_inv_cal()  'fill material order form

            go_ahead = False

            '--- reload entire table except current inv----
            fullfill_grid.Rows.Clear()

            go_ahead = True

            wait_la.Visible = False
            MessageBox.Show("Material Request updated successfully!")

            Call Open_recent(bom_open.Text, open_job)
            Call Color_Module()
        End If

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

    Sub Update_mr_fullfill(mr_name As String)

        Try

            For i = 0 To fullfill_grid.Rows.Count - 1

                If fullfill_grid.Rows(i).Cells(9).Value <> 0 Then

                    Dim qty_f As Double = 0


                    If IsNumeric(fullfill_grid.Rows(i).Cells(10).Value) = True Then
                        qty_f = CType(fullfill_grid.Rows(i).Cells(10).Value, Double)
                    End If

                    If IsNumeric(fullfill_grid.Rows(i).Cells(9).Value) = True Then
                        qty_f = qty_f + CType(fullfill_grid.Rows(i).Cells(9).Value, Double)
                    End If



                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                    cmd5.Parameters.AddWithValue("@mr_name", mr_name)

                    cmd5.CommandText = "UPDATE Material_Request.mr_data SET qty_fullfilled = @qty_fullfilled where mr_name = @mr_name and Part_No = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub update_mb_fullfill(mr_name As String)
        Try

            For i = 0 To fullfill_grid.Rows.Count - 1

                If fullfill_grid.Rows(i).Cells(9).Value <> 0 Then

                    Dim qty_f As Double = 0

                    If IsNumeric(fullfill_grid.Rows(i).Cells(9).Value) = True Then
                        qty_f = qty_f + CType(fullfill_grid.Rows(i).Cells(9).Value, Double)
                    End If

                    '--- add qty_f to fulfill in MB
                    Dim real_full As Double : real_full = 0
                    Dim cmd3 As New MySqlCommand
                    cmd3.Parameters.AddWithValue("@mr_name", mr_name)
                    cmd3.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd3.CommandText = "select qty_fullfilled from Material_Request.mr_data where mr_name = @mr_name and Part_No = @part_name and qty_fullfilled is not null;"
                    cmd3.Connection = Login.Connection
                    Dim reader3 As MySqlDataReader
                    reader3 = cmd3.ExecuteReader

                    If reader3.HasRows Then
                        While reader3.Read
                            real_full = reader3(0)
                        End While
                    End If

                    reader3.Close()

                    qty_f = qty_f + real_full


                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                    cmd5.Parameters.AddWithValue("@mr_name", mr_name)

                    cmd5.CommandText = "UPDATE Material_Request.mr_data SET qty_fullfilled = @qty_fullfilled where mr_name = @mr_name and Part_No = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()
                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Update_current_inv()

        Try
            For i = 0 To fullfill_grid.Rows.Count - 1

                If fullfill_grid.Rows(i).Cells(9).Value <> 0 Then

                    Dim current_q As Double : current_q = 0
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        While reader4.Read
                            current_q = reader4(0).ToString
                        End While
                    End If

                    reader4.Close()

                    current_q = current_q - If(IsNumeric(fullfill_grid.Rows(i).Cells(9).Value), fullfill_grid.Rows(i).Cells(9).Value, 0)
                    current_q = If(current_q < 0, 0, current_q)


                    '---- update the value
                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", fullfill_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@current_qty", current_q)

                    cmd5.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Procurement_Overview_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub

    Private Sub AssemblyBOMOverviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssemblyBOMOverviewToolStripMenuItem.Click


        If String.Equals(Me.Text, "My Material Requests") = False Then

            If ASM_mode = True Then
                ASM_overview.Visible = True
            Else
                MessageBox.Show("Not a Assembly BOM")
            End If
        Else
            MessageBox.Show("Please Open a BOM")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If ASM_mode = True Then

            wait_la2.Visible = True
            Application.DoEvents()

            Call Update_mr_fullfill_asm(Me.Text)
            Call Inventory_manage.General_inv_cal()

            total_grid.Rows.Clear()

            wait_la2.Visible = False
            MessageBox.Show("Assembly BOM updated successfully!")

            Call open_my_job(Me.Text)

        Else
            MessageBox.Show("Not an Assembly Mode")
        End If
    End Sub

    Sub Update_mr_fullfill_asm(mr_name As String)

        Try
            For i = 0 To total_grid.Rows.Count - 1

                Dim update_v As Boolean : update_v = False

                '------- get inventory ---------
                Dim current_q As Double : current_q = 0
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        current_q = reader4(0).ToString
                    End While
                End If

                reader4.Close()
                '----------------------------------

                Dim qty_f As Double = 0

                If IsNumeric(total_grid.Rows(i).Cells(8).Value) = True Then
                    qty_f = CType(total_grid.Rows(i).Cells(8).Value, Double)
                End If

                If IsNumeric(total_grid.Rows(i).Cells(11).Value) = True Then

                    If total_grid.Rows(i).Cells(11).Value <= current_q Then
                        qty_f = qty_f + CType(total_grid.Rows(i).Cells(11).Value, Double)
                        update_v = True
                    End If

                End If

                If update_v = True Then

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                    cmd5.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                    cmd5.Parameters.AddWithValue("@mr_name", mr_name)
                    cmd5.Parameters.AddWithValue("@asm", total_grid.Rows(i).Cells(17).Value)

                    cmd5.CommandText = "UPDATE Material_Request.mr_data SET qty_fullfilled = @qty_fullfilled where mr_name = @mr_name and Part_No = @part_name and Assembly_name = @asm"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                    Dim qty_ad As Double : qty_ad = 0
                    qty_ad = If(IsNumeric(total_grid.Rows(i).Cells(11).Value), total_grid.Rows(i).Cells(11).Value, 0)

                    If current_q >= qty_ad Then

                        current_q = current_q - qty_ad
                        current_q = If(current_q < 0, 0, current_q)

                        '---- update the value
                        Dim cmd51 As New MySqlCommand
                        cmd51.Parameters.Clear()
                        cmd51.Parameters.AddWithValue("@part_name", total_grid.Rows(i).Cells(1).Value)
                        cmd51.Parameters.AddWithValue("@current_qty", current_q)

                        cmd51.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                        cmd51.Connection = Login.Connection
                        cmd51.ExecuteNonQuery()

                    End If

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub UpdateOnHandQtyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateOnHandQtyToolStripMenuItem.Click

        If TabControl1.SelectedTab Is TabPage6 Then

            If fullfill_grid.Rows.Count > 0 Then

                If (fullfill_grid.CurrentCell.ColumnIndex = 0) Then

                    Dim component As String : component = ""
                    component = fullfill_grid.CurrentCell.Value.ToString

                    Fast_Inventory.Text = component
                    Fast_Inventory.ShowDialog()

                Else
                    MessageBox.Show("Please select a Part")

                End If

            End If

        ElseIf TabControl1.SelectedTab Is TabPage10 Then
            If asm_grid.Rows.Count > 0 Then

                If (asm_grid.CurrentCell.ColumnIndex = 0) Then

                    Dim component As String : component = ""
                    component = asm_grid.CurrentCell.Value.ToString

                    Fast_Inventory.Text = component
                    Fast_Inventory.ShowDialog()

                Else
                    MessageBox.Show("Please select a Part")

                End If

            End If
        End If

    End Sub

    Sub MPL_open(job As String)

        mpl_grid.Rows.Clear()
        mpl_compare.Rows.Clear()
        mpl_box1.Items.Clear()
        mpl_box2.Items.Clear()

        Dim n_r As Integer : n_r = 0

        Try

            '--get latest revision --
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "SELECT distinct n_r from Master_Packing_List.packing_l where job = @job order by n_r desc limit 1"
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
            cmd4.CommandText = "SELECT qty, part_name, part_desc, need_date from Master_Packing_List.packing_l where job = @job and n_r = @n_r"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    mpl_grid.Rows.Add(New String() {})
                    mpl_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    mpl_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    mpl_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    mpl_grid.Rows(i).Cells(3).Value = reader4(3).ToString

                    i = i + 1
                End While

            End If

            reader4.Close()


            '--fill dropboxes --------
            Dim cmdc As New MySqlCommand
            cmdc.Parameters.AddWithValue("@job", job)
            cmdc.CommandText = "SELECT distinct mpl_name from Master_Packing_List.packing_l where job = @job"
            cmdc.Connection = Login.Connection
            Dim readerc As MySqlDataReader
            readerc = cmdc.ExecuteReader

            If readerc.HasRows Then
                While readerc.Read
                    mpl_box1.Items.Add(readerc(0).ToString)
                    mpl_box2.Items.Add(readerc(0).ToString)
                End While
            End If

            readerc.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'export  MPL to excel file


        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then


                Cursor.Current = Cursors.WaitCursor

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
                For i As Integer = 0 To mpl_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = mpl_grid.Columns(i).HeaderText
                    For j As Integer = 0 To mpl_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = mpl_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\MPL.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                Cursor.Current = Cursors.Default
                MessageBox.Show("Master Packing List exported successfully!")
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        '---- Compare two MPL Request
        If Not mpl_box1.SelectedItem Is Nothing And Not mpl_box2.SelectedItem Is Nothing Then

            mpl_compare.Rows.Clear()

            Dim total_mr As New Dictionary(Of String, String)
            '-- store mr1 parts
            Dim dimen_table1 = New DataTable
            dimen_table1.Columns.Add("Part_No", GetType(String))
            dimen_table1.Columns.Add("Description", GetType(String))

            '-- store mr2 parts
            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("Description", GetType(String))

            Try
                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@mpl_name", mpl_box1.Text)
                cmd4.CommandText = "SELECT part_name, part_desc from Master_Packing_List.packing_l where mpl_name = @mpl_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then
                    While reader4.Read
                        dimen_table1.Rows.Add(reader4(0).ToString, reader4(1).ToString)
                    End While
                End If

                reader4.Close()

                '--- mr2
                Dim cmd41 As New MySqlCommand
                cmd41.Parameters.AddWithValue("@mpl_name", mpl_box2.Text)
                cmd41.CommandText = "SELECT part_name, part_desc from Master_Packing_List.packing_l where mpl_name = @mpl_name"
                cmd41.Connection = Login.Connection
                Dim reader41 As MySqlDataReader
                reader41 = cmd41.ExecuteReader

                If reader41.HasRows Then
                    While reader41.Read
                        dimen_table2.Rows.Add(reader41(0).ToString, reader41(1).ToString)
                    End While
                End If

                reader41.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            For i = 0 To dimen_table1.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table1.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table1.Rows(i).Item(0), dimen_table1.Rows(i).Item(1))
                End If
            Next

            For i = 0 To dimen_table2.Rows.Count - 1
                If total_mr.ContainsKey(dimen_table2.Rows(i).Item(0)) = False Then
                    total_mr.Add(dimen_table2.Rows(i).Item(0), dimen_table2.Rows(i).Item(1))
                End If
            Next

            '--- add to comp_Grid
            For Each kvp As KeyValuePair(Of String, String) In total_mr.ToArray
                mpl_compare.Rows.Add(New String() {kvp.Key, total_mr(kvp.Key)})
            Next

            For i = 0 To mpl_compare.Rows.Count - 1
                Try
                    Dim cmd41 As New MySqlCommand
                    cmd41.Parameters.Clear()
                    cmd41.Parameters.AddWithValue("@mpl_name", mpl_box1.Text)
                    cmd41.Parameters.AddWithValue("@part", mpl_compare.Rows(i).Cells(0).Value)
                    cmd41.CommandText = "SELECT qty from  Master_Packing_List.packing_l where mpl_name = @mpl_name and part_name = @part"
                    cmd41.Connection = Login.Connection
                    Dim reader41 As MySqlDataReader
                    reader41 = cmd41.ExecuteReader

                    If reader41.HasRows Then
                        While reader41.Read
                            mpl_compare.Rows(i).Cells(2).Value = reader41(0).ToString
                        End While
                    End If

                    reader41.Close()

                    Dim cmd42 As New MySqlCommand
                    cmd42.Parameters.Clear()
                    cmd42.Parameters.AddWithValue("@mpl_name", mpl_box2.Text)
                    cmd42.Parameters.AddWithValue("@part", mpl_compare.Rows(i).Cells(0).Value)
                    cmd42.CommandText = "SELECT qty from  Master_Packing_List.packing_l where mpl_name = @mpl_name and part_name = @part"
                    cmd42.Connection = Login.Connection
                    Dim reader42 As MySqlDataReader
                    reader42 = cmd42.ExecuteReader

                    If reader42.HasRows Then
                        While reader42.Read
                            mpl_compare.Rows(i).Cells(3).Value = reader42(0).ToString
                        End While
                    End If

                    reader42.Close()

                    mpl_compare.Rows(i).Cells(4).Value = If(IsNumeric(mpl_compare.Rows(i).Cells(3).Value), mpl_compare.Rows(i).Cells(3).Value, 0) - If(IsNumeric(mpl_compare.Rows(i).Cells(2).Value), mpl_compare.Rows(i).Cells(2).Value, 0)

                    If mpl_compare.Rows(i).Cells(4).Value <> 0 Then
                        mpl_compare.Rows(i).Cells(4).Style.BackColor = Color.CadetBlue
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            Next

        End If


    End Sub

    Private Sub total_grid_DoubleClick(sender As Object, e As EventArgs) Handles total_grid.DoubleClick

        Call Alter_cable(total_grid)

    End Sub


    '-------- show alternative cables --------
    Sub Alter_cable(mygrid As DataGridView)

        If TabControl1.SelectedTab Is TabPage3 Then
            '-- Master bom first tab

            If mygrid.Rows.Count > 0 Then

                If String.IsNullOrEmpty(mygrid.CurrentCell.Value) = False Then

                    Dim index_k = mygrid.CurrentCell.RowIndex

                    If String.Equals(mygrid.Rows(index_k).Cells(1).Value.ToString, "") = False And String.Equals(mygrid.Rows(index_k).Cells(1).Value.ToString, " ") = False Then

                        Dim component As String : component = mygrid.Rows(index_k).Cells(1).Value.ToString

                        Cables_ADA.Text = "Alternatives for " & component

                        Call fill_alt(component)
                        Cables_ADA.ShowDialog()

                    End If
                End If
            End If

        ElseIf TabControl1.SelectedTab Is TabPage6 Then


            If mygrid.Rows.Count > 0 Then

                If String.IsNullOrEmpty(mygrid.CurrentCell.Value) = False Then

                    Dim index_k = mygrid.CurrentCell.RowIndex

                    If String.Equals(mygrid.Rows(index_k).Cells(0).Value.ToString, "") = False And String.Equals(mygrid.Rows(index_k).Cells(0).Value.ToString, " ") = False Then

                        Dim component As String : component = mygrid.Rows(index_k).Cells(0).Value.ToString

                        Cables_ADA.Text = "Alternatives for " & component

                        Call fill_alt(component)
                        Cables_ADA.ShowDialog()

                    End If
                End If
            End If


        End If


    End Sub

    Private Sub fullfill_grid_DoubleClick(sender As Object, e As EventArgs) Handles fullfill_grid.DoubleClick

        '  Call Alter_cable(fullfill_grid)

        Dim index_k = fullfill_grid.CurrentCell.RowIndex


        '--- may open an assembly from here

        Label20.Text = fullfill_grid.Rows(index_k).Cells(0).Value
        mr_label.Text = p_name_p.Text

        Dim qty_as As Double : qty_as = 0
        'qty_as = fullfill_grid.Rows(index_k).Cells(7).Value - fullfill_grid.Rows(index_k).Cells(10).Value

        'If qty_as <= fullfill_grid.Rows(index_k).Cells(8).Value Then
        '    qty_as = 0
        'Else
        '    qty_as = qty_as - fullfill_grid.Rows(index_k).Cells(8).Value
        'End If

        qty_as = fullfill_grid.Rows(index_k).Cells(7).Value

        Call open_assem(fullfill_grid.Rows(index_k).Cells(0).Value, qty_as)  'open assembly

        Label19.Text = "(" & qty_as & ")"



        TabControl1.TabPages.Insert(4, TabPage10)
        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.SelectedTab = TabPage10



    End Sub

    Sub fill_alt(component As String)
        '-------- fill table -----------
        Try

            Cables_ADA.alt_grid.Rows.Clear()
            Dim ada_n As String : ada_n = "unknown"
            Dim cmd42 As New MySqlCommand
            cmd42.Parameters.Clear()
            cmd42.Parameters.AddWithValue("@part_name", component)
            cmd42.CommandText = "SELECT Legacy_ADA_Number from parts_table where part_name = @part_name"
            cmd42.Connection = Login.Connection
            Dim reader42 As MySqlDataReader
            reader42 = cmd42.ExecuteReader

            If reader42.HasRows Then
                While reader42.Read
                    If Not IsDBNull(reader42(0)) Then
                        ada_n = reader42(0)
                    End If
                End While
            End If

            reader42.Close()

            If String.IsNullOrEmpty(ada_n) = True Then
                ada_n = "unknown"
            End If

            '---------- display all the alternatives ---------
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@ADA", ada_n)
            cmd4.Parameters.AddWithValue("@part", component)
            cmd4.CommandText = "SELECT p1.Part_Name, p1.Manufacturer, inventory.inventory_qty.current_qty from parts_table as p1 INNER JOIN inventory.inventory_qty ON p1.Part_Name = inventory.inventory_qty.part_name where p1.Legacy_ADA_Number = @ADA and p1.Part_Name not like @part"


            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    Cables_ADA.alt_grid.Rows.Add(New String() {})
                    Cables_ADA.alt_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    Cables_ADA.alt_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    Cables_ADA.alt_grid.Rows(i).Cells(2).Value = reader4(2).ToString
                    i = i + 1
                End While

            End If

            reader4.Close()
            '-------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '--------------------------------
    End Sub

    Private Sub ShowShippingAddressToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowShippingAddressToolStripMenuItem.Click

        Shipping_map.ShowDialog()
    End Sub

    Private Sub PackingListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PackingListToolStripMenuItem.Click
        MPL_pc.Visible = True
    End Sub

    Private Sub MoveToPackiNgSlipToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveToPackiNgSlipToolStripMenuItem.Click

        If TabControl1.SelectedTab Is TabPage2 Then
            If MPL_pc.Visible = True Then

                'Dim index As Integer : index = 0

                'For i = 0 To MPL_pc.PR_grid.Rows.Count - 1
                '    If String.IsNullOrEmpty(MPL_pc.PR_grid.Rows(i).Cells(0).Value) Then
                '        index = i
                '        Exit For
                '    End If
                'Next

                For i = 0 To compare_grid.Rows.Count - 1
                    MPL_pc.PR_grid.Rows.Add(New String() {compare_grid.Rows(i).Cells(6).Value, compare_grid.Rows(i).Cells(0).Value, compare_grid.Rows(i).Cells(1).Value})
                Next
            End If
        End If

    End Sub

    Private Sub job_label_Click(sender As Object, e As EventArgs) Handles job_label.Click
        Job_description.ShowDialog()
    End Sub

    Private Sub RejectMaterialRequestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RejectMaterialRequestToolStripMenuItem.Click
        Reject_BOM.ShowDialog()
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        TabControl1.TabPages.Insert(4, TabPage6)
        TabControl1.TabPages.Remove(TabPage10)
        TabControl1.SelectedTab = TabPage6
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
            cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, p1.Manufacturer,  p1.Primary_Vendor, p1.MFG_type, adv.Qty from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = @adv"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    asm_grid.Rows.Add()  'add a new row
                    asm_grid.Rows(i).Cells(0).Value = readerv(0).ToString  'part
                    asm_grid.Rows(i).Cells(1).Value = readerv(1).ToString  'desc
                    asm_grid.Rows(i).Cells(2).Value = readerv(2).ToString   'mfg
                    asm_grid.Rows(i).Cells(3).Value = readerv(3).ToString  'vendor
                    asm_grid.Rows(i).Cells(4).Value = 0  'price
                    asm_grid.Rows(i).Cells(5).Value = readerv(4).ToString  'mfg type
                    asm_grid.Rows(i).Cells(6).Value = readerv(5).ToString * qty   'qty demanded
                    asm_grid.Rows(i).Cells(9).Value = 0  'fulfilled
                    asm_grid.Rows(i).Cells(10).Value = 0
                    i = i + 1
                End While
            End If

            readerv.Close()


            For j = 0 To asm_grid.Rows.Count - 1
                asm_grid.Rows(j).Cells(4).Value = Form1.Get_Latest_Cost(Login.Connection, asm_grid.Rows(j).Cells(0).Value, asm_grid.Rows(j).Cells(3).Value) * asm_grid.Rows(j).Cells(6).Value
            Next

            '////////////////////  ---------- fill current inventory values and fulfilled-----------------------

            For i = 0 To asm_grid.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                cmd5.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        asm_grid.Rows(i).Cells(7).Value = reader5(0).ToString
                    End While
                End If

                reader5.Close()

                '---- qty fulfilled --------

                Dim cmd58 As New MySqlCommand
                cmd58.Parameters.Clear()
                cmd58.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                cmd58.Parameters.AddWithValue("@job", job_label.Text)
                cmd58.Parameters.AddWithValue("@asm", Label20.Text)
                cmd58.Parameters.AddWithValue("@panel_name", mr_label.Text)
                cmd58.CommandText = "SELECT qty_fullfilled from Material_Request.my_assem where part_name = @part_name and job = @job and asm = @asm and panel_name = @panel_name"
                cmd58.Connection = Login.Connection
                Dim reader58 As MySqlDataReader
                reader58 = cmd58.ExecuteReader

                If reader58.HasRows Then
                    While reader58.Read
                        asm_grid.Rows(i).Cells(9).Value = reader58(0).ToString
                    End While
                End If

                reader58.Close()

                '--qty needed
                asm_grid.Rows(i).Cells(10).Value = asm_grid.Rows(i).Cells(6).Value - asm_grid.Rows(i).Cells(9).Value

                If CType(asm_grid.Rows(i).Cells(10).Value, Double) <> 0 Then
                    asm_grid.Rows(i).Cells(10).Style.BackColor = Color.Firebrick
                Else
                    asm_grid.Rows(i).Cells(10).Style.BackColor = Color.Gray
                End If

            Next
            '----------------------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Label21_MouseEnter(sender As Object, e As EventArgs) Handles Label21.MouseEnter
        Label21.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label21_MouseLeave(sender As Object, e As EventArgs) Handles Label21.MouseLeave
        Label21.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        'asm bom

        If String.Equals(Me.Text, "My Material Requests") = False Then

            wait_la.Visible = True
            Application.DoEvents()

            Call update_asm_bom()
            Call Update_current_inv_asm()  'update inventory current qtys
            Call Inventory_manage.General_inv_cal()  'fill material order form


            '--- reload entire table except current inv----
            go_ahead = False
            asm_grid.Rows.Clear()
            go_ahead = True

            wait_la.Visible = False
            MessageBox.Show("ASM BOM updated successfully!")

            Dim qty_as As Integer : qty_as = 1
            qty_as = CType(Label19.Text.Replace(")", "").Replace("(", ""), Double)

            Call open_assem(Label20.Text, qty_as)

            Call update_Qty_transit(Label20.Text, qty_as)  'update qty in transit


        End If
    End Sub


    Sub update_Qty_transit(asm As String, qty_t As Double)

        '-- if asm bom shopping cart is fullfilled then update qty in transit of that assembly

        Dim upd As Boolean : upd = True

        For i = 0 To asm_grid.Rows.Count - 1

            If CType(asm_grid.Rows(i).Cells(10).Value, Double) > 0 Then
                upd = False
                Exit For
            End If
        Next

        If upd = True Then


            Try

                Dim in_transit As Double : in_transit = 0

                '-- get qty_in transit 
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@part_name", asm)
                check_cmd.CommandText = "select Qty_in_order from inventory.inventory_qty where part_name = @part_name"
                check_cmd.Connection = Login.Connection
                check_cmd.ExecuteNonQuery()

                Dim reader As MySqlDataReader
                reader = check_cmd.ExecuteReader

                If reader.HasRows Then

                    While reader.Read
                        in_transit = reader(0).ToString
                    End While
                End If

                reader.Close()

                '-------------------------------


                Dim cmd52 As New MySqlCommand
                cmd52.Parameters.Clear()
                cmd52.Parameters.AddWithValue("@part_name", asm)
                cmd52.Parameters.AddWithValue("@Qty_in_order", If(qty_t + in_transit >= 0, qty_t + in_transit, 0))

                cmd52.CommandText = "UPDATE inventory.inventory_qty set Qty_in_order = @Qty_in_order where part_name = @part_name"
                cmd52.Connection = Login.Connection
                cmd52.ExecuteNonQuery()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If

    End Sub


    Sub update_asm_bom()

        Try

            For i = 0 To asm_grid.Rows.Count - 1

                If asm_grid.Rows(i).Cells(8).Value <> 0 Then

                    Dim qty_f As Double = 0


                    If IsNumeric(asm_grid.Rows(i).Cells(9).Value) = True Then
                        qty_f = CType(asm_grid.Rows(i).Cells(9).Value, Double)
                    End If

                    If IsNumeric(asm_grid.Rows(i).Cells(8).Value) = True Then
                        qty_f = qty_f + CType(asm_grid.Rows(i).Cells(8).Value, Double)
                    End If

                    Dim exist_c As Boolean : exist_c = False

                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@asm", Label20.Text)
                    cmd5.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@panel_name", mr_label.Text)
                    cmd5.CommandText = "SELECT * from Material_Request.my_assem where asm = @asm and part_name = @part_name and panel_name = @panel_name"
                    cmd5.Connection = Login.Connection
                    Dim reader5 As MySqlDataReader
                    reader5 = cmd5.ExecuteReader

                    If reader5.HasRows Then
                        While reader5.Read
                            exist_c = True
                        End While
                    End If

                    reader5.Close()

                    If exist_c = True Then

                        Dim cmd52 As New MySqlCommand
                        cmd52.Parameters.Clear()
                        cmd52.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                        cmd52.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                        cmd52.Parameters.AddWithValue("@job", job_label.Text)
                        cmd52.Parameters.AddWithValue("@asm", Label20.Text)
                        cmd52.Parameters.AddWithValue("@mr_name", mr_label.Text)

                        cmd52.CommandText = "UPDATE Material_Request.my_assem SET qty_fullfilled = @qty_fullfilled where job = @job and asm = @asm and part_name = @part_name"
                        cmd52.Connection = Login.Connection
                        cmd52.ExecuteNonQuery()

                    Else

                        Dim cmd53 As New MySqlCommand
                        cmd53.Parameters.Clear()
                        cmd53.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                        cmd53.Parameters.AddWithValue("@qty_fullfilled", qty_f)
                        cmd53.Parameters.AddWithValue("@job", job_label.Text)
                        cmd53.Parameters.AddWithValue("@asm", Label20.Text)
                        cmd53.Parameters.AddWithValue("@panel_name", mr_label.Text)

                        cmd53.CommandText = "INSERT INTO  Material_Request.my_assem(job, asm, part_name, qty_fullfilled, panel_name) VALUES (@job, @asm, @part_name, @qty_fullfilled, @panel_name)"
                        cmd53.Connection = Login.Connection
                        cmd53.ExecuteNonQuery()


                    End If

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Update_current_inv_asm()

        Try
            For i = 0 To asm_grid.Rows.Count - 1

                If asm_grid.Rows(i).Cells(8).Value <> 0 Then

                    Dim current_q As Double : current_q = 0
                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "SELECT current_qty from inventory.inventory_qty where part_name = @part_name"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        While reader4.Read
                            current_q = reader4(0).ToString
                        End While
                    End If

                    reader4.Close()

                    current_q = current_q - If(IsNumeric(asm_grid.Rows(i).Cells(8).Value), asm_grid.Rows(i).Cells(8).Value, 0)
                    current_q = If(current_q < 0, 0, current_q)


                    '---- update the value
                    Dim cmd5 As New MySqlCommand
                    cmd5.Parameters.Clear()
                    cmd5.Parameters.AddWithValue("@part_name", asm_grid.Rows(i).Cells(0).Value)
                    cmd5.Parameters.AddWithValue("@current_qty", current_q)

                    cmd5.CommandText = "UPDATE inventory.inventory_qty SET current_qty = @current_qty where part_name = @part_name"
                    cmd5.Connection = Login.Connection
                    cmd5.ExecuteNonQuery()

                End If
            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub AssemblyTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssemblyTableToolStripMenuItem.Click
        Call Export_asm()
    End Sub

    Private Sub asm_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles asm_grid.CellValueChanged
        '--security on hand
        If go_ahead = True Then

            For Each row As DataGridViewRow In asm_grid.Rows
                If row.IsNewRow Then Continue For
                If (IsNumeric(row.Cells(6).Value) = True And IsNumeric(row.Cells(8).Value)) Then

                    If CType(row.Cells(7).Value, Double) < CType(row.Cells(8).Value, Double) Then
                        row.Cells(8).Value = 0
                    End If

                End If
            Next
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        '-- search in map ---
        Cursor.Current = Cursors.WaitCursor


        Dim mr_t As New List(Of String)
        Dim p_names As New List(Of String)

        Dim mr As String : mr = ""

        Call Color_Module()

        For i = 0 To alabels.Count - 1

            p_names.Add(alabels(i).Text)

            Dim cmd50 As New MySqlCommand
            cmd50.Parameters.Clear()
            cmd50.Parameters.AddWithValue("@job", job_label.Text)
            cmd50.Parameters.AddWithValue("@Panel_name", alabels(i).Text)
            cmd50.CommandText = "SELECT mr_name from Material_Request.mr where Panel_name = @Panel_name and job = @job"
            cmd50.Connection = Login.Connection
            Dim reader50 As MySqlDataReader
            reader50 = cmd50.ExecuteReader

            If reader50.HasRows Then
                While reader50.Read
                    mr = reader50(0).ToString
                End While
            End If
            reader50.Close()

            '-get latest rev
            mr_t.Add(get_last_revision(mr))
        Next

        For i = 0 To plabels.Count - 1

            p_names.Add(plabels(i).Text)

            Dim cmd501 As New MySqlCommand
            cmd501.Parameters.Clear()
            cmd501.Parameters.AddWithValue("@job", job_label.Text)
            cmd501.Parameters.AddWithValue("@Panel_name", plabels(i).Text)
            cmd501.CommandText = "SELECT mr_name from Material_Request.mr where Panel_name = @Panel_name and job = @job"
            cmd501.Connection = Login.Connection
            Dim reader501 As MySqlDataReader
            reader501 = cmd501.ExecuteReader

            If reader501.HasRows Then
                While reader501.Read
                    mr = reader501(0).ToString
                End While
            End If
            reader501.Close()

            '-get latest rev
            mr_t.Add(get_last_revision(mr))
        Next

        '--field
        p_names.Add("Field BOM")

        Dim cmd5 As New MySqlCommand
        cmd5.Parameters.Clear()
        cmd5.Parameters.AddWithValue("@BOM_type", "Field")
        cmd5.Parameters.AddWithValue("@job", job_label.Text)

        cmd5.CommandText = "SELECT mr_name from Material_Request.mr where  BOM_type = @BOM_type and job = @job"
        cmd5.Connection = Login.Connection
        Dim reader5 As MySqlDataReader
        reader5 = cmd5.ExecuteReader

        If reader5.HasRows Then
            While reader5.Read
                mr = reader5(0).ToString
            End While
        End If
        reader5.Close()

        mr_t.Add(get_last_revision(mr))

        '-- special order
        p_names.Add("Spare Parts BOM")

        Dim cmd52 As New MySqlCommand
        cmd52.Parameters.Clear()
        cmd52.Parameters.AddWithValue("@BOM_type", "Spare_Parts")
        cmd52.Parameters.AddWithValue("@job", job_label.Text)

        cmd52.CommandText = "SELECT mr_name from Material_Request.mr where  BOM_type = @BOM_type and job = @job"
        cmd52.Connection = Login.Connection
        Dim reader52 As MySqlDataReader
        reader52 = cmd52.ExecuteReader

        If reader52.HasRows Then
            While reader52.Read
                mr = reader52(0).ToString
            End While
        End If
        reader52.Close()

        mr_t.Add(get_last_revision(mr))

        For i = 0 To mr_t.Count - 1

            Dim cmd501 As New MySqlCommand
            cmd501.Parameters.Clear()
            cmd501.Parameters.AddWithValue("@mr_name", mr_t.Item(i))
            cmd501.Parameters.AddWithValue("@Part_No", TextBox2.Text)
            cmd501.CommandText = "SELECT * from Material_Request.mr_data where mr_name = @mr_name and Part_No = @Part_No"
            cmd501.Connection = Login.Connection
            Dim reader501 As MySqlDataReader
            reader501 = cmd501.ExecuteReader

            If reader501.HasRows Then
                color_found(p_names.Item(i).ToString)
            End If

            reader501.Close()
        Next

        Cursor.Current = Cursors.Default
        MessageBox.Show("The part can be found in the Black BOM's")

    End Sub

    Sub color_found(name As String)
        '--color block

        If String.Equals(name, "Field BOM") = True Then
            F_BOM.BackColor = Color.Black

        ElseIf String.Equals(name, "Spare Parts BOM") = True Then
            SP_BOM.BackColor = Color.Black
        End If

        For i = 0 To plabels.Count - 1
            If String.Equals(plabels(i).Text, name) Then
                plabels(i).BackColor = Color.Black
            End If
        Next

        For i = 0 To alabels.Count - 1
            If String.Equals(alabels(i).Text, name) Then
                alabels(i).BackColor = Color.Black
            End If
        Next

    End Sub

    'Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    '    '--- test ---

    '    Dim my_assemblies = New List(Of String)()

    '    '--------------  add to device
    '    Dim cmd2 As New MySqlCommand
    '    cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
    '    cmd2.Connection = Login.Connection
    '    Dim reader2 As MySqlDataReader
    '    reader2 = cmd2.ExecuteReader

    '    If reader2.HasRows Then
    '        While reader2.Read
    '            my_assemblies.Add(reader2(0))
    '        End While
    '    End If

    '    reader2.Close()



    '    Dim dimen_table1 = New DataTable
    '    dimen_table1.Columns.Add("Part_No", GetType(String))
    '    dimen_table1.Columns.Add("Vendor", GetType(String))

    '    Try
    '        Dim cmd4 As New MySqlCommand
    '        cmd4.CommandText = "select distinct Part_No, Vendor from Material_Request.mr_data;"

    '        cmd4.Connection = Login.Connection
    '        Dim reader4 As MySqlDataReader
    '        reader4 = cmd4.ExecuteReader

    '        If reader4.HasRows Then

    '            While reader4.Read
    '                dimen_table1.Rows.Add(reader4(0).ToString, reader4(1).ToString)
    '            End While

    '        End If

    '        reader4.Close()

    '        For i = 0 To dimen_table1.Rows.Count - 1

    '            Dim Create_cmd2 As New MySqlCommand
    '            Create_cmd2.Parameters.Clear()
    '            Create_cmd2.Parameters.AddWithValue("@Part_No", dimen_table1.Rows(i).Item(0))

    '            Dim cost As Decimal : cost = 0

    '            If my_assemblies.Contains(dimen_table1.Rows(i).Item(0)) = True Then
    '                cost = myQuote.Cost_of_Assem(dimen_table1.Rows(i).Item(0))
    '            Else
    '                cost = Form1.Get_Latest_Cost(Login.Connection, dimen_table1.Rows(i).Item(0), dimen_table1.Rows(i).Item(1))
    '            End If

    '            Create_cmd2.Parameters.AddWithValue("@Price", cost)


    '            Create_cmd2.CommandText = "UPDATE Material_Request.mr_data  SET Price = @Price where Part_No = @Part_No"
    '            Create_cmd2.Connection = Login.Connection
    '            Create_cmd2.ExecuteNonQuery()
    '        Next

    '        MessageBox.Show("done")

    '    Catch ex As Exception
    '        MessageBox.Show(ex.ToString)
    '    End Try


    'End Sub


End Class