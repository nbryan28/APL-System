Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class Part_report



    Private Sub Procurement_report_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call EnableDoubleBuffered(O_grid)
        Call EnableDoubleBuffered(A_grid)
        ' Call EnableDoubleBuffered(asm_grid)


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


    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            O_grid.Rows.Clear()
            A_grid.Rows.Clear()
            f_grid.Rows.Clear()

            Dim my_mr = New List(Of String)()

            '--------------- populate assemblies -----------

            Dim cmd_po1 As New MySqlCommand
            cmd_po1.Parameters.AddWithValue("@part", TextBox1.Text)
            cmd_po1.CommandText = "Select legacy_ADA, Qty from adv where Part_Name = @part"

            cmd_po1.Connection = Login.Connection
            Dim readerv1 As MySqlDataReader
            readerv1 = cmd_po1.ExecuteReader

            If readerv1.HasRows Then
                Dim i As Integer : i = 0
                While readerv1.Read

                    A_grid.Rows.Add()  'add a new row
                    A_grid.Rows(i).Cells(0).Value = readerv1(0).ToString
                    A_grid.Rows(i).Cells(1).Value = readerv1(1).ToString

                    i = i + 1

                End While
            End If

            readerv1.Close()

            '-----------------------------------------
            Dim cmd_po As New MySqlCommand
            cmd_po.Parameters.AddWithValue("@part", TextBox1.Text)
            cmd_po.CommandText = "Select Part_No, Description from Material_Request.mr_data where Part_No = @part"

            cmd_po.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_po.ExecuteReader

            If readerv.HasRows Then

                While readerv.Read

                    part_l.Text = readerv(0).ToString
                    desc_l.Text = readerv(1).ToString
                End While
            End If

            readerv.Close()

            '---------------------------------------------------
            '===========   mrs  ===============================

            Dim cmd_po12 As New MySqlCommand
            cmd_po12.Parameters.AddWithValue("@part", TextBox1.Text)
            cmd_po12.CommandText = "Select distinct mr_name from Material_Request.mr_data where Part_No = @part and latest_r = 'x'"

            cmd_po12.Connection = Login.Connection
            Dim readervp As MySqlDataReader
            readervp = cmd_po12.ExecuteReader

            If readervp.HasRows Then

                While readervp.Read
                    my_mr.Add(readervp(0))
                End While
            End If

            readervp.Close()


            '---
            For i = 0 To my_mr.Count - 1

                '--- get job and description ---
                Dim job_n As String : job_n = "x"
                Dim job_desc As String : job_desc = ""
                Dim need As Double : need = 0

                Dim asm_n As Double : asm_n = 0
                Dim p_collect As Double : p_collect = 0

                Dim cmd_mr As New MySqlCommand
                cmd_mr.Parameters.AddWithValue("@mr_name", my_mr.Item(i))
                cmd_mr.CommandText = "Select job from Material_Request.mr where mr_name = @mr_name and BOM_type not like 'ASM' and released = 'Y'"

                cmd_mr.Connection = Login.Connection
                Dim reader_my As MySqlDataReader
                reader_my = cmd_mr.ExecuteReader

                If reader_my.HasRows Then
                    While reader_my.Read
                        job_n = reader_my(0)
                    End While
                End If

                reader_my.Close()

                '----------------------------
                Dim cmd_mr2 As New MySqlCommand
                cmd_mr2.Parameters.AddWithValue("@job", job_n)
                cmd_mr2.CommandText = "Select Job_description from management.projects where Job_number = @job"

                cmd_mr2.Connection = Login.Connection
                Dim reader_my2 As MySqlDataReader
                reader_my2 = cmd_mr2.ExecuteReader

                If reader_my2.HasRows Then
                    While reader_my2.Read
                        job_desc = reader_my2(0)
                    End While
                End If

                reader_my2.Close()

                '---------- getting qty need individual  -------

                '--- new procedure--
                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_No", TextBox1.Text)
                cmd.Parameters.AddWithValue("@mr_name", my_mr.Item(i))
                cmd.CommandText = "select Qty, qty_fullfilled from Material_Request.mr_data where Part_No = @Part_No and mr_name = @mr_name and latest_r = 'x' and released = 'Y'"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        need = If(IsDBNull(reader(0)), 0, reader(0)) - If(IsDBNull(reader(1)), 0, reader(1))
                    End While
                End If

                reader.Close()
                '-------------------------


                If String.Equals(job_n, "x") = False And need > 0 Then
                    O_grid.Rows.Add(New String() {job_n, job_desc, need})
                End If

            Next




            '--------------------------------
            '------ feature code -----------
            Dim cmd_po2 As New MySqlCommand
            cmd_po2.Parameters.AddWithValue("@part", TextBox1.Text)
            cmd_po2.CommandText = "Select Feature_code, solution, qty from quote_table.feature_parts where part_name = @part"

            cmd_po2.Connection = Login.Connection
            Dim readerv2 As MySqlDataReader
            readerv2 = cmd_po2.ExecuteReader

            If readerv2.HasRows Then
                Dim i As Integer : i = 0
                While readerv2.Read

                    f_grid.Rows.Add()  'add a new row
                    f_grid.Rows(i).Cells(0).Value = readerv2(0).ToString
                    f_grid.Rows(i).Cells(1).Value = readerv2(1).ToString
                    f_grid.Rows(i).Cells(2).Value = readerv2(2).ToString

                    i = i + 1

                End While
            End If

            readerv2.Close()

            Call assem_cal()

            '----------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        '------------------------------------------------


    End Sub

    Sub assem_cal()

        '------- how many assemblies that contain that part are demanded in that job


        For j = 0 To A_grid.Rows.Count - 1

            '--- get jobs that contains this assembly --------
            Dim my_mr2 = New List(Of String)()
            Dim job_n As String : job_n = "X"
            Dim job_desc As String : job_desc = ""
            Dim asm_n As Double : asm_n = 0
            Dim p_collect As Double : p_collect = 0

            Dim cmd_po12 As New MySqlCommand
            cmd_po12.Parameters.AddWithValue("@part", A_grid.Rows(j).Cells(0).Value)
            cmd_po12.CommandText = "Select distinct mr_name from Material_Request.mr_data where Part_No = @part and latest_r = 'x'"

            cmd_po12.Connection = Login.Connection
            Dim readervp As MySqlDataReader
            readervp = cmd_po12.ExecuteReader

            If readervp.HasRows Then

                While readervp.Read
                    my_mr2.Add(readervp(0))
                End While
            End If

            readervp.Close()

            For k = 0 To my_mr2.Count - 1

                '-------- get job --------
                Dim cmd_mr As New MySqlCommand
                cmd_mr.Parameters.AddWithValue("@mr_name", my_mr2.Item(k))
                cmd_mr.CommandText = "Select job from Material_Request.mr where mr_name = @mr_name and BOM_type not like 'ASM' and released = 'Y'"

                cmd_mr.Connection = Login.Connection
                Dim reader_my As MySqlDataReader
                reader_my = cmd_mr.ExecuteReader

                If reader_my.HasRows Then
                    While reader_my.Read
                        job_n = reader_my(0)
                    End While
                End If

                reader_my.Close()

                ' ---------get desc ---------------
                Dim cmd_mr2 As New MySqlCommand
                cmd_mr2.Parameters.AddWithValue("@job", job_n)
                cmd_mr2.CommandText = "Select Job_description from management.projects where Job_number = @job"

                cmd_mr2.Connection = Login.Connection
                Dim reader_my2 As MySqlDataReader
                reader_my2 = cmd_mr2.ExecuteReader

                If reader_my2.HasRows Then
                    While reader_my2.Read
                        job_desc = reader_my2(0)
                    End While
                End If

                reader_my2.Close()
                '----------------------------------------------

                Dim hm As Double : hm = 0
                hm = A_grid.Rows(j).Cells(1).Value

                Dim cmdx As New MySqlCommand
                cmdx.Parameters.AddWithValue("@Part_No", A_grid.Rows(j).Cells(0).Value)
                cmdx.Parameters.AddWithValue("@mr_name", my_mr2.Item(k))
                cmdx.CommandText = "select Qty, qty_fullfilled from Material_Request.mr_data where Part_No = @Part_No and mr_name = @mr_name and latest_r = 'x' and released = 'Y'"
                cmdx.Connection = Login.Connection
                Dim readerx As MySqlDataReader
                readerx = cmdx.ExecuteReader

                If readerx.HasRows Then
                    While readerx.Read
                        asm_n = ((If(IsDBNull(readerx(0)), 0, readerx(0)) - If(IsDBNull(readerx(1)), 0, readerx(1))) * hm)
                    End While
                End If

                readerx.Close()


                '------------- fullfilled --------------
                Dim cmdx1 As New MySqlCommand
                cmdx1.Parameters.AddWithValue("@part_name", TextBox1.Text)
                cmdx1.Parameters.AddWithValue("@asm", A_grid.Rows(j).Cells(0).Value)
                cmdx1.Parameters.AddWithValue("@job", job_n)
                cmdx1.CommandText = "select sum(qty_fullfilled) from Material_Request.my_assem where part_name = @part_name and job = @job and asm = @asm"
                cmdx1.Connection = Login.Connection
                Dim readerx1 As MySqlDataReader
                readerx1 = cmdx1.ExecuteReader

                If readerx1.HasRows Then
                    While readerx1.Read
                        p_collect = If(IsDBNull(readerx1(0)), 0, readerx1(0))
                    End While
                End If

                readerx1.Close()

                Dim new_need As Double : new_need = asm_n - p_collect



                Dim found_b As Boolean : found_b = False

                For z = 0 To O_grid.Rows.Count - 1
                    If String.Equals(O_grid.Rows(z).Cells(0).Value, job_n) = True And new_need > 0 Then
                        O_grid.Rows(z).Cells(2).Value = O_grid.Rows(z).Cells(2).Value + new_need
                        found_b = True
                    End If
                Next

                If found_b = False And new_need > 0 Then
                    O_grid.Rows.Add(New String() {job_n, job_desc, new_need})
                End If

            Next
            '--------------------------------------
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call export_Excel(O_grid)
    End Sub

    Sub export_Excel(mygrid As DataGridView)
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

                Cursor.Current = Cursors.Default

                MessageBox.Show("Table exported successfully!")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call export_Excel(A_grid)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call export_Excel(f_grid)
    End Sub
End Class