Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail



Public Class Total_Project_Cost
    Private Sub Total_Project_Cost_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            job_box.Items.Clear()

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    job_box.Items.Add(reader2(0))
                End While
            End If

            reader2.Close()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '--- export data to excel
        If Not job_box.SelectedItem Is Nothing Then


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

                    xlWorkSheet.Cells(1, 1) = "Project " & job_box.Text
                    xlWorkSheet.Cells(3, 1) = "Inventory Cost"
                    xlWorkSheet.Cells(3, 2) = gi_box.Text
                    xlWorkSheet.Cells(4, 1) = "Project Specific Cost"
                    xlWorkSheet.Cells(4, 2) = ps_box.Text
                    xlWorkSheet.Cells(5, 1) = "Add and Return Cost"
                    xlWorkSheet.Cells(5, 2) = ar_box.Text
                    xlWorkSheet.Cells(6, 1) = "Bulk Cost"
                    xlWorkSheet.Cells(6, 2) = bulk_box.Text
                    xlWorkSheet.Cells(7, 1) = "Shipping Cost"
                    xlWorkSheet.Cells(7, 2) = ship_box.Text
                    xlWorkSheet.Cells(9, 1) = "Total Project Cost"
                    xlWorkSheet.Cells(9, 2) = total_box.Text



                    SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                    If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                        xlWorkBook.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                    End If

                    xlWorkBook.Close(False)
                    xlApp.Quit()


                    Marshal.ReleaseComObject(xlApp)
                    MessageBox.Show("Data exported successfully!")

                Catch ex As Exception
                    MessageBox.Show(ex.ToString)
                End Try
            End If
        End If
    End Sub

    Private Sub job_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles job_box.SelectedValueChanged

        '---- calculate Totals ----
        '-- if job is not closed -

        '------- clear boxes ------
        gi_box.Clear()
        ps_box.Clear()
        ar_box.Clear()
        bulk_box.Clear()
        ship_box.Clear()
        total_box.Clear()

        Try

            Dim is_closed As Boolean : is_closed = False

            Dim cmd21 As New MySqlCommand
            cmd21.Parameters.AddWithValue("@job", job_box.Text)
            cmd21.CommandText = "SELECT * from Material_Request.mr where job = @job and finished = 'Y'"
            cmd21.Connection = Login.Connection
            Dim reader21 As MySqlDataReader
            reader21 = cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    is_closed = True
                    status_l.Text = "Status: Closed"
                End While

            Else
                status_l.Text = "Status: Open"
            End If

            reader21.Close()

            If is_closed = False Then

                Call Cal_Totals(job_box.Text)
            Else

                Call show_totals(job_box.Text)
            End If

            total_box.Text = Decimal.Round(If(IsNumeric(gi_box.Text) = True, CType(gi_box.Text, Double), 0) + If(IsNumeric(ps_box.Text) = True, CType(ps_box.Text, Double), 0) + If(IsNumeric(ar_box.Text) = True, CType(ar_box.Text, Double), 0) + If(IsNumeric(bulk_box.Text) = True, CType(bulk_box.Text, Double), 0) + If(IsNumeric(ship_box.Text) = True, CType(ship_box.Text, Double), 0), 2, MidpointRounding.AwayFromZero).ToString("N")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try



    End Sub

    Sub show_totals(job As String)

        '---- display data from total cost ----
        Try
            Dim cmd21 As New MySqlCommand
            cmd21.Parameters.AddWithValue("@job", job_box.Text)
            cmd21.CommandText = "SELECT gi_cost, ps_cost, ar_cost, bulk_cost, shipping_cost from management.total_cost where job = @job"
            cmd21.Connection = Login.Connection
            Dim reader21 As MySqlDataReader
            reader21 = cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    gi_box.Text = If(reader21(0) Is DBNull.Value, 0, Decimal.Round(reader21(0).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    ps_box.Text = If(reader21(1) Is DBNull.Value, 0, Decimal.Round(reader21(1).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    ar_box.Text = If(reader21(2) Is DBNull.Value, 0, Decimal.Round(reader21(2).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    bulk_box.Text = If(reader21(3) Is DBNull.Value, 0, Decimal.Round(reader21(3).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    ship_box.Text = If(reader21(4) Is DBNull.Value, 0, Decimal.Round(reader21(4).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                End While
            End If

            reader21.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Sub Cal_Totals(job As String)

        '----------- Cal totals of job ---------------

        Dim Inv_c As Double : Inv_c = 0
        Dim PS_c As Double : PS_c = 0
        Dim AR_c As Double : AR_c = 0
        Dim bulk As Double : bulk = 0   'bulk
        Dim ship As Double : ship = 0   'ship


        '--- table general inv ---------
        Dim dimen_table = New DataTable

        dimen_table.Columns.Add("Part_No", GetType(String))
        dimen_table.Columns.Add("Price", GetType(String))
        dimen_table.Columns.Add("Qty", GetType(String))
        dimen_table.Columns.Add("Min", GetType(String))


        '--- table project specific ---------
        Dim ps_table = New DataTable

        ps_table.Columns.Add("Part_No", GetType(String))
        ps_table.Columns.Add("Qty", GetType(String))
        ps_table.Columns.Add("Cost", GetType(String))

        '--- table add-return ---------
        Dim ar_table = New DataTable

        ar_table.Columns.Add("Part_No", GetType(String))
        ar_table.Columns.Add("Qty", GetType(String))
        ar_table.Columns.Add("Cost", GetType(String))
        ar_table.Columns.Add("task", GetType(String))

        '----------------------------------------



        '--------------- fill tables ---------------
        Try

            '-----------/////////////// fill only inventory table ////////////-------------------

            Dim mr_name As String : mr_name = "not_found"
            Dim cmd21 As New MySqlCommand
            cmd21.Parameters.AddWithValue("@job", job)
            cmd21.CommandText = "select mr_name from Material_Request.mr where job = @job and (BOM_type = 'old_BOM' or BOM_Type = 'MB' or BOM_type = 'ASM') order by release_date desc limit 1"
            cmd21.Connection = Login.Connection
            Dim reader21 As MySqlDataReader
            reader21 = cmd21.ExecuteReader

            If reader21.HasRows Then
                While reader21.Read
                    mr_name = reader21(0).ToString
                End While
            End If

            reader21.Close()

            mr_name = Procurement_Overview.get_last_revision(mr_name)

            '--/////////////////////

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr_name", mr_name)
            cmd4.CommandText = "SELECT  Part_No, Price,  Qty from Material_Request.mr_data where mr_name = @mr_name"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader


            If reader4.HasRows Then
                While reader4.Read
                    dimen_table.Rows.Add(reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, 0)
                End While
            End If

            reader4.Close()

            '-----------------

            For i = 0 To dimen_table.Rows.Count - 1

                Dim cmd5 As New MySqlCommand
                cmd5.Parameters.Clear()
                cmd5.Parameters.AddWithValue("@part_name", dimen_table.Rows(i).Item(0).ToString)
                cmd5.CommandText = "SELECT min_qty from inventory.inventory_qty where part_name = @part_name"
                cmd5.Connection = Login.Connection
                Dim reader5 As MySqlDataReader
                reader5 = cmd5.ExecuteReader

                If reader5.HasRows Then
                    While reader5.Read
                        dimen_table.Rows(i).Item(3) = reader5(0)
                    End While
                End If

                reader5.Close()

                '-- get totals --

                If CType(dimen_table.Rows(i).Item(3), Double) > 0 Then
                    Inv_c = Inv_c + (If(IsNumeric(dimen_table.Rows(i).Item(1)), dimen_table.Rows(i).Item(1), 0) * If(IsNumeric(dimen_table.Rows(i).Item(1)), dimen_table.Rows(i).Item(1), 0))
                End If

                '-------------------

            Next

            '-----------/////////////// fill Project specific ////////////-------------------
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@job", job)
            cmd41.CommandText = "SELECT Part_No, qty_purchased, cost from Tracking_Reports.my_tracking_reports where job = @job"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader


            If reader41.HasRows Then
                While reader41.Read
                    ps_table.Rows.Add(reader41(0).ToString, reader41(1).ToString, reader41(2).ToString)
                End While
            End If

            reader41.Close()

            For i = 0 To ps_table.Rows.Count - 1
                PS_c = PS_c + (If(IsNumeric(ps_table.Rows(i).Item(1)), ps_table.Rows(i).Item(1), 0) * If(IsNumeric(ps_table.Rows(i).Item(2)), ps_table.Rows(i).Item(2), 0))
            Next

            '------------------------------------------------------------

            '-----------/////////////// fill add return  ////////////-------------------

            Dim cmd42 As New MySqlCommand
            cmd42.Parameters.AddWithValue("@job", job)
            cmd42.CommandText = "SELECT Part_No, qty, Cost, task from Tracking_Reports.add_return where job = @job"
            cmd42.Connection = Login.Connection
            Dim reader42 As MySqlDataReader
            reader42 = cmd42.ExecuteReader


            If reader42.HasRows Then
                While reader42.Read
                    ar_table.Rows.Add(reader42(0).ToString, reader42(1).ToString, reader42(2).ToString, reader42(3).ToString)
                End While
            End If

            reader42.Close()

            For i = 0 To ar_table.Rows.Count - 1

                If String.Equals(dimen_table.Rows(i).Item(3), "Add") Then
                    AR_c = AR_c + (If(IsNumeric(ar_table.Rows(i).Item(1)), ar_table.Rows(i).Item(1), 0) * If(IsNumeric(ar_table.Rows(i).Item(2)), ar_table.Rows(i).Item(2), 0))
                Else
                    AR_c = AR_c + (If(IsNumeric(ar_table.Rows(i).Item(1)), ar_table.Rows(i).Item(1), 0) * If(IsNumeric(ar_table.Rows(i).Item(2)), ar_table.Rows(i).Item(2), 0) * -1)
                End If
            Next

            '---------- shipping and bulk ------
            Dim cmd211 As New MySqlCommand
            cmd211.Parameters.AddWithValue("@job", job_box.Text)
            cmd211.CommandText = "SELECT bulk_cost, shipping_cost from management.total_cost where job = @job"
            cmd211.Connection = Login.Connection
            Dim reader211 As MySqlDataReader
            reader211 = cmd211.ExecuteReader

            If reader211.HasRows Then
                While reader211.Read
                    bulk = If(reader211(0) Is DBNull.Value, 0, Decimal.Round(reader211(0).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                    ship = If(reader211(1) Is DBNull.Value, 0, Decimal.Round(reader211(1).ToString, 2, MidpointRounding.AwayFromZero).ToString("N"))
                End While
            End If

            reader211.Close()
            '-----------------------------------

            '----------- put all data in textboxes --------------
            gi_box.Text = Decimal.Round(Inv_c, 2, MidpointRounding.AwayFromZero).ToString("N")
            ps_box.Text = Decimal.Round(PS_c, 2, MidpointRounding.AwayFromZero).ToString("N")
            ar_box.Text = Decimal.Round(AR_c, 2, MidpointRounding.AwayFromZero).ToString("N")
            bulk_box.Text = bulk
            ship_box.Text = ship
            '----------------------------------------------------

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        If String.Equals(status_l.Text, "Status: Open") = True And job_box.SelectedItem Is Nothing = False Then

            Dim exist_c As Boolean : exist_c = False

            Dim cmd5 As New MySqlCommand
            cmd5.Parameters.Clear()
            cmd5.Parameters.AddWithValue("@job", job_box.SelectedItem)
            cmd5.CommandText = "SELECT * from management.total_cost where job = @job"
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

                '-- update
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@job", job_box.Text)
                Create_cmd.Parameters.AddWithValue("@bulk_cost", If(IsNumeric(bulk_box.Text), bulk_box.Text, 0))
                Create_cmd.Parameters.AddWithValue("@shipping_cost", If(IsNumeric(ship_box.Text), ship_box.Text, 0))


                Create_cmd.CommandText = "UPDATE management.total_cost  SET  shipping_cost = @shipping_cost, bulk_cost = @bulk_cost  where job = @job"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            Else

                '-- insert
                Dim Create_cmd As New MySqlCommand
                Create_cmd.Parameters.AddWithValue("@job", job_box.Text)
                Create_cmd.Parameters.AddWithValue("@bulk_cost", If(IsNumeric(bulk_box.Text), bulk_box.Text, 0))
                Create_cmd.Parameters.AddWithValue("@shipping_cost", If(IsNumeric(ship_box.Text), ship_box.Text, 0))

                Create_cmd.CommandText = "INSERT INTO management.total_cost(job, bulk_cost, shipping_cost) VALUES (@job, @bulk_cost, @shipping_cost)"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

            End If

        Else
                MessageBox.Show("Sorry the job is closed or you have not select one")

        End If

    End Sub
End Class