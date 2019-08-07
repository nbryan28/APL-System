Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection

Public Class Report_tracking_mn
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If String.Equals(Proc_Material_R.open_job, "Open Project:") = False Then

            '---------- Lets create a tracking report
            Try
                '-- get id of BOM
                Dim id_bom As String : id_bom = 1
                Dim cmd21 As New MySqlCommand
                cmd21.Parameters.AddWithValue("@mr_name", Proc_Material_R.Text)
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
                check_cmd2.Parameters.AddWithValue("@job", Proc_Material_R.open_job)
                check_cmd2.Parameters.AddWithValue("@id_bom", id_bom)
                check_cmd2.CommandText = "delete from Tracking_Reports.my_tracking_reports where job = @job and id_bom = @id_bom"
                check_cmd2.Connection = Login.Connection
                check_cmd2.ExecuteNonQuery()

                '---insert data

                For i = 0 To orders_grid.Rows.Count - 1

                    If IsNumeric(orders_grid.Rows(i).Cells(4).Value.ToString) = True And String.Equals(orders_grid.Rows(i).Cells(4).Value.ToString, "0") = False Then
                        Dim Create_cmd6 As New MySqlCommand
                        Create_cmd6.Parameters.Clear()
                        Create_cmd6.Parameters.AddWithValue("@job", Proc_Material_R.open_job)
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

                        '  Call Form1.Command_h(current_user, "(" & orders_grid.Rows(i).Cells(4).Value.ToString & ")" & " " & orders_grid.Rows(i).Cells(0).Value.ToString & " purchased. PO: " & If(orders_grid.Rows(i).Cells(5).Value Is Nothing, "", orders_grid.Rows(i).Cells(5).Value.ToString), Proc_Material_R.open_job)
                    End If

                Next

                MessageBox.Show("Report updated")

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub Report_tracking_mn_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
            cmd21.Parameters.AddWithValue("@mr_name", Proc_Material_R.Text)
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
            cmd4.Parameters.AddWithValue("@job", Proc_Material_R.open_job)
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



        For i = 0 To Proc_Material_R.total_grid.Rows.Count - 1

            Try

                '--check if part exist in inventory
                Dim exist_c As Boolean : exist_c = False
                Dim check_cmd As New MySqlCommand
                check_cmd.Parameters.AddWithValue("@part_name", Proc_Material_R.total_grid.Rows(i).Cells(1).Value)
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
                    If my_assemblies.Contains(Proc_Material_R.total_grid.Rows(i).Cells(1).Value) = True Then
                        Call Get_need_buy(Proc_Material_R.total_grid.Rows(i).Cells(1).Value, Proc_Material_R.total_grid.Rows(i).Cells(0).Value)
                    Else

                        If orders_grid.Rows.Count > 0 Then

                            Dim new_qty As Double = Isintable(Proc_Material_R.total_grid.Rows(i).Cells(1).Value, Proc_Material_R.total_grid.Rows(i).Cells(0).Value)

                            If new_qty > 0 Then
                                orders_grid.Rows.Add(New String() {Proc_Material_R.total_grid.Rows(i).Cells(1).Value, Proc_Material_R.total_grid.Rows(i).Cells(3).Value, Proc_Material_R.total_grid.Rows(i).Cells(5).Value, Proc_Material_R.total_grid.Rows(i).Cells(6).Value, new_qty})
                            End If

                        Else
                            orders_grid.Rows.Add(New String() {})
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(0).Value = Proc_Material_R.total_grid.Rows(i).Cells(1).Value 'part name
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(1).Value = Proc_Material_R.total_grid.Rows(i).Cells(3).Value 'manu
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(2).Value = Proc_Material_R.total_grid.Rows(i).Cells(5).Value  'vendor
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(3).Value = Proc_Material_R.total_grid.Rows(i).Cells(6).Value  'cost
                            orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(4).Value = Proc_Material_R.total_grid.Rows(i).Cells(0).Value  'qty
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'export  to excel file

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then

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

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Special_orders_" & Proc_Material_R.open_job & ".xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Data exported successfully!")
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If orders_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(orders_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

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

                        'If orders_grid.Rows.Count > 0 Then

                        '    Dim new_qty As Double = Isintable(datatable.Rows(i).Item(0), datatable.Rows(i).Item(4))

                        '    If new_qty > 0 Then
                        '        orders_grid.Rows.Add(New String() {datatable.Rows(i).Item(0), datatable.Rows(i).Item(1), datatable.Rows(i).Item(2), datatable.Rows(i).Item(3), new_qty})
                        '    End If

                        'Else
                        orders_grid.Rows.Add(New String() {})
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(0).Value = datatable.Rows(i).Item(0) 'part name
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(1).Value = datatable.Rows(i).Item(1) 'manu
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(2).Value = datatable.Rows(i).Item(2)  'vendor
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(3).Value = datatable.Rows(i).Item(3)  'cost
                        orders_grid.Rows(orders_grid.Rows.Count - 1).Cells(4).Value = datatable.Rows(i).Item(4)  'qty
                        '  End If

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

End Class