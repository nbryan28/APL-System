Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection

Public Class Material_order
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Material_order_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call EnableDoubleBuffered(MO_grid)

        '  Call Inventory_manage.refresh_table()

        Try

            '--- fill assemblies list --
            'Dim my_assem As New Dictionary(Of String, String)

            'Dim cmd2 As New MySqlCommand
            'cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            'cmd2.Connection = Login.Connection
            'Dim reader2 As MySqlDataReader
            'reader2 = cmd2.ExecuteReader

            'If reader2.HasRows Then
            '    While reader2.Read
            '        my_assem.Add(reader2(0).ToString, reader2(0).ToString)
            '    End While
            'End If

            'reader2.Close()
            '------------------------------
            Dim my_assemblies = New List(Of String)()

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


            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT Part_No, description, Qty_needed from inventory.Material_orders"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then

                Dim i As Integer : i = 0
                While reader4.Read
                    ' If my_assem.ContainsKey(reader4(0).ToString) = False Then
                    MO_grid.Rows.Add(New String() {})
                    MO_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    MO_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    MO_grid.Rows(i).Cells(4).Value = reader4(2).ToString
                    i = i + 1
                    '  End If
                End While

            End If

            reader4.Close()

            '---- remove the ones that are above min already (on hand + qty_transit - demand) > min
            'For i = MO_grid.Rows.Count - 1 To 0 Step -1
            '    If Is_above_min(MO_grid.Rows(i).Cells(0).Value) = True Then
            '        MO_grid.Rows.RemoveAt(i)
            '    End If
            'Next
            '----------------------------------------

            For i = 0 To MO_grid.Rows.Count - 1
                '--- get manufacturer and vendor -------

                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@part", MO_grid.Rows(i).Cells(0).Value)
                cmd.CommandText = "SELECT Manufacturer, Primary_Vendor from parts_table where Part_Name = @part"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        MO_grid.Rows(i).Cells(2).Value = reader(0).ToString
                        MO_grid.Rows(i).Cells(3).Value = reader(1).ToString
                    End While

                End If

                reader.Close()

                '------------  cost ---------------------
                If my_assemblies.Contains(MO_grid.Rows(i).Cells(0).Value) = True Then
                    MO_grid.Rows(i).Cells(7).Value = "$" & myQuote.Cost_of_Assem(MO_grid.Rows(i).Cells(0).Value)
                Else
                    MO_grid.Rows(i).Cells(7).Value = "$" & Form1.Get_Latest_Cost(Login.Connection, MO_grid.Rows(i).Cells(0).Value, MO_grid.Rows(i).Cells(3).Value)
                End If
            Next




        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Sub Refresh_table()

        '----- refresh material order form ------------
        Dim my_assemblies = New List(Of String)()

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

        MO_grid.Rows.Clear()

        Try

            Dim cmd4 As New MySqlCommand
            cmd4.CommandText = "SELECT Part_No, description,  Qty_needed from inventory.Material_orders"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                Dim i As Integer : i = 0
                While reader4.Read
                    MO_grid.Rows.Add(New String() {})
                    MO_grid.Rows(i).Cells(0).Value = reader4(0).ToString
                    MO_grid.Rows(i).Cells(1).Value = reader4(1).ToString
                    MO_grid.Rows(i).Cells(4).Value = reader4(2).ToString
                    i = i + 1
                End While

            End If

            reader4.Close()

            For i = 0 To MO_grid.Rows.Count - 1
                '--- get manufacturer and vendor -------

                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@part", MO_grid.Rows(i).Cells(0).Value)
                cmd.CommandText = "SELECT Manufacturer, Primary_Vendor from parts_table where Part_Name = @part"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        MO_grid.Rows(i).Cells(2).Value = reader(0).ToString
                        MO_grid.Rows(i).Cells(3).Value = reader(1).ToString
                    End While

                End If

                reader.Close()

                '------------  cost ---------------------
                If my_assemblies.Contains(MO_grid.Rows(i).Cells(0).Value) = True Then
                    MO_grid.Rows(i).Cells(7).Value = "$" & myQuote.Cost_of_Assem(MO_grid.Rows(i).Cells(0).Value)
                Else
                    MO_grid.Rows(i).Cells(7).Value = "$" & Form1.Get_Latest_Cost(Login.Connection, MO_grid.Rows(i).Cells(0).Value, MO_grid.Rows(i).Cells(3).Value)
                End If


            Next

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-------- update order_qty in inventory management table ----------

        Try

            For i = 0 To MO_grid.Rows.Count - 1

                If String.IsNullOrEmpty(MO_grid.Rows(i).Cells(5).Value) = False And String.Equals(MO_grid.Rows(i).Cells(5).Value, "") = False Then

                    Dim cmd4 As New MySqlCommand
                    cmd4.Parameters.AddWithValue("@part_name", MO_grid.Rows(i).Cells(0).Value)
                    cmd4.CommandText = "SELECT part_name, Qty_in_order from inventory.inventory_qty where part_name = @part_name"
                    cmd4.Connection = Login.Connection
                    Dim reader4 As MySqlDataReader
                    reader4 = cmd4.ExecuteReader

                    If reader4.HasRows Then
                        While reader4.Read
                            Dim Create_cmd As New MySqlCommand
                            Create_cmd.Parameters.AddWithValue("@part_name", MO_grid.Rows(i).Cells(0).Value)
                            Create_cmd.Parameters.AddWithValue("@Qty_in_order", MO_grid.Rows(i).Cells(5).Value)
                            Create_cmd.Parameters.AddWithValue("@es_date_of_arrival", MO_grid.Rows(i).Cells(6).Value)
                            Create_cmd.Parameters.AddWithValue("@old_QTY", reader4(1))

                            If IsNumeric(reader4(1)) = True Then
                                reader4.Close()
                                Create_cmd.CommandText = "UPDATE inventory.inventory_qty  SET  Qty_in_order = @Qty_in_order + @old_QTY, es_date_of_arrival = @es_date_of_arrival  where part_Name = @part_Name"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()
                                Exit While

                            Else
                                reader4.Close()
                                Create_cmd.CommandText = "UPDATE inventory.inventory_qty  SET  Qty_in_order = @Qty_in_order, es_date_of_arrival = @es_date_of_arrival where part_Name = @part_Name"
                                Create_cmd.Connection = Login.Connection
                                Create_cmd.ExecuteNonQuery()
                                Exit While
                            End If
                        End While
                    Else
                        reader4.Close()
                    End If



                End If

            Next

            Call Inventory_manage.General_inv_cal()   'recalculate inventory values
            Call Refresh_table()
            ' MessageBox.Show("Done")

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'export  MPaterial order to excel file

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
                For i As Integer = 0 To MO_grid.ColumnCount - 1

                    xlWorkSheet.Cells(1, i + 1) = MO_grid.Columns(i).HeaderText
                    For j As Integer = 0 To MO_grid.RowCount - 1
                        xlWorkSheet.Cells(j + 2, i + 1) = MO_grid.Rows(j).Cells(i).Value
                    Next j

                Next i

                Try
                    xlWorkBook.SaveAs(FolderBrowserDialog1.SelectedPath & "\Material_Orders.xlsx")
                Catch ex As Exception

                End Try
                xlWorkBook.Close(True)
                xlApp.Quit()


                Marshal.ReleaseComObject(xlApp)

                MessageBox.Show("Material Order exported successfully!")
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If MO_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(MO_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Function Is_above_min(part As String) As Boolean
        '------- if the on hand + qty in transit - demand >= min then remove from grid

        Is_above_min = False

        Dim onhand As Double : onhand = 0
        Dim qty_transit As Double : qty_transit = 0
        Dim demand As Double : demand = 0
        Dim min As Double : min = 0

        Dim cmd As New MySqlCommand
        cmd.Parameters.AddWithValue("@part", part)
        cmd.CommandText = "SELECT min_qty, current_qty, Qty_in_order from inventory.inventory_qty where part_Name = @part"
        cmd.Connection = Login.Connection
        Dim reader As MySqlDataReader
        reader = cmd.ExecuteReader

        If reader.HasRows Then
            While reader.Read
                min = reader(0).ToString
                onhand = reader(1).ToString
                qty_transit = reader(2).ToString
            End While
        End If

        reader.Close()

        demand = Inventory_manage.cal_demand(part)

        If (onhand + qty_transit - demand) >= min Then
            Is_above_min = True
        End If

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '-- generate a ASM BOM --------
        Try
            Dim my_assem As New List(Of String)

            Dim dimen_table = New DataTable
            dimen_table.Columns.Add("Assembly", GetType(String))
            dimen_table.Columns.Add("qty", GetType(String))

            '----- datatable ready to be copy to mysql table mr_data
            Dim dimen_table2 = New DataTable
            dimen_table2.Columns.Add("Part_No", GetType(String))
            dimen_table2.Columns.Add("Description", GetType(String))
            dimen_table2.Columns.Add("Manufacturer", GetType(String))
            dimen_table2.Columns.Add("Vendor", GetType(String))
            dimen_table2.Columns.Add("Price", GetType(String))
            dimen_table2.Columns.Add("Qty", GetType(Double))
            dimen_table2.Columns.Add("MFG", GetType(String))
            dimen_table2.Columns.Add("part_status", GetType(String))
            dimen_table2.Columns.Add("part_Type", GetType(String))
            dimen_table2.Columns.Add("asm", GetType(String))

            Dim is_assem As Boolean : is_assem = False

            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    my_assem.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()

            '----- check if there is at least one assembly -----
            For i = 0 To MO_grid.Rows.Count - 1
                If my_assem.Contains(MO_grid.Rows(i).Cells(0).Value.ToString) = True Then
                    dimen_table.Rows.Add(MO_grid.Rows(i).Cells(0).Value.ToString, MO_grid.Rows(i).Cells(4).Value.ToString)
                    is_assem = True
                End If
            Next
            '-----------------------------------------

            If is_assem = True Then

                '----- ASM Bom name ----------
                Dim idb As Integer : idb = 1
                Dim asm_bom_name As String : asm_bom_name = "ASM_BOM_"
                Dim today_d As String : today_d = System.DateTime.Now.ToString
                asm_bom_name = asm_bom_name & today_d
                asm_bom_name = asm_bom_name.Replace("/", "-")
                asm_bom_name = asm_bom_name.Replace(":", "-")
                asm_bom_name = asm_bom_name.Replace(" ", "_")
                '------------------------------------

                Dim cmd21 As New MySqlCommand
                cmd21.Parameters.AddWithValue("@job", "ADA_INV")
                cmd21.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"
                cmd21.Connection = Login.Connection
                Dim reader21 As MySqlDataReader
                reader21 = cmd21.ExecuteReader

                If reader21.HasRows Then
                    While reader21.Read
                        idb = reader21(0)
                    End While
                End If

                reader21.Close()


                '-------- create mr ---------------------------
                Dim main_cmd As New MySqlCommand
                main_cmd.Parameters.AddWithValue("@mr_name", asm_bom_name)
                main_cmd.Parameters.AddWithValue("@created_by", current_user)
                main_cmd.Parameters.AddWithValue("@released", "Y")
                main_cmd.Parameters.AddWithValue("@released_by", current_user)
                main_cmd.Parameters.AddWithValue("@job", "ADA_INV")
                main_cmd.Parameters.AddWithValue("@id_bom", idb + 1)
                main_cmd.Parameters.AddWithValue("@need_date", Date.Today)
                main_cmd.Parameters.AddWithValue("@finished", "Y")
                main_cmd.CommandText = "INSERT INTO Material_Request.mr(mr_name, Date_Created , created_by, released, released_by, release_date, job, id_bom, finished, BOM_type, need_date) VALUES (@mr_name, now(), @created_by, @released, @released_by, now(), @job, @id_bom, @finished, 'ASM', @need_date)"
                main_cmd.Connection = Login.Connection
                main_cmd.ExecuteNonQuery()


                For i = 0 To dimen_table.Rows.Count - 1

                    Dim adv_n As String : adv_n = "sdfsdfas"
                    Dim cmd_a As New MySqlCommand
                    cmd_a.Parameters.AddWithValue("@assem", dimen_table.Rows(i).Item(0))
                    cmd_a.CommandText = "Select ADV_Number from devices where Legacy_ADA_Number = @assem"
                    cmd_a.Connection = Login.Connection

                    Dim reader_k As MySqlDataReader
                    reader_k = cmd_a.ExecuteReader


                    If reader_k.HasRows Then
                        While reader_k.Read
                            adv_n = reader_k(0)
                        End While
                    End If

                    reader_k.Close()


                    Dim cmd_pd As New MySqlCommand
                    cmd_pd.Parameters.AddWithValue("@device", adv_n)
                    cmd_pd.CommandText = "SELECT p1.Part_Name, p1.Part_Description, p1.Manufacturer,  p1.Primary_Vendor, adv.Qty, p1.MFG_type, p1.Part_Status, p1.Part_Type from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = @device"
                    cmd_pd.Connection = Login.Connection
                    Dim readerv As MySqlDataReader
                    readerv = cmd_pd.ExecuteReader

                    If readerv.HasRows Then
                        While readerv.Read
                            dimen_table2.Rows.Add(readerv(0).ToString, readerv(1).ToString, readerv(2).ToString, readerv(3).ToString, 0, readerv(4).ToString * dimen_table.Rows(i).Item(1), readerv(5).ToString, readerv(6).ToString, readerv(7).ToString, dimen_table.Rows(i).Item(0))  'add a new row
                        End While
                    End If

                    readerv.Close()
                Next

                For i = 0 To dimen_table2.Rows.Count - 1
                    dimen_table2.Rows(i).Item(4) = Form1.Get_Latest_Cost(Login.Connection, dimen_table2.Rows(i).Item(0), dimen_table2.Rows(i).Item(3))
                Next



                For i = 0 To dimen_table2.Rows.Count - 1

                    Dim Create_cmd6 As New MySqlCommand
                    Create_cmd6.Parameters.Clear()
                    Create_cmd6.Parameters.AddWithValue("@mr_name", asm_bom_name)
                    Create_cmd6.Parameters.AddWithValue("@Part_No", If(dimen_table2.Rows(i).Item(0) Is Nothing, "unknown part", dimen_table2.Rows(i).Item(0)))
                    Create_cmd6.Parameters.AddWithValue("@Description", If(dimen_table2.Rows(i).Item(1) Is Nothing, "", dimen_table2.Rows(i).Item(1)))
                    Create_cmd6.Parameters.AddWithValue("@ADA_Number", "")
                    Create_cmd6.Parameters.AddWithValue("@Manufacturer", If(dimen_table2.Rows(i).Item(2) Is Nothing, "", dimen_table2.Rows(i).Item(2)))
                    Create_cmd6.Parameters.AddWithValue("@Vendor", If(dimen_table2.Rows(i).Item(3) Is Nothing, "", dimen_table2.Rows(i).Item(3)))
                    Create_cmd6.Parameters.AddWithValue("@Price", If(dimen_table2.Rows(i).Item(4) Is Nothing, "", dimen_table2.Rows(i).Item(4)))
                    Create_cmd6.Parameters.AddWithValue("@Qty", If(dimen_table2.Rows(i).Item(5) Is Nothing, "", dimen_table2.Rows(i).Item(5)))
                    Create_cmd6.Parameters.AddWithValue("@subtotal", dimen_table2.Rows(i).Item(4) * dimen_table2.Rows(i).Item(5))
                    Create_cmd6.Parameters.AddWithValue("@mfg_type", If(dimen_table2.Rows(i).Item(6) Is Nothing, "", dimen_table2.Rows(i).Item(6)))
                    Create_cmd6.Parameters.AddWithValue("@assembly", If(dimen_table2.Rows(i).Item(9) Is Nothing, "", dimen_table2.Rows(i).Item(9)))
                    Create_cmd6.Parameters.AddWithValue("@Notes", "")
                    Create_cmd6.Parameters.AddWithValue("@Part_status", If(dimen_table2.Rows(i).Item(7) Is Nothing, "", dimen_table2.Rows(i).Item(7)))
                    Create_cmd6.Parameters.AddWithValue("@Part_type", If(dimen_table2.Rows(i).Item(8) Is Nothing, "", dimen_table2.Rows(i).Item(8)))
                    Create_cmd6.Parameters.AddWithValue("@need_by_date", "")


                    Create_cmd6.CommandText = "INSERT INTO Material_Request.mr_data(mr_name, Part_No, Description,ADA_Number, Manufacturer, Vendor,  Price, Qty, subtotal, mfg_type, Assembly_name, Part_status, Part_type, Notes,  need_by_date, released, latest_r) VALUES (@mr_name, @Part_No, @Description, @ADA_Number, @Manufacturer, @Vendor, @Price, @Qty, @subtotal, @mfg_type, @assembly, @Part_status, @Part_type, @Notes,  @need_by_date, 'Y', 'x')"
                    Create_cmd6.Connection = Login.Connection
                    Create_cmd6.ExecuteNonQuery()

                Next


                '---- enter data to ASM BOM table ---

                For j = 0 To dimen_table.Rows.Count - 1
                    Dim main_cmd7 As New MySqlCommand
                    main_cmd7.Parameters.Clear()
                    main_cmd7.Parameters.AddWithValue("@mr_name", asm_bom_name)
                    main_cmd7.Parameters.AddWithValue("@asm", dimen_table.Rows(j).Item(0))
                    main_cmd7.Parameters.AddWithValue("@qty", dimen_table.Rows(j).Item(1))
                    main_cmd7.CommandText = "INSERT INTO Tracking_Reports.ASM_BOM(mr_name, ASM, qty_needed) VALUES (@mr_name,@asm,@qty)"
                    main_cmd7.Connection = Login.Connection
                    main_cmd7.ExecuteNonQuery()
                Next

                '-----------------------------

                If enable_mess = True Then
                    'write mail
                    Dim mail_n As String : mail_n = "Assembly Request has been Released for Project ADA_INV" & vbCrLf & vbCrLf _
                         & "Released by: " & current_user & vbCrLf _
                         & "Released Date: " & Now & vbCrLf _
                         & "Project: ADA_INV" & vbCrLf _
                         & "BOM File Name: " & asm_bom_name & vbCrLf

                    Call Sent_mail.Sent_multiple_emails("Inventory", "Assembly Request has been Released for Project ADA_INV", mail_n)
                End If


                MessageBox.Show("ASM BOM was created succesfully")

            Else
                MessageBox.Show("No Assemblies found!")

            End If


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub
End Class