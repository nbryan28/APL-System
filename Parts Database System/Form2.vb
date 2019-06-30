
Imports MySql.Data.MySqlClient

Public Class Login

    Public Connection As MySqlConnection

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'secret_word = TextBox1.Text
        'Visible = False

        ''Open a mysql connection with the password specified
        ''IP Address 192.168.68.112 of my computer or 192.168.68.34 engraving room
        'Try
        '    Connection = New MySqlConnection("datasource= 192.168.68.112;port=3306;username=root;password=" & secret_word & ";database=parts;Allow User Variables=True")
        '    Connection.Open()
        'Catch ex As Exception
        '    MessageBox.Show("Wrong password!!!")
        'End Try

        ''check if connection is open
        'If (Connection.State = ConnectionState.Open) Then

        '    ok_press = True
        '    Form1.DataGridView1.EnableHeadersVisualStyles = False

        '    ' ======================  Start filling DataViewGrid1 with PARTS data ============================
        '    Dim table As New DataTable
        '    Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Primary_Vendor from Parts_Table order by Part_Name", Connection)
        '    'Dim adapter2 As New MySqlDataAdapter("SELECT @row:=@row+1 as Alternate, ADA_id as id, ADA_Part_Name, SUM(Unit_Cost) as Total_Cost from ada_parts, (SELECT @row := 0) m where ADA_Part_Name = '" & ADA_name & "' group by ADA_id", Connection)

        '    adapter.Fill(table)     'DataViewGrid1 fill
        '    Form1.DataGridView1.DataSource = table
        '    '   Form1.Specs_Table.DataSource = table   'Specifications table fill

        '    'Setting Columns size for Parts Datagrid
        '    For i = 0 To Form1.DataGridView1.ColumnCount - 1
        '        With Form1.DataGridView1.Columns(i)
        '            .Width = 380
        '        End With
        '    Next i

        '    'Setting Columns size for Specifications table

        '    '  For i = 0 To Form1.Specs_Table.ColumnCount - 1
        '    'With Form1.Specs_Table.Columns(i)
        '    '.Width = 380
        '    'End With
        '    '  Next i


        '    ' ======================  Start filling Vendor Table with vendors data =============================
        '    Dim table3 As New DataTable
        '    Dim adapter3 As New MySqlDataAdapter("SELECT Part_name, vendor_name, Cost, Purchase_date from vendors_table order by Part_Name", Connection)

        '    adapter3.Fill(table3)     'Specifications table fill
        '    Form1.Vendors_grid.DataSource = table3


        '    '======================== Start filling KITS TABLE with kits ==========================
        '    Dim table_kits As New DataTable
        '    Dim adapter_kits As New MySqlDataAdapter("SELECT KIT_Name, Legacy_ADA_Number as ADA_Number, Description from KITS", Connection)

        '    adapter_kits.Fill(table_kits)     'kits grid fill
        '    Form1.Kit_Table.DataSource = table_kits
        '    Form1.Kit_Table.DataSource = table_kits   'kits table fill

        '    'Setting Columns size for KITS Datagrid
        '    For i = 0 To Form1.Kit_Table.ColumnCount - 1
        '        With Form1.Kit_Table.Columns(i)
        '            .Width = 550
        '        End With
        '    Next i


        '    '===========================Start filling DEVICE TABLE with devices =============================
        '    Dim table_dev As New DataTable
        '    Dim adapter_dev As New MySqlDataAdapter("SELECT DEVICE_Name, Legacy_ADA_Number as ADA_Number, Description, Material_Cost, Labor_Cost, (Material_Cost + Labor_Cost) as Cost from DEVICES", Connection)

        '    adapter_dev.Fill(table_dev)     'Device fill
        '    Form1.Device_Table.DataSource = table_dev
        '    Form1.Device_Table.DataSource = table_dev   'devices table fill

        '    'Setting Columns size for devices Datagrid
        '    For i = 0 To Form1.Device_Table.ColumnCount - 1
        '        With Form1.Device_Table.Columns(i)
        '            .Width = 330
        '        End With
        '    Next i

        '    '========================== Start filling ATRONIX_NUMBER with atronix_numbers =============================

        '    'Dim table_an As New DataTable
        '    'Dim adapter_an As New MySqlDataAdapter("(SELECT ACN_Number as Atronix_Number, Part_Name, Legacy_ADA from acn) UNION ALL (SELECT AKN_Number as Atronix_Number, Part_Name, Legacy_ADA from akn) UNION ALL SELECT ADV_Number as Atronix_Number, Part_Name, Legacy_ADA from adv", Connection)

        '    'adapter_an.Fill(table_an)     'Atronix Number datagrid
        '    'Form1.Cross_Table.DataSource = table_an
        '    'Form1.Cross_Table.DataSource = table_an   'Atronix Number datagrid table fill

        '    ''Setting Columns size for atronix number Datagrid
        '    'For i = 0 To Form1.Cross_Table.ColumnCount - 1
        '    '    With Form1.Cross_Table.Columns(i)
        '    '        .Width = 430
        '    '    End With
        '    'Next i


        '    '============= Start filling ADA TABLE ===============================

        '    Dim table_ada As New DataTable

        '    Dim adapter_ada As New MySqlDataAdapter("select * from ((Select Legacy_ADA_Number as ADA_Number, Part_Name as Reference_Name from parts_table)  UNION ALL (Select Legacy_ADA_Number as ADA_Number, KIT_Name as Reference_Name from kits)  UNION ALL (Select Legacy_ADA_Number as ADA_Number, DEVICE_Name as Reference_Name from Devices) ) as t order by ADA_Number desc", Connection)

        '    adapter_ada.Fill(table_ada)
        '    Form1.ADA_grid.DataSource = table_ada
        '    Form1.ADA_grid.Columns(0).Width = 720
        '    Form1.ADA_grid.Columns(1).Width = 850




        '    '============================= Start filling KIT/DEVICE Visualizer table ==============================

        '    Dim table_v As New DataTable
        '    Dim adapter_v As New MySqlDataAdapter("(SELECT KIT_Name as DEVICE_KIT_Name from Kits) UNION ALL (SELECT DEVICE_Name as DEVICE_KIT_Name from Devices)", Connection)

        '    adapter_v.Fill(table_v)     'Visualizer 
        '    Form1.Device_Grid.DataSource = table_v
        '    Form1.Device_Grid.DataSource = table_v   'Visualizer table fill

        '    'Setting Columns size for Visualizer Datagrid
        '    For i = 0 To Form1.Device_Grid.ColumnCount - 1
        '        With Form1.Device_Grid.Columns(i)
        '            .Width = 700
        '        End With
        '    Next i



        '    'Setting Columns size for vendor datagrid
        '    For i = 0 To Form1.Vendors_grid.ColumnCount - 1
        '        With Form1.Vendors_grid.Columns(i)
        '            .Width = 460
        '        End With
        '    Next i

        '    Form1.Vendors_grid.Columns(3).Width = 150

        '    'Add Context Menu to Parts Table (Datagridview1 in form1)
        '    ' Form1.DataGridView1.Columns(0).ContextMenuStrip = Form1.menu_BOM


        '    'Populate Part Type Combo boxes
        '    Try
        '        Dim cmd_v As New MySqlCommand
        '        cmd_v.CommandText = "SELECT Part_Type from Parts_Type_table"

        '        cmd_v.Connection = Connection
        '        Dim readerv As MySqlDataReader
        '        readerv = cmd_v.ExecuteReader

        '        If readerv.HasRows Then
        '            While readerv.Read
        '                Form1.PartTypebox1.Items.Add(readerv(0))
        '            End While
        '        End If

        '        readerv.Close()

        '    Catch ex As Exception
        '        MessageBox.Show(ex.ToString)
        '    End Try

        '    'CHECK VERSION CONTROL
        '    Try
        '        version_c = "AAPL v0.9"
        '        Dim cmd_ver As New MySqlCommand
        '        cmd_ver.CommandText = "SELECT version from version_control"

        '        cmd_ver.Connection = Connection
        '        Dim readerv As MySqlDataReader
        '        readerv = cmd_ver.ExecuteReader

        '        If readerv.HasRows Then
        '            While readerv.Read
        '                If String.Equals(version_c, readerv(0).ToString) = False Then

        '                    Dim result As DialogResult = MessageBox.Show("A new version is out! Please click yes to update AAPL", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message
        '                    Dim proc As New System.Diagnostics.Process()
        '                    Dim appPath As String = Application.StartupPath()

        '                    If (result = DialogResult.Yes) Then
        '                        '--------- open update ---
        '                        Try
        '                            proc = Process.Start(appPath & "\Updater.exe")
        '                        Catch ex As Exception
        '                            MessageBox.Show("Cannot access the server")
        '                        End Try
        '                        '-----------------------------------
        '                    End If
        '                End If
        '            End While
        '        End If

        '        readerv.Close()

        '    Catch ex As Exception
        '        Message
        '    End Try

        'End If

    End Sub

    'Connect to database
    Sub Access_database()

        'IP Address my laptop 192.168.68.112 of my computer or 192.168.68.34 engraving room, APL server 10.21.2.8, user:root pass:aapl_1369  (old: AAPL server 192.168.69.138 pass 113133)
        Try
            ' Connection = New MySqlConnection("datasource=10.21.2.8;port=3306;username=root;password=aapl_1369;database=parts;Allow User Variables=True")
            Connection = New MySqlConnection("datasource=localhost ;port=3306;username=root;password=atronixatl;database=parts;Allow User Variables=True")
            Connection.Open()
        Catch ex As Exception
            MessageBox.Show("Can't connect!!!")
        End Try

        'check if connection is open
        If (Connection.State = ConnectionState.Open) Then

            ok_press = True

            Form1.DataGridView1.EnableHeadersVisualStyles = False
            Form1.TabPage1.BackColor = Color.FromArgb(61, 60, 78)

            Form1.fc_grid.Rows.Clear()


            ' ======================  Start filling DataViewGrid1 with PARTS data ============================
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name as Part_Number, Manufacturer, Part_Description, Legacy_ADA_Number as ADA_Number, Part_Status, Part_Type, Notes, Primary_Vendor, MFG_type from parts_table order by Part_Name", Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            Form1.DataGridView1.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill

            'Setting Columns size for Parts Datagrid
            For i = 0 To Form1.DataGridView1.ColumnCount - 1
                With Form1.DataGridView1.Columns(i)
                    .Width = 200
                End With
            Next i

            Form1.DataGridView1.Columns(0).Width = 480
            Form1.DataGridView1.Columns(2).Width = 480


            ' ======================  Start filling Vendor Table with vendors data =============================
            Dim table3 As New DataTable
            Dim adapter3 As New MySqlDataAdapter("SELECT Part_name, vendor_name, Cost, Purchase_date from vendors_table order by Part_Name", Connection)

            adapter3.Fill(table3)     'Specifications table fill
            Form1.Vendors_grid.DataSource = table3


            'Setting Columns size for vendor datagrid
            For i = 0 To Form1.Vendors_grid.ColumnCount - 1
                With Form1.Vendors_grid.Columns(i)
                    .Width = 460
                End With
            Next i

            Form1.Vendors_grid.Columns(3).Width = 150


            '======================== Start filling KITS TABLE with kits ==========================
            'Dim table_kits As New DataTable
            'Dim adapter_kits As New MySqlDataAdapter("SELECT Legacy_ADA_Number as ADA_Number, Description from kits order by Legacy_ADA_Number", Connection)

            'adapter_kits.Fill(table_kits)     'kits grid fill
            'Form1.Kit_Table.DataSource = table_kits
            'Form1.Kit_Table.DataSource = table_kits   'kits table fill

            ''Setting Columns size for KITS Datagrid
            'For i = 0 To Form1.Kit_Table.ColumnCount - 1
            '    With Form1.Kit_Table.Columns(i)
            '        .Width = 850
            '    End With
            'Next i

            '================ feature code ====================================

            '--------- load solutions ---------------
            Dim cmd_vp As New MySqlCommand
            cmd_vp.CommandText = "SELECT solution_name from quote_table.feature_solutions order by solution_name"

            cmd_vp.Connection = Connection
            Dim readervp As MySqlDataReader
            readervp = cmd_vp.ExecuteReader

            If readervp.HasRows Then
                While readervp.Read
                    Form1.sol_box.Items.Add(readervp(0))
                End While
            End If

            readervp.Close()

            Form1.sol_box.Text = "EIP-MS/EIP-RIO" '"SWDMS/EIPRIO"


            '--- load fc table ------

            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@solution", "EIP-MS/EIP-RIO")
            cmd4.CommandText = "Select Feature_code, description, type, specific_type, labor_cost, bulk_cost from quote_table.feature_codes where solution = @solution"
            cmd4.Connection = Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    Form1.fc_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString, reader4(3).ToString, If(IsDBNull(reader4(4)), "0", reader4(4).ToString), If(IsDBNull(reader4(5)), "0", reader4(5).ToString), 0})
                End While
            End If

            reader4.Close()

            Call Form1.update_total_fc()


            '===========================Start filling DEVICE TABLE with devices =============================
            Dim table_dev As New DataTable
            Dim adapter_dev As New MySqlDataAdapter("SELECT Legacy_ADA_Number as ADA_Number, Description, Bulk_Cost, Labor_Cost from devices order by Legacy_ADA_Number", Connection)

            adapter_dev.Fill(table_dev)     'Device fill
            Form1.Device_Table.DataSource = table_dev
            Form1.Device_Table.DataSource = table_dev   'devices table fill

            'Setting Columns size for devices Datagrid
            For i = 0 To Form1.Device_Table.ColumnCount - 1
                With Form1.Device_Table.Columns(i)
                    .Width = 530
                End With
            Next i


            '===================================== Start filling ADA TABLE ===============================

            'Dim table_ada As New DataTable
            'Dim adapter_ada As New MySqlDataAdapter("select * from ((Select Legacy_ADA_Number as ADA_Number, Part_Name as Reference_Name from parts_table where Legacy_ADA_Number not like '%none%')  UNION ALL (Select Legacy_ADA_Number as ADA_Number, KIT_Name as Reference_Name from kits)  UNION ALL (Select Legacy_ADA_Number as ADA_Number, DEVICE_Name as Reference_Name from Devices) ) as t order by ADA_Number desc", Connection)

            'adapter_ada.Fill(table_ada)
            'Form1.ADA_grid.DataSource = table_ada
            'Form1.ADA_grid.Columns(0).Width = 720
            'Form1.ADA_grid.Columns(1).Width = 850


            '===============================================================================================

            'Populate Part Type Combo boxes
            Try
                Dim cmd_v As New MySqlCommand
                cmd_v.CommandText = "SELECT Part_Type from parts_type_table"

                cmd_v.Connection = Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_v.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        Form1.PartTypebox1.Items.Add(readerv(0))  'change
                    End While
                End If

                readerv.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            '----------  CHECK VERSION CONTROL ----------
            Try
                version_c = "APL v0.99br"
                Dim cmd_ver As New MySqlCommand
                cmd_ver.CommandText = "SELECT version from version_control"

                cmd_ver.Connection = Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_ver.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        If String.Equals(version_c, readerv(0).ToString) = False Then

                            Dim result As DialogResult = MessageBox.Show("A new version is out! Please click yes to update APL", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) 'confirmation message
                            Dim proc As New System.Diagnostics.Process()
                            Dim appPath As String = Application.StartupPath()

                            If (result = DialogResult.Yes) Then
                                '--------- open update ---
                                Try
                                    proc = Process.Start(appPath & "\Updater.exe")
                                Catch ex As Exception
                                    MessageBox.Show("Cannot access the server")
                                End Try
                                '-----------------------------------
                            End If
                        End If
                    End While
                End If

                readerv.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            ' ---------------------------- Populate  Combo boxes ----------------------------------
            Try
                Dim cmd_v As New MySqlCommand
                cmd_v.CommandText = "SHOW TABLES in jobs"

                cmd_v.Connection = Connection
                Dim readerv As MySqlDataReader
                readerv = cmd_v.ExecuteReader

                If readerv.HasRows Then
                    While readerv.Read
                        Form1.projects_box.Items.Add(readerv(0).ToString.Substring(readerv(0).ToString.IndexOf("_") + 1))
                    End While
                End If

                readerv.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

            current_user = ""


        End If

        Form1.TabControl1.TabPages.Remove(Form1.TabPage10)
        Form1.WindowState = FormWindowState.Maximized


    End Sub


End Class