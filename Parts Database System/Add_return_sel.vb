Imports MySql.Data.MySqlClient

Public Class Add_return_sel
    Private Sub Add_return_sel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Part_Description, Manufacturer, Part_Type,  Primary_Vendor, MFG_type, Part_Status from parts_table order by Part_Name", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            allParts.DataSource = table

            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210

            job_box.Items.Clear()
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT distinct job from Material_Request.mr where job is not null order by job"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    job_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'show all parts
        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT Part_Name, Part_Description, Manufacturer, Part_Type, Primary_Vendor, MFG_type, Part_Status from parts_table order by Part_Name", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            allParts.DataSource = table

            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210


        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '============== Partial match ============================
        Try
            Dim cmdstr As String : cmdstr = "SELECT Part_Name as Part_Number, Part_Description, Manufacturer, Part_Type, Primary_Vendor, MFG_type, Part_status  from parts_table where Part_Name LIKE  @search"
            Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
            cmd.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")

            Dim table_s As New DataTable
            Dim adapter_s As New MySqlDataAdapter(cmd)

            adapter_s.Fill(table_s)
            allParts.DataSource = table_s

            allParts.Columns(0).Width = 480
            allParts.Columns(1).Width = 480
            allParts.Columns(2).Width = 210
            allParts.Columns(3).Width = 210
            allParts.Columns(4).Width = 210
            allParts.Columns(5).Width = 210
            allParts.Columns(6).Width = 210

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Add_return.addr_grid.Rows.Add({job_box.Text, "", qty_box.Text, part_label.Text, "", "", Form1.Get_Latest_Cost(Login.Connection, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(0).Value, allParts.Rows(allParts.CurrentCell.RowIndex).Cells(4).Value)})
    End Sub

    Private Sub allParts_DoubleClick(sender As Object, e As EventArgs) Handles allParts.DoubleClick
        Dim index_k = allParts.CurrentCell.RowIndex

        part_label.Text = allParts.Rows(index_k).Cells(0).Value

    End Sub
End Class