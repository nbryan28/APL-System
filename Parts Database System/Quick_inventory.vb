Imports MySql.Data.MySqlClient

Public Class Quick_inventory
    Private Sub Quick_inventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' ======================  Start filling DataViewGrid1 with PARTS data ============================
        Try
            Dim table As New DataTable
            Dim adapter As New MySqlDataAdapter("SELECT part_name, description, current_qty, Qty_in_order as Qty_in_Transit, es_date_of_arrival from inventory.inventory_qty order by part_name", Login.Connection)


            adapter.Fill(table)     'DataViewGrid1 fill
            DataGridView1.DataSource = table
            '   Form1.Specs_Table.DataSource = table   'Specifications table fill


            DataGridView1.Columns(0).Width = 480
            DataGridView1.Columns(1).Width = 580
            DataGridView1.Columns(2).Width = 300
            DataGridView1.Columns(3).Width = 200
            DataGridView1.Columns(4).Width = 400

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In DataGridView1.Rows
            If String.IsNullOrEmpty(row.Cells.Item("part_name").Value) = False Then
                If String.Compare(row.Cells.Item("part_name").Value.ToString, TextBox1.Text, True) = 0 Then
                    rowindex = row.Index
                    DataGridView1.CurrentCell = DataGridView1.Rows(rowindex).Cells(0)
                    found_po = True
                    Exit For
                End If

            End If
        Next

        If found_po = False Then
            MsgBox("Part not found in Inventory!")
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If DataGridView1.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(DataGridView1.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub
End Class