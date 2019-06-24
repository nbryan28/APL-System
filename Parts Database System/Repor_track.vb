Imports MySql.Data.MySqlClient
Imports System.Reflection

Public Class Repor_track
    Private Sub Label3_Click(sender As Object, e As EventArgs)
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        Try
            ComboBox2.Items.Clear()
            Dim id As Integer : id = 1
            Dim cmd_v As New MySqlCommand
            cmd_v.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd_v.CommandText = "select count(distinct id_bom) from Material_Request.mr where job = @job"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    id = readerv(0)
                End While
            End If

            readerv.Close()

            '------- fill combobox ---------
            For i = 1 To id
                Dim cmd_v1 As New MySqlCommand
                cmd_v1.Parameters.Clear()
                cmd_v1.Parameters.AddWithValue("@job", ComboBox1.Text)
                cmd_v1.Parameters.AddWithValue("@id_bom", i)
                cmd_v1.CommandText = "select mr_name from Material_Request.mr where job = @job and id_bom = @id_bom order by release_date desc limit 1"

                cmd_v1.Connection = Login.Connection
                Dim readerv1 As MySqlDataReader
                readerv1 = cmd_v1.ExecuteReader

                If readerv1.HasRows Then
                    While readerv1.Read
                        ComboBox2.Items.Add(readerv1(0))
                    End While
                End If

                readerv1.Close()
            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub Repor_track_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call EnableDoubleBuffered(orders_grid)
        '---- laod active jobs
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

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        '============ COPY MENU ================
        If orders_grid.CurrentCell.Value.ToString <> String.Empty Then
            Clipboard.SetText(orders_grid.CurrentCell.Value.ToString)
        Else
            Clipboard.Clear()
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged

        Dim mr As String : mr = "none"
        Dim id_box As String : id_box = 1

        If Not ComboBox2.SelectedItem Is Nothing Then
            mr = ComboBox2.SelectedItem.ToString
        End If

        orders_grid.Rows.Clear()

        Try
            Dim cmd4 As New MySqlCommand
            cmd4.Parameters.AddWithValue("@mr", mr)
            cmd4.CommandText = "SELECT id_bom from Material_Request.mr where mr_name = @mr"
            cmd4.Connection = Login.Connection
            Dim reader4 As MySqlDataReader
            reader4 = cmd4.ExecuteReader

            If reader4.HasRows Then
                While reader4.Read
                    id_box = reader4(0)
                End While
            End If

            reader4.Close()

            '--------------------
            Dim cmd41 As New MySqlCommand
            cmd41.Parameters.AddWithValue("@id_bom", id_box)
            cmd41.Parameters.AddWithValue("@job", ComboBox1.Text)
            cmd41.CommandText = "SELECT Part_No, qty_needed, PO, es_date_of_arrival, mfg, primary_vendor, cost from Tracking_Reports.my_tracking_reports where job = @job and id_bom = @id_bom"
            cmd41.Connection = Login.Connection
            Dim reader41 As MySqlDataReader
            reader41 = cmd41.ExecuteReader

            If reader41.HasRows Then
                Dim i As Integer : i = 0
                While reader41.Read
                    orders_grid.Rows.Add(New String() {})
                    orders_grid.Rows(i).Cells(0).Value = reader41(0).ToString
                    orders_grid.Rows(i).Cells(1).Value = reader41(4).ToString
                    orders_grid.Rows(i).Cells(2).Value = reader41(5).ToString
                    orders_grid.Rows(i).Cells(3).Value = reader41(6).ToString
                    orders_grid.Rows(i).Cells(4).Value = reader41(1).ToString
                    orders_grid.Rows(i).Cells(5).Value = reader41(2).ToString
                    orders_grid.Rows(i).Cells(6).Value = reader41(3).ToString
                    i = i + 1
                End While

            End If

            reader41.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer

        For Each row As DataGridViewRow In orders_grid.Rows
            If String.Compare(row.Cells.Item("Column10").Value.ToString, TextBox1.Text) = 0 Then
                rowindex = row.Index
                orders_grid.CurrentCell = orders_grid.Rows(rowindex).Cells(0)
                found_po = True
                Exit For
            End If
        Next
        If found_po = False Then
            MsgBox(" Part not found! ")
        End If
    End Sub


    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub
End Class