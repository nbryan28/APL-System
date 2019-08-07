Imports MySql.Data.MySqlClient

Public Class Open_package
    Private Sub Open_package_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Items.Clear()
        open_grid.Rows.Clear()

        Try

            Dim check_cmd2 As New MySqlCommand
            check_cmd2.CommandText = "select distinct job from Material_Request.mr where released is null and job is not null order by job"
            check_cmd2.Connection = Login.Connection
            check_cmd2.ExecuteNonQuery()

            Dim reader2 As MySqlDataReader
            reader2 = check_cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    ComboBox1.Items.Add(reader2(0))
                End While
            End If

            reader2.Close()


            '--------------------------
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select mr_name, Date_Created, created_by, job from Material_Request.mr where released is null and BOM_Type = 'MB' order by mr_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(i).Cells(1).Value = reader(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader(1)
                    open_grid.Rows(i).Cells(3).Value = reader(2).ToString
                    open_grid.Rows(i).Cells(4).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub



    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        '---- filter boms -------
        Try
            open_grid.Rows.Clear()
            '--------------------------
            Dim check_cmd As New MySqlCommand
            check_cmd.Parameters.AddWithValue("@job", ComboBox1.SelectedItem.ToString)
            check_cmd.CommandText = "select mr_name, Date_Created, created_by, job from Material_Request.mr where released is null and BOM_Type = 'MB' and job = @job order by mr_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(i).Cells(1).Value = reader(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader(1)
                    open_grid.Rows(i).Cells(3).Value = reader(2).ToString
                    open_grid.Rows(i).Cells(4).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- show all

        open_grid.Rows.Clear()
        Try

            '--------------------------
            Dim check_cmd As New MySqlCommand
            check_cmd.CommandText = "select mr_name, Date_Created, created_by, job from Material_Request.mr where released is null and BOM_Type = 'MB' order by mr_name"
            check_cmd.Connection = Login.Connection
            check_cmd.ExecuteNonQuery()

            Dim reader As MySqlDataReader
            reader = check_cmd.ExecuteReader

            If reader.HasRows Then
                Dim i As Integer : i = 0
                While reader.Read
                    open_grid.Rows.Add(New String() {})
                    open_grid.Rows(i).Cells(1).Value = reader(0).ToString
                    open_grid.Rows(i).Cells(2).Value = reader(1)
                    open_grid.Rows(i).Cells(3).Value = reader(2).ToString
                    open_grid.Rows(i).Cells(4).Value = reader(3).ToString
                    i = i + 1
                End While
            End If

            reader.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub open_grid_DoubleClick(sender As Object, e As EventArgs) Handles open_grid.DoubleClick

        If open_grid.Rows.Count > 0 Then

            Dim index_k = open_grid.CurrentCell.RowIndex
            Dim mr_name As String : mr_name = open_grid.Rows(index_k).Cells(1).Value
            Dim job As String : job = open_grid.Rows(index_k).Cells(4).Value

            Call BOM_types.Open_BOM_package(mr_name, job)
            Me.Visible = False
        End If
    End Sub

    Private Sub Open_package_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = False Then
            Me.Close()
        End If
    End Sub
End Class