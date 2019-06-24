Imports MySql.Data.MySqlClient


'-------- NOTE: THE Remote_IO table has a column called SOLUTIOM instead of solutiom ------------

Public Class RIO_wiz

    Public trigger As Boolean
    Public cell_ch As Boolean

    Private Sub RIO_wiz_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        trigger = False
        RIO_grid.Rows.Clear()
        rio_grid2.Rows.Clear()

        RIO_grid.Rows.Add(New String() {"Required", myQuote.inputs_io, myQuote.outputs_io, myQuote.motion_io})
        RIO_grid.Rows.Add(New String() {"Available"})



        Try

            '------ sol box fill--------------
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT solution_name from quote_table.feature_solutions order by solution_name desc"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    sol_box.Items.Add(readerv(0))
                End While
            End If

            readerv.Close()
            '-----------------------------

            sol_box.Text = "SWDMS/EIPRIO"

            Dim cmd_panel As New MySqlCommand
            cmd_panel.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd_panel.CommandText = "SELECT feature_code from quote_table.Remote_IO where solutiom = @sol"
            cmd_panel.Connection = Login.Connection
            Dim reader_panel As MySqlDataReader
            reader_panel = cmd_panel.ExecuteReader

            If reader_panel.HasRows Then
                While reader_panel.Read
                    rio_grid2.Rows.Add(New String() {reader_panel(0)})
                End While
            End If

            reader_panel.Close()



        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        For i = 0 To rio_grid2.Rows.Count - 1
            Call Rio_needed(i)
        Next

        cell_ch = True
        trigger = True
        Call eval_ava()

    End Sub

    Function Max_number(num1 As Double, num2 As Double, num3 As Double) As Double

        'Return the max number of the parameters passed
        Max_number = num1

        Dim array_t(4) As Double
        array_t(0) = num1
        array_t(1) = num2
        array_t(2) = num3

        Max_number = array_t.Max

    End Function

    Private Sub sol_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles sol_box.SelectedValueChanged

        If trigger = True Then

            cell_ch = False

            RIO_grid.Rows.Clear()
            rio_grid2.Rows.Clear()

            RIO_grid.Rows.Add(New String() {"Required", myQuote.inputs_io, myQuote.outputs_io, myQuote.motion_io})
            RIO_grid.Rows.Add(New String() {"Available"})


            Try

                Dim cmd_panel As New MySqlCommand
                cmd_panel.Parameters.AddWithValue("@sol", sol_box.Text)
                cmd_panel.CommandText = "SELECT feature_code from quote_table.Remote_IO where solutiom = @sol"
                cmd_panel.Connection = Login.Connection
                Dim reader_panel As MySqlDataReader
                reader_panel = cmd_panel.ExecuteReader

                If reader_panel.HasRows Then
                    While reader_panel.Read
                        rio_grid2.Rows.Add(New String() {reader_panel(0)})
                    End While
                End If

                reader_panel.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try


            For i = 0 To rio_grid2.Rows.Count - 1
                Call Rio_needed(i)
            Next

            cell_ch = True

        End If

    End Sub

    Sub Rio_needed(index As Integer)
        'calculate the needed and available RIO

        Dim n_inputs As Double : n_inputs = 0
        Dim n_outputs As Double : n_outputs = 0
        Dim n_motion As Double : n_motion = 0


        Try
            Dim cmd_panel As New MySqlCommand
            cmd_panel.Parameters.AddWithValue("@sol", sol_box.Text)
            cmd_panel.Parameters.AddWithValue("@feature", rio_grid2.Rows(index).Cells(0).Value)
            cmd_panel.CommandText = "SELECT inputs, output, motion from quote_table.Remote_IO where solutiom = @sol and feature_code = @feature"
            cmd_panel.Connection = Login.Connection
            Dim reader_panel As MySqlDataReader
            reader_panel = cmd_panel.ExecuteReader

            If reader_panel.HasRows Then
                While reader_panel.Read

                    If reader_panel(0) > 0 Then
                        n_inputs = Math.Ceiling(myQuote.inputs_io / reader_panel(0))
                    End If

                    If reader_panel(1) > 0 Then
                        n_outputs = Math.Ceiling(myQuote.outputs_io / reader_panel(1))
                    End If

                    If reader_panel(2) > 0 Then
                        n_motion = Math.Ceiling(myQuote.motion_io / reader_panel(2))
                    End If

                End While
            End If

            rio_grid2.Rows(index).Cells(1).Value = Max_number(n_inputs, n_outputs, n_motion)
            rio_grid2.Rows(index).Cells(2).Value = Max_number(n_inputs, n_outputs, n_motion)

            reader_panel.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


    End Sub

    Private Sub rio_grid2_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles rio_grid2.CellValueChanged

        Call eval_ava()

    End Sub

    Sub eval_ava()

        If trigger = True And cell_ch = True Then

            Dim av_inputs As Double = av_inputs = 0
            Dim av_outputs As Double = av_outputs = 0
            Dim av_motion As Double = av_motion = 0


            For i = 0 To rio_grid2.Rows.Count - 1

                If IsNumeric(rio_grid2.Rows(i).Cells(2).Value) = True Then

                    If rio_grid2.Rows(i).Cells(2).Value > 0 Then

                        Try
                            Dim cmd_panel As New MySqlCommand
                            cmd_panel.Parameters.AddWithValue("@sol", sol_box.Text)
                            cmd_panel.Parameters.AddWithValue("@feature", rio_grid2.Rows(i).Cells(0).Value)
                            cmd_panel.CommandText = "SELECT inputs, output, motion from quote_table.Remote_IO where solutiom = @sol and feature_code = @feature"
                            cmd_panel.Connection = Login.Connection
                            Dim reader_panel As MySqlDataReader
                            reader_panel = cmd_panel.ExecuteReader

                            If reader_panel.HasRows Then
                                While reader_panel.Read
                                    av_inputs = av_inputs + CType(reader_panel(0), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)
                                    av_outputs = av_outputs + CType(reader_panel(1), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)
                                    av_motion = av_motion + CType(reader_panel(2), Double) * CType(rio_grid2.Rows(i).Cells(2).Value, Double)

                                End While
                            End If

                            reader_panel.Close()

                        Catch ex As Exception
                            MessageBox.Show(ex.ToString)
                        End Try

                    End If
                End If

            Next

            RIO_grid.Rows(1).Cells(1).Value = Math.Ceiling(av_inputs + 1)
            RIO_grid.Rows(1).Cells(2).Value = Math.Ceiling(av_outputs + 1)
            RIO_grid.Rows(1).Cells(3).Value = Math.Ceiling(av_motion + 1)


            If RIO_grid.Rows(1).Cells(1).Value < RIO_grid.Rows(0).Cells(1).Value Then
                RIO_grid.Rows(1).Cells(1).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(1).Style.BackColor = Color.LightGray
            End If

            If RIO_grid.Rows(1).Cells(2).Value < RIO_grid.Rows(0).Cells(2).Value Then
                RIO_grid.Rows(1).Cells(2).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(2).Style.BackColor = Color.LightGray
            End If

            If RIO_grid.Rows(1).Cells(3).Value < RIO_grid.Rows(0).Cells(3).Value Then
                RIO_grid.Rows(1).Cells(3).Style.BackColor = Color.Brown
            Else
                RIO_grid.Rows(1).Cells(3).Style.BackColor = Color.LightGray
            End If

            '---- warning label (you have not compensate -----
            If RIO_grid.Rows(1).Cells(1).Value < RIO_grid.Rows(0).Cells(1).Value Or RIO_grid.Rows(1).Cells(2).Value < RIO_grid.Rows(0).Cells(2).Value Or RIO_grid.Rows(1).Cells(3).Value < RIO_grid.Rows(0).Cells(3).Value Then
                warning.Text = "You have not compensate all the IOs required"
            ElseIf RIO_grid.Rows(1).Cells(1).Value >= RIO_grid.Rows(0).Cells(1).Value Or RIO_grid.Rows(1).Cells(2).Value >= RIO_grid.Rows(0).Cells(2).Value Or RIO_grid.Rows(1).Cells(3).Value >= RIO_grid.Rows(0).Cells(3).Value Then
                warning.Text = ""
            End If


            If rio_grid2.Rows.Count > 1 Then
                mycost.Text = "Cost $: " & 363.14 * If(IsNumeric(rio_grid2.Rows(0).Cells(2).Value), rio_grid2.Rows(0).Cells(2).Value, 0) + 258.14 * If(IsNumeric(rio_grid2.Rows(1).Cells(2).Value), rio_grid2.Rows(1).Cells(2).Value, 0)
            End If

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myQuote.IO_custom = True
        myQuote.RIO_press = True
        Call myQuote.General_calculation()
        Me.Close()
    End Sub
End Class