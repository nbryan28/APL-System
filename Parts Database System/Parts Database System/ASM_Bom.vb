Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ASM_Bom


    Private Sub ASM_Bom_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            '    '--------------  add to device
            Dim cmd2 As New MySqlCommand
            cmd2.CommandText = "SELECT distinct Legacy_ADA_Number from devices"
            cmd2.Connection = Login.Connection
            Dim reader2 As MySqlDataReader
            reader2 = cmd2.ExecuteReader

            If reader2.HasRows Then
                While reader2.Read
                    ComboBox1.Items.Add(reader2(0).ToString)
                End While
            End If

            reader2.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Try
            part_assembly.Rows.Clear()
            Dim adv_n As String : adv_n = ""

            Dim cmd_a As New MySqlCommand
            cmd_a.Parameters.AddWithValue("@assem", ComboBox1.SelectedItem.ToString)
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
            cmd_pd.CommandText = "SELECT adv.Qty, p1.Manufacturer, p1.Part_Name, p1.Part_Description from parts_table as p1 INNER JOIN adv ON p1.Part_Name = adv.Part_Name where adv.ADV_Number = @device"
            cmd_pd.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_pd.ExecuteReader

            If readerv.HasRows Then
                Dim i As Integer : i = 0
                While readerv.Read

                    part_assembly.Rows.Add()  'add a new row
                    part_assembly.Rows(i).Cells(0).Value = readerv(0).ToString
                    part_assembly.Rows(i).Cells(1).Value = readerv(1).ToString
                    part_assembly.Rows(i).Cells(2).Value = readerv(2).ToString
                    part_assembly.Rows(i).Cells(3).Value = readerv(3).ToString


                    i = i + 1
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If IsNumeric(TextBox1.Text) = True Then

            For i = 0 To part_assembly.Rows.Count - 1
                part_assembly.Rows(i).Cells(4).Value = part_assembly.Rows(i).Cells(0).Value * TextBox1.Text
            Next

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim appPath As String = Application.StartupPath()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim xlWorkSheet As Excel.Worksheet
        xlApp.DisplayAlerts = False

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
        Else

            Try
                Dim wb As Excel.Workbook = xlApp.Workbooks.Open(appPath & "\ASSEM.xlsx")
                xlWorkSheet = wb.Sheets("ASSEMBLY")

                'Start filling the values

                xlWorkSheet.Cells(1, 1) = ComboBox1.SelectedItem.ToString
                xlWorkSheet.Cells(2, 10) = TextBox1.Text

                For i = 0 To part_assembly.Rows.Count - 1
                    xlWorkSheet.Cells(i + 4, 1) = part_assembly.Rows(i).Cells(0).Value
                    xlWorkSheet.Cells(i + 4, 2) = part_assembly.Rows(i).Cells(1).Value
                    xlWorkSheet.Cells(i + 4, 3) = part_assembly.Rows(i).Cells(2).Value
                    xlWorkSheet.Cells(i + 4, 4) = part_assembly.Rows(i).Cells(3).Value
                Next


                SaveFileDialog1.Filter = "Excel Files|*.xlsx;*.xlsm"

                If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                    wb.SaveCopyAs(SaveFileDialog1.FileName.ToString)
                End If

                wb.Close(False)


                Marshal.ReleaseComObject(xlApp)
                MessageBox.Show("Assembly BOM generated successfully!")


            Catch ex As Exception
                MessageBox.Show("File not found or corrupted")
            End Try

        End If

    End Sub

    Private Sub AssemblyAndComponentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AssemblyAndComponentsToolStripMenuItem.Click
        Assembly_comp.Visible = True
    End Sub
End Class