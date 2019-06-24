Imports MySql.Data.MySqlClient
Imports System.Text.RegularExpressions

Public Class Search_order

    Dim jobs_list As New List(Of String)()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ListBox1.Items.Clear()


        '------- populate a list with all the jobs

        Dim parts_f_list As New List(Of String)()

        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SHOW TABLES in jobs"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    jobs_list.Add(readerv(0).ToString)
                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        '============== Find Partial matches ============================

        For k = 0 To jobs_list.Count - 1

            Try
                Dim cmdstr As String : cmdstr = "SELECT distinct Part_Description from  jobs." & jobs_list(k) & " where Part_Description LIKE  @search"
                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                cmd.Parameters.AddWithValue("@search", "%" & TextBox1.Text & "%")

                Dim readerf As MySqlDataReader
                readerf = cmd.ExecuteReader

                If readerf.HasRows Then
                    While readerf.Read
                        parts_f_list.Add(readerf(0).ToString)
                    End While
                End If

                readerf.Close()


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next


        Dim final_list As List(Of String) = parts_f_list.Distinct().ToList

        For Each c In final_list
            ListBox1.Items.Add(c)
        Next

        Label1.Text = "Select one of the " & final_list.Count & " parts found"

    End Sub



    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick

        'Populate listbox2 with all the POs related to the part selected

        'ListBox2.Items.Add(ListBox1.SelectedItem)
        ListBox2.Items.Clear()
        Dim po_f_list As New List(Of String)()

        For k = 0 To jobs_list.Count - 1

            Try
                Dim cmdstr As String : cmdstr = "SELECT distinct PO from  jobs." & jobs_list(k) & " where Part_Description =  @sel_p ORDER BY PO desc"
                Dim cmd As New MySqlCommand(cmdstr, Login.Connection)
                cmd.Parameters.AddWithValue("@sel_p", ListBox1.SelectedItem)

                Dim readerp As MySqlDataReader
                readerp = cmd.ExecuteReader

                If readerp.HasRows Then
                    While readerp.Read
                        po_f_list.Add(readerp(0).ToString)
                    End While
                End If

                readerp.Close()


            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Next


        Dim final_po_list As List(Of String) = po_f_list.Distinct().ToList

        For Each c In final_po_list
            ListBox2.Items.Add(c)
        Next


    End Sub


End Class