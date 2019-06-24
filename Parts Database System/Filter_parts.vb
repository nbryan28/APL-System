Public Class Filter_parts

    Public my_index As Integer

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then

            For i = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, True)
            Next

        Else
            For i = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(i, False)
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '-- remove all the rows that are not found in the list
        Dim my_collection = New List(Of String)()

        For Each item In CheckedListBox1.CheckedItems
            my_collection.Add(item.ToString)
        Next


        For j = Form1.DataGridView1.Rows.Count - 1 To 0 Step -1
            If Form1.DataGridView1.Rows(j).Cells(my_index).Value Is DBNull.Value = False Then
                If my_collection.Contains(Form1.DataGridView1.Rows(j).Cells(my_index).Value) = False Then
                    Form1.DataGridView1.Rows.RemoveAt(j)

                End If
            End If
        Next

        Me.Close()
    End Sub
End Class