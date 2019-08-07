Public Class Parts_needed


    Private Sub Parts_needed_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'call searcher
        If (e.KeyCode = Keys.F AndAlso e.Modifiers = Keys.Control) Then
            mydatagrid = 2
            Search_grid.ShowDialog()
        End If
    End Sub

    Private Sub type_box_SelectedValueChanged_1(sender As Object, e As EventArgs) Handles type_box.SelectedValueChanged
        'Update ada set parts qty grid when type in the dropbox is changed

        Dim type As String : type = ""
        Dim myindex As Integer : myindex = ADA_Setup.Get_set_number(ADA_Setup.ADA_set_list, Me.Text)

        If Not type_box.SelectedItem Is Nothing Then
            type = type_box.SelectedItem.ToString
        End If

        If String.Equals(type, "") = False And myindex > -1 Then

            Call ADA_Setup.my_ADA_parts(type, myindex)

        End If
    End Sub


End Class