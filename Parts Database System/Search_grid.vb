Public Class Search_grid



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        mysearch(mydatagrid)
    End Sub

    Sub mysearch(option_t As Integer)

        Dim found_po As Boolean : found_po = False
        Dim rowindex As Integer : rowindex = 0


        If String.IsNullOrEmpty(my_search.Text) = False Then

            If option_t = 1 Then

                For Each row As DataGridViewRow In ADA_Setup.PR_grid.Rows
                    If row.IsNewRow Then Continue For
                    If String.Compare(row.Cells.Item("ADA_number").Value.ToString, my_search.Text) = 0 Then
                        rowindex = row.Index
                        ADA_Setup.PR_grid.CurrentCell = ADA_Setup.PR_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If
                Next
                If found_po = False Then
                    MsgBox("ADA Part not found!")
                End If

            ElseIf option_t = 2 Then

                For Each row As DataGridViewRow In Parts_needed.Parts_need_grid.Rows
                    If row.IsNewRow Then Continue For
                    If String.Compare(row.Cells.Item("ADA_number").Value.ToString, my_search.Text) = 0 Then
                        rowindex = row.Index
                        Parts_needed.Parts_need_grid.CurrentCell = Parts_needed.Parts_need_grid.Rows(rowindex).Cells(0)
                        found_po = True
                        Exit For
                    End If
                Next
                If found_po = False Then
                        MsgBox("ADA Part not found!")
                    End If

                    ElseIf option_t = 3 Then

                        For Each row As DataGridViewRow In PartCost_summary.Parts_cost_t_grid.Rows
                            If row.IsNewRow Then Continue For
                            If String.Compare(row.Cells.Item("ADA_number").Value.ToString, my_search.Text) = 0 Then
                                rowindex = row.Index
                                PartCost_summary.Parts_cost_t_grid.CurrentCell = PartCost_summary.Parts_cost_t_grid.Rows(rowindex).Cells(0)
                                found_po = True
                                Exit For
                            End If
                        Next
                        If found_po = False Then
                            MsgBox("ADA Part not found!")
                        End If
                    Else

                    End If

        End If
    End Sub


End Class