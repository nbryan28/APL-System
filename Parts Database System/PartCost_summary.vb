Public Class PartCost_summary
    Private Sub type_box_SelectedValueChanged(sender As Object, e As EventArgs) Handles type_box_c.SelectedValueChanged
        'Update part cost parts qty grid when type in the dropbox is changed

        Dim type As String : type = ""

        If Not type_box_c.SelectedItem Is Nothing Then
            type = type_box_c.SelectedItem.ToString
        End If

        If String.Equals(type, "") = False Then

            Call ADA_Setup.Cal_PartCost(type)

        End If
    End Sub

    Private Sub PartCost_summary_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = Keys.F AndAlso e.Modifiers = Keys.Control) Then
            mydatagrid = 3
            Search_grid.ShowDialog()
        End If
    End Sub

    Private Sub PartCost_summary_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'load values
        Dim labor_tt As Double : labor_tt = 0
        Dim mat_tt As Double : mat_tt = 0
        Dim total_tt As Double = total_tt = 0

        If IsNumeric(ADA_Setup.totals_grid.Rows(0).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(0).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(2).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(2).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(4).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(4).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(7).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(7).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(9).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(9).Cells(1).Value()
        End If

        '---------------------------------------
        If IsNumeric(ADA_Setup.totals_grid.Rows(1).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(1).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(3).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(3).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(5).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(5).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(8).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(8).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(10).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(10).Cells(1).Value()
        End If

        big_materials.Text = "Total Materials Cost: $" & mat_tt
        big_labor.Text = "Total Labor Cost: $" & labor_tt
        big_cost.Text = "Total Cost: $" & mat_tt + labor_tt

        type_box_c.Text = "Starter Panel"

    End Sub

    Private Sub Parts_cost_t_grid_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Parts_cost_t_grid.CellValueChanged
        'load values
        Dim labor_tt As Double : labor_tt = 0
        Dim mat_tt As Double : mat_tt = 0
        Dim total_tt As Double = total_tt = 0

        If ADA_Setup.Visible = True Then

            If IsNumeric(ADA_Setup.totals_grid.Rows(0).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(0).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(2).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(2).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(4).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(4).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(7).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(7).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(9).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(9).Cells(1).Value()
            End If

            '---------------------------------------
            If IsNumeric(ADA_Setup.totals_grid.Rows(1).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(1).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(3).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(3).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(5).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(5).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(8).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(8).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(10).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(10).Cells(1).Value()
            End If

            big_materials.Text = "Total Materials Cost: $" & mat_tt
            big_labor.Text = "Total Labor Cost: $" & labor_tt
            big_cost.Text = "Total Cost: $" & mat_tt + labor_tt

        End If

    End Sub

    Private Sub Parts_cost_t_grid_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles Parts_cost_t_grid.RowsAdded
        Dim labor_tt As Double : labor_tt = 0
        Dim mat_tt As Double : mat_tt = 0
        Dim total_tt As Double = total_tt = 0

        If ADA_Setup.Visible = True Then

            If IsNumeric(ADA_Setup.totals_grid.Rows(0).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(0).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(2).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(2).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(4).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(4).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(7).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(7).Cells(1).Value()
            End If
            If IsNumeric(ADA_Setup.totals_grid.Rows(9).Cells(1).Value()) = True Then
                mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(9).Cells(1).Value()
            End If

            '---------------------------------------
            If IsNumeric(ADA_Setup.totals_grid.Rows(1).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(1).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(3).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(3).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(5).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(5).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(8).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(8).Cells(1).Value()
            End If

            If IsNumeric(ADA_Setup.totals_grid.Rows(10).Cells(1).Value()) = True Then
                labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(10).Cells(1).Value()
            End If

            big_materials.Text = "Total Materials Cost: $" & mat_tt
            big_labor.Text = "Total Labor Cost: $" & labor_tt
            big_cost.Text = "Total Cost: $" & mat_tt + labor_tt

        End If
    End Sub

    Private Sub Parts_cost_t_grid_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles Parts_cost_t_grid.RowsRemoved
        Dim labor_tt As Double : labor_tt = 0
        Dim mat_tt As Double : mat_tt = 0
        Dim total_tt As Double = total_tt = 0

        If IsNumeric(ADA_Setup.totals_grid.Rows(0).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(0).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(2).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(2).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(4).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(4).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(7).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(7).Cells(1).Value()
        End If
        If IsNumeric(ADA_Setup.totals_grid.Rows(9).Cells(1).Value()) = True Then
            mat_tt = mat_tt + ADA_Setup.totals_grid.Rows(9).Cells(1).Value()
        End If

        '---------------------------------------
        If IsNumeric(ADA_Setup.totals_grid.Rows(1).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(1).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(3).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(3).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(5).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(5).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(8).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(8).Cells(1).Value()
        End If

        If IsNumeric(ADA_Setup.totals_grid.Rows(10).Cells(1).Value()) = True Then
            labor_tt = labor_tt + ADA_Setup.totals_grid.Rows(10).Cells(1).Value()
        End If

        big_materials.Text = "Total Materials Cost: $" & mat_tt
        big_labor.Text = "Total Labor Cost: $" & labor_tt
        big_cost.Text = "Total Cost: $" & mat_tt + labor_tt
    End Sub
End Class