Public Class Set_wizard
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If String.IsNullOrEmpty(TextBox1.Text) = False Or String.Equals(TextBox1.Text, "") = False Then

            If myQuote.my_Set_info.ContainsKey(TextBox1.ToString) = False Then

                If TextBox1.Text.Contains("ADA") = True Then

                    'TextBox1.Text = "ADA" & myQuote.counter + 1
                    myQuote.counter = myQuote.counter + 1
                End If



                myQuote.Panel_grid.Columns.Add(TextBox1.Text, TextBox1.Text)
                myQuote.PLC_grid.Columns.Add(TextBox1.Text, TextBox1.Text)
                myQuote.Field_grid.Columns.Add(TextBox1.Text, TextBox1.Text)
                myQuote.Control_grid.Columns.Add(TextBox1.Text, TextBox1.Text)


                myQuote.my_Set_info.Add(TextBox1.ToString, TextBox2.ToString)



                For i = 0 To myQuote.Panel_grid.Columns.Count - 1
                    myQuote.Panel_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.PLC_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.Field_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                    myQuote.Control_grid.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                Next

                Me.Close()
            Else
                MessageBox.Show("Set already exist!")
            End If


        Else
                MessageBox.Show("Please enter ADA Set name")
        End If



    End Sub

    Private Sub Set_wizard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = "ADA" & myQuote.counter
    End Sub
End Class