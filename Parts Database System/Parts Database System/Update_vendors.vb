Imports MySql.Data.MySqlClient

Public Class Update_vendors


    Public index_s As Integer



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '----------------------------------- LOAD VENDORS ----------------------------------------
        If String.IsNullOrEmpty(Part_textbox.Text) = False Then
            Try

                Dim exist_p As Boolean : exist_p = False
                Dim p_vendor As String : p_vendor = "00000"
                vendor_grid.Rows.Clear()

                '------ get vendor if parts exist ----
                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_Name", Part_textbox.Text)
                cmd.CommandText = "SELECT Primary_Vendor from parts_table where Part_Name = @Part_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    While reader.Read
                        exist_p = True
                        p_vendor = reader(0).ToString
                    End While
                End If

                reader.Close()


                '-----------------------------------

                Dim cmd4 As New MySqlCommand
                cmd4.Parameters.AddWithValue("@part_name", Part_textbox.Text)
                cmd4.CommandText = "SELECT Vendor_Name, Cost, Purchase_Date from vendors_table where Part_name = @part_name"
                cmd4.Connection = Login.Connection
                Dim reader4 As MySqlDataReader
                reader4 = cmd4.ExecuteReader

                If reader4.HasRows Then

                    While reader4.Read
                        vendor_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString})
                    End While

                    part_n.Text = Part_textbox.Text

                    reader4.Close()

                Else

                    reader4.Close()
                    vendor_grid.Rows.Clear()
                    date_box.Text = ""
                    vendor_box.Text = ""
                    Price_box.Text = ""

                    If exist_p = False Then
                        MessageBox.Show("Please, add " & Part_textbox.Text & " to the parts table first!")
                        part_n.Text = "Part selected:"
                    Else

                        Dim result As Integer = MessageBox.Show("No Pricing record was found for " & Part_textbox.Text & ", Would you like to create a new entry?", "Not Record Found", MessageBoxButtons.YesNoCancel)

                        If result = DialogResult.Yes Then

                            Dim qty_change As String : qty_change = 0
                            qty_change = InputBox("Please enter the Cost of this part")

                            If IsNumeric(qty_change) = True Then

                                If qty_change >= 0 Then
                                    Dim Create_cmd6 As New MySqlCommand
                                    Create_cmd6.Parameters.Clear()
                                    Create_cmd6.Parameters.AddWithValue("@Part_name", Part_textbox.Text)
                                    Create_cmd6.Parameters.AddWithValue("@Vendor_Name", p_vendor)
                                    Create_cmd6.Parameters.AddWithValue("@Cost", Decimal.Parse(qty_change))
                                    Create_cmd6.Parameters.AddWithValue("@Purchase_Date", Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)

                                    Create_cmd6.CommandText = "INSERT INTO vendors_table(Part_Name, Vendor_Name, Cost, Purchase_Date) VALUES (@Part_name, @Vendor_Name, @Cost, @Purchase_Date)"
                                    Create_cmd6.Connection = Login.Connection
                                    Create_cmd6.ExecuteNonQuery()
                                End If

                                part_n.Text = Part_textbox.Text
                            Else
                                MsgBox("Please enter a number.")
                                part_n.Text = "Part selected:"
                            End If

                            Call Refresh_tables()

                        Else
                            part_n.Text = "Part selected:"
                        End If

                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub vendor_grid_DoubleClick(sender As Object, e As EventArgs) Handles vendor_grid.DoubleClick
        'Show vendor data in its corresponding textboxes while double clicking

        If vendor_grid.Rows.Count > 0 Then

            For j = 0 To vendor_grid.Rows.Count - 1
                vendor_grid.Rows(j).DefaultCellStyle.BackColor = Color.WhiteSmoke
            Next

            Dim index_k = vendor_grid.CurrentCell.RowIndex

            vendor_box.Text = If(String.IsNullOrEmpty(vendor_grid.Rows(index_k).Cells(0).Value) = True, "", vendor_grid.Rows(index_k).Cells(0).Value)
            Price_box.Text = If(String.IsNullOrEmpty(vendor_grid.Rows(index_k).Cells(1).Value) = True, "", vendor_grid.Rows(index_k).Cells(1).Value)

            Dim new_date() As String
            Dim left_o() As String
            Dim better_date As String : better_date = ""

            If String.IsNullOrEmpty(vendor_grid.Rows(index_k).Cells(2).Value()) = False Then
                new_date = vendor_grid.Rows(index_k).Cells(2).Value().ToString.Split("/")
                left_o = new_date(2).Split(" ")
                better_date = left_o(0) & "-" & new_date(0) & "-" & new_date(1)
            End If


            date_box.Text = better_date

            vendor_grid.Rows(index_k).DefaultCellStyle.BackColor = Color.Tan
            index_s = index_k

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If String.Equals(part_n.Text, "Part selected:") = False Then
            'Enter new vendor entry

            Try
                Dim Create_cmd2 As New MySqlCommand
                Dim date_t As String : date_t = Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day

                Create_cmd2.Parameters.AddWithValue("@Part_Name", part_n.Text)
                Create_cmd2.Parameters.AddWithValue("@Primary_Vendor", If(String.IsNullOrEmpty(vendor_box.Text) = True, "Unknown", vendor_box.Text))

                'Check if price is valid
                If String.Equals(Price_box.Text, "") = False And IsNumeric(Price_box.Text) Then
                    Create_cmd2.Parameters.AddWithValue("@Price", Price_box.Text)
                Else
                    Create_cmd2.Parameters.AddWithValue("@Price", "0")
                End If

                If date_box.Text Like "####-##-##" Or date_box.Text Like "####-#-##" Or date_box.Text Like "####-#-#" Or date_box.Text Like "####-##-#" Then
                    date_t = date_box.Text
                Else
                    date_t = Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day
                End If

                Create_cmd2.Parameters.AddWithValue("@Date", date_t)
                Create_cmd2.CommandText = "INSERT INTO vendors_table (Part_Name, Vendor_Name, Cost, Purchase_Date) VALUES (@Part_Name, @Primary_Vendor,  @Price, @Date);"
                Create_cmd2.Connection = Login.Connection
                Create_cmd2.ExecuteNonQuery()
                Call Refresh_tables()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub


    Sub Refresh_tables()

        ''---------- REFRESH DATA IN GRID ----------------------
        vendor_grid.Rows.Clear()

        Dim cmd4 As New MySqlCommand
        cmd4.Parameters.AddWithValue("@part_name", part_n.Text)
        cmd4.CommandText = "SELECT Vendor_Name, Cost, Purchase_Date from vendors_table where Part_name = @part_name"
        cmd4.Connection = Login.Connection
        Dim reader4 As MySqlDataReader
        reader4 = cmd4.ExecuteReader

        If reader4.HasRows Then

            While reader4.Read
                vendor_grid.Rows.Add(New String() {reader4(0).ToString, reader4(1).ToString, reader4(2).ToString})
            End While

            ' part_n.Text = Part_textbox.Text

        End If

        reader4.Close()
        index_s = -1
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        '-------- Delete entry --------

        If index_s > -1 Then
            Try

                Dim Create_vd As New MySqlCommand
                Dim dlgR As DialogResult
                dlgR = MessageBox.Show("Are you sure you want to delete this Vendor entry?", "Attention!", MessageBoxButtons.YesNo)

                If dlgR = DialogResult.Yes Then

                    Dim new_date As String : new_date = ""

                    '--current values ---
                    If String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(2).Value) = False Then
                        new_date = better_date(vendor_grid.Rows(index_s).Cells(2).Value)
                    End If

                    Create_vd.Parameters.AddWithValue("@Part_Name", part_n.Text)
                    Create_vd.Parameters.AddWithValue("@Primary_Vendor", If(String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(0).Value) = True, "", vendor_grid.Rows(index_s).Cells(0).Value))
                    Create_vd.Parameters.AddWithValue("@Price", If(String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(1).Value) = True, "", vendor_grid.Rows(index_s).Cells(1).Value))
                    Create_vd.Parameters.AddWithValue("@Date", new_date)
                    Create_vd.CommandText = "DELETE FROM vendors_table where Part_Name = @Part_Name and Vendor_Name = @Primary_Vendor and Cost = @Price and Purchase_date = @Date"
                    Create_vd.Connection = Login.Connection
                    Create_vd.ExecuteNonQuery()

                    Call Refresh_tables()

                End If
            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        Else
            MessageBox.Show("Please double click the entry you want to delete, and click the 'Remove' button")
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        '------- Update entry ---------
        If index_s > -1 Then

            Try
                Dim Create_cmd As New MySqlCommand
                Dim new_date As String : new_date = ""

                '--current values ---
                If String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(2).Value) = False Then
                    new_date = better_date(vendor_grid.Rows(index_s).Cells(2).Value)
                End If

                Create_cmd.Parameters.AddWithValue("@old_vendor", If(String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(0).Value) = True, "", vendor_grid.Rows(index_s).Cells(0).Value))
                Create_cmd.Parameters.AddWithValue("@old_Cost", If(String.IsNullOrEmpty(vendor_grid.Rows(index_s).Cells(1).Value) = True, "", vendor_grid.Rows(index_s).Cells(1).Value))
                Create_cmd.Parameters.AddWithValue("@old_Purchase_date", new_date)

                '--- new values
                Create_cmd.Parameters.AddWithValue("@part_name", part_n.Text)
                Create_cmd.Parameters.AddWithValue("@vendor", vendor_box.Text)
                Create_cmd.Parameters.AddWithValue("@Cost", Price_box.Text)
                Create_cmd.Parameters.AddWithValue("@Purchase_date", date_box.Text)

                Create_cmd.CommandText = "UPDATE vendors_table  SET  Vendor_Name = @vendor, Cost =  @Cost, Purchase_Date = @Purchase_date  where Part_Name = @part_name and Vendor_Name = @old_vendor and Cost = @old_Cost and Purchase_Date = @old_Purchase_date"
                Create_cmd.Connection = Login.Connection
                Create_cmd.ExecuteNonQuery()

                Call Refresh_tables()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try

        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub Update_vendors_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        index_s = -1
    End Sub

    Function better_date(date_t As String)

        Dim new_date() As String
        Dim left_o() As String
        better_date = ""


        new_date = date_t.Split("/")
        left_o = new_date(2).Split(" ")
        better_date = left_o(0) & "-" & new_date(0) & "-" & new_date(1)

    End Function
End Class