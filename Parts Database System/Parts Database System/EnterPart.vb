Imports System.Globalization
Imports MySql.Data.MySqlClient

Public Class EnterPart


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Private Sub EnterPart_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged

        '========= UNELEGANT WAY OF CLEANING CONTROLS ===========================

        Dim ctrl As Control
        For Each ctrl In Me.GroupBox1.Controls
            If (ctrl.GetType() Is GetType(TextBox)) Then
                ctrl.Text = ""
            End If
        Next

        For Each ctrl In Me.GroupBox2.Controls
            If (ctrl.GetType() Is GetType(TextBox)) Then
                ctrl.Text = ""
            End If
        Next

    End Sub



    Private Sub EnterPart_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        '----------- Populate type combobox --------
        Try
            Dim cmd_v As New MySqlCommand
            cmd_v.CommandText = "SELECT Part_Type from parts_type_table"

            cmd_v.Connection = Login.Connection
            Dim readerv As MySqlDataReader
            readerv = cmd_v.ExecuteReader

            If readerv.HasRows Then
                While readerv.Read
                    Type_t.Items.Add(readerv(0))

                End While
            End If

            readerv.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        '----------------Populate Units combobox -----------------------
        Try
            Dim cmd_u As New MySqlCommand
            cmd_u.CommandText = "SELECT distinct Units from parts_table where Units is not null"

            cmd_u.Connection = Login.Connection
            Dim readeru As MySqlDataReader
            readeru = cmd_u.ExecuteReader

            If readeru.HasRows Then
                While readeru.Read
                    unit_t.Items.Add(readeru(0))
                End While
            End If

            readeru.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '----------------------------------- ENTER A NEW PART TO THE PARTS TABLE------------------------------------------

        Dim min_flag As Boolean : min_flag = True


        '---------MAKE SURE TEXTBOXES ARE FILL UP (PART, STATUS AND TYPE) ------------------------
        If String.Equals(part_t.Text, "") = True Or String.Equals(status_t.Text, "") = True Or String.Equals(Type_t.Text, "") = True Then
            MessageBox.Show("Please Fill Part Name, Part Status and Part Type fields")
        Else

            '----------- MAKE SURE THE PART DOES NOT EXIST ------------
            Try
                Dim cmd As New MySqlCommand
                Dim Create_cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                cmd.CommandText = "SELECT Part_Name from parts_table where Part_Name = @Part_Name"
                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader
                reader = cmd.ExecuteReader

                If String.Equals(min_t.Text, "") = False And IsNumeric(min_t.Text) = False Then
                    min_flag = False
                    MessageBox.Show("Invalid Minimum Order Quantity")
                    reader.Close()
                End If

                If min_flag = True Then

                    If Not reader.HasRows Then

                        'IF IT DOES NOT EXIST THEN CREATE IT
                        reader.Close()

                        'check for empty and invalid min qty
                        If String.Equals(min_t.Text, "") Then
                            min_t.Text = 1
                        ElseIf CType(min_t.Text, Integer) <= 0 Then
                            min_t.Text = 1
                        End If

                        Create_cmd.Parameters.AddWithValue("@Part_Name", part_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Manufacturer", manu_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Description", desc_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Notes", notes_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Status", status_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Part_Type", Type_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Units", unit_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Min_Order_Qty", min_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Legacy_ADA_Number", legacy_t.Text)
                        Create_cmd.Parameters.AddWithValue("@Primary_Vendor", p_vendor_t.Text)
                        Create_cmd.Parameters.AddWithValue("@MFG_type", ComboBox1.Text)


                        Create_cmd.CommandText = "INSERT INTO parts_table(Part_Name, Manufacturer, Part_Description, Notes, Part_Status, Part_Type, Units, Min_Order_Qty, Legacy_ADA_Number, Primary_Vendor, MFG_type) VALUES (@Part_Name, @Manufacturer, @Part_Description, @Notes, @Part_Status, @Part_Type, @Units, @Min_Order_Qty, @Legacy_ADA_Number, @Primary_Vendor, @MFG_type)"
                        Create_cmd.Connection = Login.Connection
                        Create_cmd.ExecuteNonQuery()

                        '--- insert into inventory
                        Dim Create_cmdi As New MySqlCommand
                        Create_cmdi.Parameters.AddWithValue("@part_name", part_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@description", desc_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@manufacturer", manu_t.Text)
                        Create_cmdi.Parameters.AddWithValue("@MFG_type", ComboBox1.Text)
                        Create_cmdi.Parameters.AddWithValue("@units", unit_t.Text)

                        Create_cmdi.CommandText = "INSERT INTO inventory.inventory_qty(part_name, description, manufacturer, MFG_type, units, min_qty, max_qty, current_qty, Qty_in_order) VALUES (@part_name, @description, @manufacturer, @MFG_type, @units,0,0,0,0)"
                        Create_cmdi.Connection = Login.Connection
                        Create_cmdi.ExecuteNonQuery()

                        '---------------------------

                        MessageBox.Show("Part " & part_t.Text & " created succesfully")


                    Else
                        MessageBox.Show("Part " & part_t.Text & " already exist!")
                        'DO NOT FORGET TO CLOSE THE READER
                        reader.Close()
                    End If

                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)

            End Try
        End If


    End Sub

End Class