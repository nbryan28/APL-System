Imports MySql.Data.MySqlClient

Public Class Barc_Search


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim part_name_t As String : part_name_t = name_box.Text

        'clear
        part_t.Text = ""
        manu_t.Text = ""
        desc_t.Text = ""
        notes_t.Text = ""
        status_t.Text = ""
        type_t.Text = ""
        units_t.Text = ""
        Min_t.Text = ""
        legacy_t.Text = ""
        mfg_type_t.Text = ""

        Try
            Dim cmd As New MySqlCommand
            cmd.Parameters.AddWithValue("@Part_Name", part_name_t)
            cmd.CommandText = "SELECT * from parts_table where part_name = @Part_Name"
            cmd.Connection = Login.Connection
            Dim reader As MySqlDataReader
            reader = cmd.ExecuteReader

            If reader.HasRows Then

                While reader.Read
                    part_t.Text = reader(0).ToString
                    manu_t.Text = reader(1).ToString
                    desc_t.Text = reader(2).ToString
                    notes_t.Text = reader(3).ToString
                    status_t.Text = reader(4).ToString
                    type_t.Text = reader(5).ToString
                    units_t.Text = reader(6).ToString
                    Min_t.Text = reader(7).ToString
                    legacy_t.Text = reader(8).ToString
                    mfg_type_t.Text = reader(10).ToString

                End While

            End If

            reader.Close()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)

        End Try

        Dim path_im As String = "O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Pictures\"
        Dim component As String : component = part_name_t.ToString.Replace("/", "-")


        Try
            If Not Image.FromFile(path_im & component & ".JPG") Is Nothing Then
                PictureBox1.Image = Image.FromFile(path_im & component & ".JPG")
            Else
                PictureBox1.Image = PictureBox1.InitialImage
            End If
        Catch ex As Exception
            PictureBox1.Image = PictureBox1.InitialImage
        End Try

        Call Show_alternatives(part_t.Text, legacy_t.Text)

    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub



    Sub Show_alternatives(name_part As String, ada_part As String)

        alt_grid.Rows.Clear()
        If String.IsNullOrEmpty(ada_part) = False And String.Equals(ada_part, "") = False And String.IsNullOrEmpty(name_part) = False And String.Equals(name_part, "") = False Then
            Try
                Dim cmd As New MySqlCommand
                cmd.Parameters.AddWithValue("@ADA_name", ada_part)
                cmd.Parameters.AddWithValue("@name_part", name_part)
                cmd.CommandText = "select Part_Name,  Part_Description, Legacy_ADA_Number, Manufacturer from parts_table where LEGACY_ADA_Number = @ADA_name and Part_Name <> @name_part"

                cmd.Connection = Login.Connection
                Dim reader As MySqlDataReader

                reader = cmd.ExecuteReader

                If reader.HasRows Then
                    Dim i As Integer : i = 0
                    While reader.Read

                        alt_grid.Rows.Add()
                        alt_grid.Rows(i).Cells(0).Value = reader(0).ToString  'part name                       
                        alt_grid.Rows(i).Cells(1).Value = reader(1).ToString 'Part description
                        alt_grid.Rows(i).Cells(2).Value = reader(2).ToString  'Ada
                        alt_grid.Rows(i).Cells(4).Value = reader(3).ToString  'mfg

                        i = i + 1
                    End While
                End If

                reader.Close()

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub Barc_Search_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class