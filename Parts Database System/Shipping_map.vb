Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Net.Mail

Public Class Shipping_map


    Private Sub Shipping_map_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Dim cmd51 As New MySqlCommand
            cmd51.Parameters.Clear()
            cmd51.Parameters.AddWithValue("@mr_name", Procurement_Overview.Text)
            cmd51.CommandText = "SELECT shipping_ad, ship_notes from Material_Request.mr where mr_name = @mr_name"
            cmd51.Connection = Login.Connection
            Dim reader51 As MySqlDataReader
            reader51 = cmd51.ExecuteReader

            If reader51.HasRows Then
                While reader51.Read
                    If reader51.IsDBNull(0) = False Then
                        TextBox1.Text = reader51(0).ToString
                    End If

                    If reader51.IsDBNull(1) = False Then
                        TextBox2.Text = reader51(1).ToString
                    End If
                End While
            End If

            reader51.Close()

            Dim map_q() As String : map_q = TextBox1.Text.ToString.Split(" "c)
            Dim new_q As String = ""

            For i = 0 To map_q.Length - 1
                new_q = new_q & map_q(i) & "+"
            Next


            ' WebBrowser1.Navigate("https://www.google.com/maps/place/910+Belcourt+Pkwy,+Roswell,+GA+30076")
            WebBrowser1.Navigate("https://www.google.com/maps/place/" & new_q)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class