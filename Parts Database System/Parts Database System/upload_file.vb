Imports System.Reflection
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Net.Mail
Imports System.IO
Imports MySql.Data.MySqlClient


Public Class upload_file
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        OpenFileDialog1.Filter = "Pdf Files|*.pdf"

        If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Dim sourcepath As String = TextBox1.Text
            Dim DestPath As String = "O:\atlanta\Users\Bryan Dahlqvist\Parts Database System\Parts_Specs\" & OpenFileDialog1.SafeFileName
            File.Copy(sourcepath, DestPath, True)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try


        MessageBox.Show("File upload successfully!")
        Me.Visible = False
        MyDash.Visible = True
    End Sub

    Private Sub upload_file_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
    End Sub

    Private Sub upload_file_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible = True Then
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub
End Class