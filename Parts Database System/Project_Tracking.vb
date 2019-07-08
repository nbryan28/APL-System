Public Class Project_Tracking
    Private Sub Project_Tracking_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '----remove tabs ------
        TabControl1.TabPages.Remove(TabPage2)
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        TabControl1.TabPages.Remove(TabPage5)
        TabControl1.TabPages.Remove(TabPage6)
        TabControl1.TabPages.Remove(TabPage7)
        TabControl1.TabPages.Remove(TabPage8)
        TabControl1.TabPages.Remove(TabPage9)
        TabControl1.TabPages.Remove(TabPage10)
        TabControl1.TabPages.Remove(TabPage11)
        TabControl1.TabPages.Remove(TabPage12)
        TabControl1.TabPages.Remove(TabPage13)
        TabControl1.TabPages.Remove(TabPage14)
        TabControl1.TabPages.Remove(TabPage15)

        Dim tronco(8) As Project_status
        Dim y As Integer : y = 19

        For i = 0 To 7

            'branches
            tronco(i) = New Project_status
            tronco(i).Location = New Point(14, y)

            Panel3.Controls.Add(tronco(i))
            y = y + 279

        Next


    End Sub

    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter
        Label3.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave
        Label3.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label6_MouseEnter(sender As Object, e As EventArgs) Handles Label6.MouseEnter
        Label6.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label6_MouseLeave(sender As Object, e As EventArgs) Handles Label6.MouseLeave
        Label6.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label8_MouseEnter(sender As Object, e As EventArgs) Handles Label8.MouseEnter
        Label8.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label8_MouseLeave(sender As Object, e As EventArgs) Handles Label8.MouseLeave
        Label8.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles Label5.MouseEnter
        Label5.ForeColor = Color.CadetBlue
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles Label5.MouseLeave
        Label5.ForeColor = Color.WhiteSmoke
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        If Log_out = True Then
            Me.Visible = False
            MyDash.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub
End Class