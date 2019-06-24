Public Class Menu_db


    Private Sub Label8_MouseEnter(sender As Object, e As EventArgs) Handles Label8.MouseEnter

        Label8.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label8_MouseLeave(sender As Object, e As EventArgs) Handles Label8.MouseLeave

        Label8.ForeColor = Color.White
    End Sub

    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter

        Label3.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave

        Label3.ForeColor = Color.White
    End Sub


    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter

        Label2.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave

        Label2.ForeColor = Color.White
    End Sub

    Private Sub Label11_MouseEnter(sender As Object, e As EventArgs) Handles Label11.MouseEnter

        Label11.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label11_MouseLeave(sender As Object, e As EventArgs) Handles Label11.MouseLeave

        Label11.ForeColor = Color.White
    End Sub

    Private Sub Label9_MouseEnter(sender As Object, e As EventArgs) Handles Label9.MouseEnter

        Label9.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label9_MouseLeave(sender As Object, e As EventArgs) Handles Label9.MouseLeave

        Label9.ForeColor = Color.White
    End Sub

    Private Sub Label10_MouseEnter(sender As Object, e As EventArgs) Handles Label10.MouseEnter

        Label10.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label10_MouseLeave(sender As Object, e As EventArgs) Handles Label10.MouseLeave

        Label10.ForeColor = Color.White
    End Sub




    Private Sub Label14_MouseEnter(sender As Object, e As EventArgs) Handles Label14.MouseEnter
        Label14.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label14_MouseLeave(sender As Object, e As EventArgs) Handles Label14.MouseLeave
        Label14.ForeColor = Color.White
    End Sub


    Private Sub Label6_MouseEnter(sender As Object, e As EventArgs) Handles Label6.MouseEnter
        Label6.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label6_MouseLeave(sender As Object, e As EventArgs) Handles Label6.MouseLeave
        Label6.ForeColor = Color.White
    End Sub

    Private Sub Label15_MouseEnter(sender As Object, e As EventArgs) Handles Label15.MouseEnter
        Label15.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label15_MouseLeave(sender As Object, e As EventArgs) Handles Label15.MouseLeave
        Label15.ForeColor = Color.White
    End Sub




    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.ForeColor = Color.White
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles Label5.MouseEnter
        Label5.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles Label5.MouseLeave
        Label5.ForeColor = Color.White
    End Sub

    Private Sub Label7_MouseEnter(sender As Object, e As EventArgs) Handles Label7.MouseEnter
        Label7.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label7_MouseLeave(sender As Object, e As EventArgs) Handles Label7.MouseLeave
        Label7.ForeColor = Color.White
    End Sub

    Private Sub Label12_MouseEnter(sender As Object, e As EventArgs) Handles Label12.MouseEnter
        Label12.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label12_MouseLeave(sender As Object, e As EventArgs) Handles Label12.MouseLeave
        Label12.ForeColor = Color.White
    End Sub


    Private Sub Label18_MouseEnter(sender As Object, e As EventArgs) Handles Label18.MouseEnter
        Label18.ForeColor = Color.DarkCyan
    End Sub

    Private Sub Label18_MouseLeave(sender As Object, e As EventArgs) Handles Label18.MouseLeave
        Label18.ForeColor = Color.White
    End Sub





    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

        '======== USER ACCOUNTS FORM ============

        If Procurement_management = True Then

            User_form.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

        '============= CONNECTIONS ========================

        If Procurement_management = True Then
            Connections_win.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

        '============= IMPORT CSV ========================
        ImportCSV.Visible = True
        Me.Visible = False

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

        If Procurement_management = True Or Engineer_management = True Then

            '================ ENTER NEW PART ===================
            EnterPart.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")

        End If

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click

        '=============== ADD KITS =======================

        If Procurement_management = True Then
            add_KIT.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

        '=============== UPDATE/REMOVE PARTS =======================

        If Procurement_management = True Then
            UR_part.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub





    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        ' ================= ADD DEVICES ==============================

        If Procurement_management = True Then
            add_DEVICE.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub


    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        '===============UPDATE KIT===========
        If Procurement_management = True Then
            edit_KIT.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        '===============UPDATE DEVICE===========
        If Procurement_management = True Then
            edit_DEVICE.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        '===============PROCESS PO===========
        If Procurement_management = True Then
            EnterPO.Visible = True
            Me.Visible = False
        Else
            MessageBox.Show("You do not have permission to access this feauture")
        End If
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

        '================ REPORTS ===================

        If Procurement_management = True Or Procurement = True Then

            Reports.Visible = True
            Me.Visible = False

        End If

    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click

        '================ PROJECT MANAGEMENT ===================
        Projects_m.Visible = True
        Me.Visible = False

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        '================ PROCUREMENT ===================
        ADA_Setup.Visible = True
        Me.Visible = False

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        '=========== VENDORS AND PRICES ENTER AND UPDATE
        If Procurement_management = True Or Procurement = True Then
            Update_vendors.Visible = True
            Me.Visible = False
        End If
    End Sub



End Class