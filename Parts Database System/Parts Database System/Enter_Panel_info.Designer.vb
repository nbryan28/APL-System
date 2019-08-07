<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Enter_Panel_info
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Enter_Panel_info))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.q_box = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.desc_box = New System.Windows.Forms.TextBox()
        Me.p_name1 = New System.Windows.Forms.ComboBox()
        Me.p_name2 = New System.Windows.Forms.ComboBox()
        Me.p_name3 = New System.Windows.Forms.ComboBox()
        Me.options_v = New System.Windows.Forms.CheckedListBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(12, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(148, 32)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Panel Name:"
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(121, 187)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 32)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Qty:"
        '
        'q_box
        '
        Me.q_box.BackColor = System.Drawing.Color.WhiteSmoke
        Me.q_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.q_box.Location = New System.Drawing.Point(231, 188)
        Me.q_box.Name = "q_box"
        Me.q_box.Size = New System.Drawing.Size(69, 31)
        Me.q_box.TabIndex = 3
        Me.q_box.Text = "1"
        Me.q_box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(474, 188)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 53)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Done"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(19, 108)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(141, 32)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Description:"
        '
        'desc_box
        '
        Me.desc_box.BackColor = System.Drawing.Color.WhiteSmoke
        Me.desc_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.desc_box.Location = New System.Drawing.Point(231, 110)
        Me.desc_box.Name = "desc_box"
        Me.desc_box.Size = New System.Drawing.Size(448, 31)
        Me.desc_box.TabIndex = 14
        '
        'p_name1
        '
        Me.p_name1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name1.FormattingEnabled = True
        Me.p_name1.Items.AddRange(New Object() {"ADA", "CP", "VFD"})
        Me.p_name1.Location = New System.Drawing.Point(274, 43)
        Me.p_name1.Name = "p_name1"
        Me.p_name1.Size = New System.Drawing.Size(90, 31)
        Me.p_name1.TabIndex = 15
        '
        'p_name2
        '
        Me.p_name2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name2.FormattingEnabled = True
        Me.p_name2.Location = New System.Drawing.Point(434, 43)
        Me.p_name2.Name = "p_name2"
        Me.p_name2.Size = New System.Drawing.Size(90, 31)
        Me.p_name2.TabIndex = 16
        '
        'p_name3
        '
        Me.p_name3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name3.FormattingEnabled = True
        Me.p_name3.Location = New System.Drawing.Point(589, 43)
        Me.p_name3.Name = "p_name3"
        Me.p_name3.Size = New System.Drawing.Size(90, 31)
        Me.p_name3.TabIndex = 17
        '
        'options_v
        '
        Me.options_v.BackColor = System.Drawing.Color.LightGray
        Me.options_v.CheckOnClick = True
        Me.options_v.FormattingEnabled = True
        Me.options_v.Items.AddRange(New Object() {"(0) None", "(1) Brake", "(2) Motor-Fan", "(3) Braking Resistor", "(4) Brake and Motor-Fan", "(5) Motor-Fan and Braking Resistor", "(6) Brake and Braking Resistor", "(7) Brake, Motor-Fan and Braking Resistor"})
        Me.options_v.Location = New System.Drawing.Point(723, 39)
        Me.options_v.Name = "options_v"
        Me.options_v.Size = New System.Drawing.Size(384, 204)
        Me.options_v.TabIndex = 19
        Me.options_v.Visible = False
        '
        'Enter_Panel_info
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 23.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1119, 297)
        Me.Controls.Add(Me.options_v)
        Me.Controls.Add(Me.p_name3)
        Me.Controls.Add(Me.p_name2)
        Me.Controls.Add(Me.p_name1)
        Me.Controls.Add(Me.desc_box)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.q_box)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximumSize = New System.Drawing.Size(1137, 344)
        Me.Name = "Enter_Panel_info"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Panel Info"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents q_box As TextBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents desc_box As TextBox
    Friend WithEvents p_name1 As ComboBox
    Friend WithEvents p_name2 As ComboBox
    Friend WithEvents p_name3 As ComboBox
    Friend WithEvents options_v As CheckedListBox
End Class
