<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Sent_mail
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Sent_mail))
        Me.mail_t = New System.Windows.Forms.TextBox()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.user_box = New System.Windows.Forms.ComboBox()
        Me.title_t = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'mail_t
        '
        Me.mail_t.BackColor = System.Drawing.Color.Silver
        Me.mail_t.Location = New System.Drawing.Point(18, 179)
        Me.mail_t.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.mail_t.Multiline = True
        Me.mail_t.Name = "mail_t"
        Me.mail_t.Size = New System.Drawing.Size(971, 259)
        Me.mail_t.TabIndex = 0
        '
        'Button14
        '
        Me.Button14.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button14.BackColor = System.Drawing.Color.DarkGray
        Me.Button14.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button14.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button14.ForeColor = System.Drawing.Color.Black
        Me.Button14.Location = New System.Drawing.Point(406, 463)
        Me.Button14.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(171, 45)
        Me.Button14.TabIndex = 13
        Me.Button14.Text = "Send"
        Me.Button14.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(30, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 25)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "To"
        '
        'user_box
        '
        Me.user_box.BackColor = System.Drawing.Color.LightGray
        Me.user_box.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.user_box.FormattingEnabled = True
        Me.user_box.Location = New System.Drawing.Point(79, 37)
        Me.user_box.Name = "user_box"
        Me.user_box.Size = New System.Drawing.Size(348, 33)
        Me.user_box.TabIndex = 15
        '
        'title_t
        '
        Me.title_t.BackColor = System.Drawing.Color.Gainsboro
        Me.title_t.Location = New System.Drawing.Point(79, 112)
        Me.title_t.Multiline = True
        Me.title_t.Name = "title_t"
        Me.title_t.Size = New System.Drawing.Size(910, 33)
        Me.title_t.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(13, 115)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 25)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Title: "
        '
        'Sent_mail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1008, 526)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.title_t)
        Me.Controls.Add(Me.user_box)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.mail_t)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximumSize = New System.Drawing.Size(1026, 573)
        Me.Name = "Sent_mail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sent Message"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents mail_t As TextBox
    Friend WithEvents Button14 As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents user_box As ComboBox
    Friend WithEvents title_t As TextBox
    Friend WithEvents Label2 As Label
End Class
