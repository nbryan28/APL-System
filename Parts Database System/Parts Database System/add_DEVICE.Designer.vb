<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class add_DEVICE
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(add_DEVICE))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.bulk_c = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.labor_c = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.legacy = New System.Windows.Forms.TextBox()
        Me.advName = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.descr = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.DimGray
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(12, 96)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1526, 747)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.DarkGray
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.bulk_c)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.labor_c)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.legacy)
        Me.GroupBox1.Controls.Add(Me.advName)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.descr)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(15, 33)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1483, 663)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter Device data"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(538, 338)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(95, 23)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "Bulk Cost $"
        '
        'bulk_c
        '
        Me.bulk_c.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.bulk_c.Location = New System.Drawing.Point(678, 335)
        Me.bulk_c.Name = "bulk_c"
        Me.bulk_c.Size = New System.Drawing.Size(326, 30)
        Me.bulk_c.TabIndex = 45
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(509, 251)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(106, 23)
        Me.Label4.TabIndex = 44
        Me.Label4.Text = "Labor Cost $"
        '
        'labor_c
        '
        Me.labor_c.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.labor_c.Location = New System.Drawing.Point(678, 251)
        Me.labor_c.Name = "labor_c"
        Me.labor_c.Size = New System.Drawing.Size(329, 30)
        Me.labor_c.TabIndex = 43
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(483, 107)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(132, 23)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Assembly Name"
        '
        'legacy
        '
        Me.legacy.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.legacy.Location = New System.Drawing.Point(673, 107)
        Me.legacy.Name = "legacy"
        Me.legacy.Size = New System.Drawing.Size(334, 30)
        Me.legacy.TabIndex = 41
        '
        'advName
        '
        Me.advName.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.advName.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.advName.Location = New System.Drawing.Point(34, 547)
        Me.advName.Name = "advName"
        Me.advName.ReadOnly = True
        Me.advName.Size = New System.Drawing.Size(291, 30)
        Me.advName.TabIndex = 36
        '
        'Label15
        '
        Me.Label15.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(30, 502)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(235, 23)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "For internal purposes (ignore)"
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.Location = New System.Drawing.Point(720, 440)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(177, 51)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Add Device"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(519, 182)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 23)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Description"
        '
        'descr
        '
        Me.descr.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.descr.Location = New System.Drawing.Point(673, 179)
        Me.descr.Name = "descr"
        Me.descr.Size = New System.Drawing.Size(331, 30)
        Me.descr.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(639, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(147, 28)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Add ASSEMBLY"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.Location = New System.Drawing.Point(1481, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 39)
        Me.Label2.TabIndex = 8
        '
        'add_DEVICE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1557, 855)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "add_DEVICE"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add Device"
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents advName As TextBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents descr As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents legacy As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents bulk_c As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents labor_c As TextBox
    Friend WithEvents Panel1 As Panel
End Class
