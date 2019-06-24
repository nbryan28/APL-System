<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Total_Project_Cost
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Total_Project_Cost))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.total_box = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ship_box = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.bulk_box = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ar_box = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ps_box = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.gi_box = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.job_box = New System.Windows.Forms.ComboBox()
        Me.Records_found = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.status_l = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(68, Byte), Integer))
        Me.GroupBox1.Controls.Add(Me.total_box)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.ship_box)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.bulk_box)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.ar_box)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ps_box)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.gi_box)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(26, 116)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1339, 670)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'total_box
        '
        Me.total_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.total_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.total_box.Location = New System.Drawing.Point(560, 579)
        Me.total_box.Name = "total_box"
        Me.total_box.ReadOnly = True
        Me.total_box.Size = New System.Drawing.Size(439, 31)
        Me.total_box.TabIndex = 23
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI Semibold", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.LightBlue
        Me.Label6.Location = New System.Drawing.Point(328, 579)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(192, 28)
        Me.Label6.TabIndex = 22
        Me.Label6.Text = "Total Project Cost: $"
        '
        'ship_box
        '
        Me.ship_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ship_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ship_box.Location = New System.Drawing.Point(560, 465)
        Me.ship_box.Name = "ship_box"
        Me.ship_box.Size = New System.Drawing.Size(439, 31)
        Me.ship_box.TabIndex = 21
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label5.Location = New System.Drawing.Point(372, 471)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(148, 25)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Shipping Cost: $"
        '
        'bulk_box
        '
        Me.bulk_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.bulk_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bulk_box.Location = New System.Drawing.Point(560, 364)
        Me.bulk_box.Name = "bulk_box"
        Me.bulk_box.Size = New System.Drawing.Size(439, 31)
        Me.bulk_box.TabIndex = 19
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label4.Location = New System.Drawing.Point(411, 367)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 25)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Bulk Cost: $"
        '
        'ar_box
        '
        Me.ar_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ar_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ar_box.Location = New System.Drawing.Point(560, 260)
        Me.ar_box.Name = "ar_box"
        Me.ar_box.ReadOnly = True
        Me.ar_box.Size = New System.Drawing.Size(439, 31)
        Me.ar_box.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(358, 260)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(162, 25)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Add and Return: $"
        '
        'ps_box
        '
        Me.ps_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ps_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ps_box.Location = New System.Drawing.Point(560, 167)
        Me.ps_box.Name = "ps_box"
        Me.ps_box.ReadOnly = True
        Me.ps_box.Size = New System.Drawing.Size(439, 31)
        Me.ps_box.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(360, 167)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(160, 25)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Project Specific: $"
        '
        'gi_box
        '
        Me.gi_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.gi_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gi_box.Location = New System.Drawing.Point(560, 83)
        Me.gi_box.Name = "gi_box"
        Me.gi_box.ReadOnly = True
        Me.gi_box.Size = New System.Drawing.Size(439, 31)
        Me.gi_box.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(339, 83)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(181, 25)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "General Inventory: $"
        '
        'job_box
        '
        Me.job_box.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.job_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.job_box.FormattingEnabled = True
        Me.job_box.Location = New System.Drawing.Point(162, 39)
        Me.job_box.Name = "job_box"
        Me.job_box.Size = New System.Drawing.Size(368, 33)
        Me.job_box.TabIndex = 1
        '
        'Records_found
        '
        Me.Records_found.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Records_found.AutoSize = True
        Me.Records_found.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Records_found.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Records_found.Location = New System.Drawing.Point(21, 42)
        Me.Records_found.Name = "Records_found"
        Me.Records_found.Size = New System.Drawing.Size(135, 25)
        Me.Records_found.TabIndex = 11
        Me.Records_found.Text = "Select Project: "
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(953, 29)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(152, 53)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Apply Changes"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label33
        '
        Me.Label33.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label33.BackColor = System.Drawing.Color.Transparent
        Me.Label33.ForeColor = System.Drawing.Color.White
        Me.Label33.Image = CType(resources.GetObject("Label33.Image"), System.Drawing.Image)
        Me.Label33.Location = New System.Drawing.Point(1320, 28)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(45, 39)
        Me.Label33.TabIndex = 63
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.AutoSize = True
        Me.Button1.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(728, 31)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(172, 51)
        Me.Button1.TabIndex = 64
        Me.Button1.Text = "Export to Excel"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'status_l
        '
        Me.status_l.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.status_l.AutoSize = True
        Me.status_l.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.status_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.status_l.Location = New System.Drawing.Point(21, 88)
        Me.status_l.Name = "status_l"
        Me.status_l.Size = New System.Drawing.Size(66, 25)
        Me.status_l.TabIndex = 24
        Me.status_l.Text = "Status:"
        '
        'Total_Project_Cost
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1402, 798)
        Me.Controls.Add(Me.status_l)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Records_found)
        Me.Controls.Add(Me.job_box)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Total_Project_Cost"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Total Project Cost"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents job_box As ComboBox
    Friend WithEvents Records_found As Label
    Friend WithEvents gi_box As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents total_box As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents ship_box As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents bulk_box As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents ar_box As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ps_box As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label33 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents status_l As Label
End Class
