<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Job_description
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Job_description))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.job_name = New System.Windows.Forms.TextBox()
        Me.desc_job = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.a_ship = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ship_a = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.CPM = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.onsite = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.onsite)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.CPM)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.a_ship)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.ship_a)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.desc_job)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.job_name)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Location = New System.Drawing.Point(12, 21)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1008, 553)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Project Information"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(316, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 23)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Job Number"
        '
        'job_name
        '
        Me.job_name.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.job_name.Location = New System.Drawing.Point(447, 65)
        Me.job_name.Name = "job_name"
        Me.job_name.ReadOnly = True
        Me.job_name.Size = New System.Drawing.Size(219, 31)
        Me.job_name.TabIndex = 2
        '
        'desc_job
        '
        Me.desc_job.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.desc_job.Location = New System.Drawing.Point(229, 143)
        Me.desc_job.Name = "desc_job"
        Me.desc_job.ReadOnly = True
        Me.desc_job.Size = New System.Drawing.Size(544, 31)
        Me.desc_job.TabIndex = 13
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(103, 148)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 23)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "Description"
        '
        'a_ship
        '
        Me.a_ship.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.a_ship.Location = New System.Drawing.Point(229, 282)
        Me.a_ship.Name = "a_ship"
        Me.a_ship.ReadOnly = True
        Me.a_ship.Size = New System.Drawing.Size(725, 31)
        Me.a_ship.TabIndex = 30
        '
        'Label16
        '
        Me.Label16.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(54, 287)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(145, 23)
        Me.Label16.TabIndex = 29
        Me.Label16.Text = "Alternate Address"
        '
        'ship_a
        '
        Me.ship_a.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ship_a.Location = New System.Drawing.Point(229, 213)
        Me.ship_a.Name = "ship_a"
        Me.ship_a.ReadOnly = True
        Me.ship_a.Size = New System.Drawing.Size(725, 31)
        Me.ship_a.TabIndex = 28
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(57, 213)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(142, 23)
        Me.Label9.TabIndex = 27
        Me.Label9.Text = "Shipping Address"
        '
        'CPM
        '
        Me.CPM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.CPM.Location = New System.Drawing.Point(229, 355)
        Me.CPM.Name = "CPM"
        Me.CPM.ReadOnly = True
        Me.CPM.Size = New System.Drawing.Size(521, 31)
        Me.CPM.TabIndex = 32
        '
        'Label19
        '
        Me.Label19.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(29, 358)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(172, 23)
        Me.Label19.TabIndex = 31
        Me.Label19.Text = "Customer Project PM"
        '
        'onsite
        '
        Me.onsite.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.onsite.Location = New System.Drawing.Point(229, 436)
        Me.onsite.Name = "onsite"
        Me.onsite.ReadOnly = True
        Me.onsite.Size = New System.Drawing.Size(521, 31)
        Me.onsite.TabIndex = 34
        '
        'Label18
        '
        Me.Label18.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(29, 436)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(180, 23)
        Me.Label18.TabIndex = 33
        Me.Label18.Text = "Onsite Contact Person"
        '
        'Job_description
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1032, 586)
        Me.Controls.Add(Me.GroupBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(1050, 633)
        Me.Name = "Job_description"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Project Information"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label3 As Label
    Friend WithEvents job_name As TextBox
    Friend WithEvents desc_job As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents a_ship As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents ship_a As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents CPM As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents onsite As TextBox
    Friend WithEvents Label18 As Label
End Class
