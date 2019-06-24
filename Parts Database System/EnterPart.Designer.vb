<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EnterPart
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EnterPart))
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.notes_t = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.legacy_t = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.unit_t = New System.Windows.Forms.ComboBox()
        Me.min_t = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.desc_t = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.status_t = New System.Windows.Forms.ComboBox()
        Me.Type_t = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.part_t = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.manu_t = New System.Windows.Forms.ComboBox()
        Me.p_vendor_t = New System.Windows.Forms.ComboBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(42, 80)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1239, 683)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage1.Location = New System.Drawing.Point(4, 32)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1231, 647)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = " Parts Data  "
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.Silver
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(3, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1222, 635)
        Me.Panel1.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.BackColor = System.Drawing.Color.LightGray
        Me.GroupBox2.Controls.Add(Me.p_vendor_t)
        Me.GroupBox2.Controls.Add(Me.manu_t)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.ComboBox1)
        Me.GroupBox2.Controls.Add(Me.notes_t)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.legacy_t)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.unit_t)
        Me.GroupBox2.Controls.Add(Me.min_t)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.desc_t)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Location = New System.Drawing.Point(19, 287)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1168, 331)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Additional data"
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(847, 88)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(85, 23)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "MFG Type"
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Panel", "Field", "Assembly", "Bulk"})
        Me.ComboBox1.Location = New System.Drawing.Point(968, 84)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(130, 31)
        Me.ComboBox1.TabIndex = 19
        '
        'notes_t
        '
        Me.notes_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.notes_t.Location = New System.Drawing.Point(175, 138)
        Me.notes_t.Name = "notes_t"
        Me.notes_t.Size = New System.Drawing.Size(586, 30)
        Me.notes_t.TabIndex = 18
        '
        'Label17
        '
        Me.Label17.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(80, 138)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(55, 23)
        Me.Label17.TabIndex = 17
        Me.Label17.Text = "Notes"
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.Location = New System.Drawing.Point(530, 263)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(121, 42)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Add Part"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(582, 198)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(179, 23)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Primary Vendor Name"
        '
        'legacy_t
        '
        Me.legacy_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.legacy_t.Location = New System.Drawing.Point(175, 195)
        Me.legacy_t.Name = "legacy_t"
        Me.legacy_t.Size = New System.Drawing.Size(357, 30)
        Me.legacy_t.TabIndex = 16
        '
        'Label16
        '
        Me.Label16.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(39, 198)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(112, 23)
        Me.Label16.TabIndex = 15
        Me.Label16.Text = "ADA Number"
        '
        'Label14
        '
        Me.Label14.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(883, 37)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(49, 23)
        Me.Label14.TabIndex = 14
        Me.Label14.Text = "Units"
        '
        'unit_t
        '
        Me.unit_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.unit_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.unit_t.FormattingEnabled = True
        Me.unit_t.Location = New System.Drawing.Point(968, 29)
        Me.unit_t.Name = "unit_t"
        Me.unit_t.Size = New System.Drawing.Size(130, 31)
        Me.unit_t.TabIndex = 13
        '
        'min_t
        '
        Me.min_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.min_t.Location = New System.Drawing.Point(968, 138)
        Me.min_t.Name = "min_t"
        Me.min_t.Size = New System.Drawing.Size(130, 30)
        Me.min_t.TabIndex = 6
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(812, 138)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 23)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Min Order Qty"
        '
        'desc_t
        '
        Me.desc_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.desc_t.Location = New System.Drawing.Point(175, 85)
        Me.desc_t.Name = "desc_t"
        Me.desc_t.Size = New System.Drawing.Size(586, 30)
        Me.desc_t.TabIndex = 4
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(39, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 23)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Description"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(39, 41)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(113, 23)
        Me.Label6.TabIndex = 1
        Me.Label6.Text = "Manufacturer"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.LightGray
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.status_t)
        Me.GroupBox1.Controls.Add(Me.Type_t)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.part_t)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(19, 21)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1168, 285)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Primary Information"
        '
        'Label13
        '
        Me.Label13.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label13.Image = CType(resources.GetObject("Label13.Image"), System.Drawing.Image)
        Me.Label13.Location = New System.Drawing.Point(805, 26)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(204, 200)
        Me.Label13.TabIndex = 7
        '
        'status_t
        '
        Me.status_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.status_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.status_t.FormattingEnabled = True
        Me.status_t.Items.AddRange(New Object() {"", "Preferred", "Special Order"})
        Me.status_t.Location = New System.Drawing.Point(370, 101)
        Me.status_t.Name = "status_t"
        Me.status_t.Size = New System.Drawing.Size(298, 31)
        Me.status_t.TabIndex = 6
        '
        'Type_t
        '
        Me.Type_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Type_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Type_t.FormattingEnabled = True
        Me.Type_t.Location = New System.Drawing.Point(370, 157)
        Me.Type_t.Name = "Type_t"
        Me.Type_t.Size = New System.Drawing.Size(298, 31)
        Me.Type_t.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(238, 157)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(80, 23)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Part Type"
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(238, 101)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Part Status"
        '
        'part_t
        '
        Me.part_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.part_t.Location = New System.Drawing.Point(370, 50)
        Me.part_t.Name = "part_t"
        Me.part_t.Size = New System.Drawing.Size(349, 30)
        Me.part_t.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.LightGray
        Me.Label3.Location = New System.Drawing.Point(168, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(161, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Part Name/Number"
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(525, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 28)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Enter New Part Information"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.ForeColor = System.Drawing.Color.DimGray
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.Location = New System.Drawing.Point(1200, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 37)
        Me.Label2.TabIndex = 2
        '
        'manu_t
        '
        Me.manu_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.manu_t.FormattingEnabled = True
        Me.manu_t.Items.AddRange(New Object() {"ACME", "ACS", "Advanced Antivibration Componants", "Advantage Industrial", "AIC", "ALLEN BRADLEY", "ARISTA", "Atronix", "Atronix/MHS", "Automation Direct", "AWC", "B&D", "Banner", "Brady", "Bussman", "Eaton", "Fastenal", "FastSigns", "Fibox", "Fuseco", "GA Automation", "GCC", "GOLDX", "GRACEPORT", "Graphic Products", "Heyco", "HOFFMAN", "Honeywell", "HUBBEL-WIEGMANN", "Hytrol", "Ideal", "IDEC", "Integra", "ISLATROL", "LEVITON", "Marathon", "McNaughton-McKay", "Mencom", "Mersen", "METALUX", "Metro Metal Works", "MetroWire", "MHS", "MMW", "Mouser", "Murr", "Omron", "OPTO22", "Panduit", "Pepperl&Fuchs", "PHILIPS", "Phoenix Contact", "Photocraft", "Planet", "PULS", "Rittal", "SherTek", "Siemens", "SquareD", "TECHCONTROLS", "TRENDnet", "Turck", "UL", "Wago"})
        Me.manu_t.Location = New System.Drawing.Point(175, 34)
        Me.manu_t.Name = "manu_t"
        Me.manu_t.Size = New System.Drawing.Size(370, 31)
        Me.manu_t.TabIndex = 21
        '
        'p_vendor_t
        '
        Me.p_vendor_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_vendor_t.FormattingEnabled = True
        Me.p_vendor_t.Items.AddRange(New Object() {"ACME", "ACS", "Advanced Antivibration Componants", "Advantage Industrial", "AIC", "ALLEN BRADLEY", "ARISTA", "Atronix", "Atronix/MHS", "Automation Direct", "AWC", "B&D", "Banner", "Brady", "Bussman", "Eaton", "Fastenal", "FastSigns", "Fibox", "Fuseco", "GA Automation", "GCC", "GOLDX", "GRACEPORT", "Graphic Products", "Heyco", "HOFFMAN", "Honeywell", "HUBBEL-WIEGMANN", "Hytrol", "Ideal", "IDEC", "Integra", "ISLATROL", "LEVITON", "Marathon", "McNaughton-McKay", "Mencom", "Mersen", "METALUX", "Metro Metal Works", "MetroWire", "MHS", "MMW", "Mouser", "Murr", "Omron", "OPTO22", "Panduit", "Pepperl&Fuchs", "PHILIPS", "Phoenix Contact", "Photocraft", "Planet", "PULS", "Rittal", "SherTek", "Siemens", "SquareD", "TECHCONTROLS", "TRENDnet", "Turck", "UL", "Wago"})
        Me.p_vendor_t.Location = New System.Drawing.Point(777, 198)
        Me.p_vendor_t.Name = "p_vendor_t"
        Me.p_vendor_t.Size = New System.Drawing.Size(321, 31)
        Me.p_vendor_t.TabIndex = 22
        '
        'EnterPart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1311, 788)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TabControl1)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EnterPart"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add New Part"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Panel1 As Panel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents status_t As ComboBox
    Friend WithEvents Type_t As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents part_t As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents notes_t As TextBox
    Friend WithEvents Label17 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents legacy_t As TextBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents unit_t As ComboBox
    Friend WithEvents min_t As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents desc_t As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents manu_t As ComboBox
    Friend WithEvents p_vendor_t As ComboBox
End Class
