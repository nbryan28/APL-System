<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Barc_Search
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Barc_Search))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox11 = New System.Windows.Forms.TextBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.part_t = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.manu_t = New System.Windows.Forms.TextBox()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.desc_t = New System.Windows.Forms.TextBox()
        Me.not_label = New System.Windows.Forms.Label()
        Me.notes_t = New System.Windows.Forms.TextBox()
        Me.mfg_type_t = New System.Windows.Forms.TextBox()
        Me.status_t = New System.Windows.Forms.TextBox()
        Me.type_t = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.units_t = New System.Windows.Forms.TextBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.Min_t = New System.Windows.Forms.TextBox()
        Me.legacy_t = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.alt_grid = New System.Windows.Forms.DataGridView()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ADA_number = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.name_box = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.alt_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Gainsboro
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Image)
        Me.Label1.Location = New System.Drawing.Point(12, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 71)
        Me.Label1.TabIndex = 0
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.LightGray
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(132, 23)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(262, 31)
        Me.TextBox1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Gainsboro
        Me.Label2.Location = New System.Drawing.Point(128, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(281, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Please Scan Label to Retrieve Data "
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(15, 132)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1807, 836)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 32)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1799, 800)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = " Part Info "
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 6)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1765, 773)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Controls.Add(Me.Button7)
        Me.GroupBox3.Controls.Add(Me.TextBox9)
        Me.GroupBox3.Controls.Add(Me.TextBox7)
        Me.GroupBox3.Controls.Add(Me.TextBox11)
        Me.GroupBox3.Controls.Add(Me.Label46)
        Me.GroupBox3.Controls.Add(Me.Label42)
        Me.GroupBox3.Controls.Add(Me.Label43)
        Me.GroupBox3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox3.Location = New System.Drawing.Point(563, 406)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1181, 334)
        Me.GroupBox3.TabIndex = 60
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Quantities"
        '
        'Button3
        '
        Me.Button3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button3.BackColor = System.Drawing.Color.CadetBlue
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(810, 248)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(180, 54)
        Me.Button3.TabIndex = 58
        Me.Button3.Text = "Generate Barcode"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.BackColor = System.Drawing.Color.Gray
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Button1.Location = New System.Drawing.Point(810, 152)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(172, 54)
        Me.Button1.TabIndex = 57
        Me.Button1.Text = "Check Part Out"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button7.BackColor = System.Drawing.Color.Teal
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.ForeColor = System.Drawing.Color.Black
        Me.Button7.Location = New System.Drawing.Point(810, 50)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(172, 54)
        Me.Button7.TabIndex = 56
        Me.Button7.Text = "Check Part In"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'TextBox9
        '
        Me.TextBox9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox9.BackColor = System.Drawing.Color.LightGray
        Me.TextBox9.Location = New System.Drawing.Point(391, 137)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.ReadOnly = True
        Me.TextBox9.Size = New System.Drawing.Size(135, 30)
        Me.TextBox9.TabIndex = 50
        '
        'TextBox7
        '
        Me.TextBox7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox7.BackColor = System.Drawing.Color.LightGray
        Me.TextBox7.Location = New System.Drawing.Point(391, 204)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = True
        Me.TextBox7.Size = New System.Drawing.Size(272, 30)
        Me.TextBox7.TabIndex = 53
        '
        'TextBox11
        '
        Me.TextBox11.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.TextBox11.BackColor = System.Drawing.Color.LightGray
        Me.TextBox11.Location = New System.Drawing.Point(391, 67)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.ReadOnly = True
        Me.TextBox11.Size = New System.Drawing.Size(135, 30)
        Me.TextBox11.TabIndex = 55
        '
        'Label46
        '
        Me.Label46.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label46.AutoSize = True
        Me.Label46.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label46.Location = New System.Drawing.Point(223, 67)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(125, 25)
        Me.Label46.TabIndex = 54
        Me.Label46.Text = "Inventory Qty"
        '
        'Label42
        '
        Me.Label42.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label42.AutoSize = True
        Me.Label42.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label42.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label42.Location = New System.Drawing.Point(264, 204)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(84, 25)
        Me.Label42.TabIndex = 51
        Me.Label42.Text = "Location"
        '
        'Label43
        '
        Me.Label43.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label43.AutoSize = True
        Me.Label43.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label43.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label43.Location = New System.Drawing.Point(211, 137)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(137, 25)
        Me.Label43.TabIndex = 49
        Me.Label43.Text = "Re-Order Level"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.InitialImage = CType(resources.GetObject("PictureBox1.InitialImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(60, 418)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(422, 322)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 59
        Me.PictureBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.part_t)
        Me.GroupBox2.Controls.Add(Me.Label40)
        Me.GroupBox2.Controls.Add(Me.Label39)
        Me.GroupBox2.Controls.Add(Me.Label37)
        Me.GroupBox2.Controls.Add(Me.manu_t)
        Me.GroupBox2.Controls.Add(Me.Label36)
        Me.GroupBox2.Controls.Add(Me.desc_t)
        Me.GroupBox2.Controls.Add(Me.not_label)
        Me.GroupBox2.Controls.Add(Me.notes_t)
        Me.GroupBox2.Controls.Add(Me.mfg_type_t)
        Me.GroupBox2.Controls.Add(Me.status_t)
        Me.GroupBox2.Controls.Add(Me.type_t)
        Me.GroupBox2.Controls.Add(Me.Label35)
        Me.GroupBox2.Controls.Add(Me.units_t)
        Me.GroupBox2.Controls.Add(Me.Label38)
        Me.GroupBox2.Controls.Add(Me.Label41)
        Me.GroupBox2.Controls.Add(Me.Min_t)
        Me.GroupBox2.Controls.Add(Me.legacy_t)
        Me.GroupBox2.Controls.Add(Me.Label34)
        Me.GroupBox2.Controls.Add(Me.Label29)
        Me.GroupBox2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox2.Location = New System.Drawing.Point(27, 40)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1717, 346)
        Me.GroupBox2.TabIndex = 58
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Specifications"
        '
        'part_t
        '
        Me.part_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.part_t.BackColor = System.Drawing.Color.LightGray
        Me.part_t.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.part_t.Location = New System.Drawing.Point(394, 54)
        Me.part_t.Multiline = True
        Me.part_t.Name = "part_t"
        Me.part_t.ReadOnly = True
        Me.part_t.Size = New System.Drawing.Size(311, 30)
        Me.part_t.TabIndex = 23
        '
        'Label40
        '
        Me.Label40.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label40.AutoSize = True
        Me.Label40.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label40.Location = New System.Drawing.Point(242, 53)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(119, 25)
        Me.Label40.TabIndex = 22
        Me.Label40.Text = "Part Number"
        '
        'Label39
        '
        Me.Label39.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label39.AutoSize = True
        Me.Label39.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label39.Location = New System.Drawing.Point(258, 107)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(100, 25)
        Me.Label39.TabIndex = 24
        Me.Label39.Text = "Part Status"
        '
        'Label37
        '
        Me.Label37.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label37.AutoSize = True
        Me.Label37.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label37.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label37.Location = New System.Drawing.Point(236, 161)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(127, 25)
        Me.Label37.TabIndex = 28
        Me.Label37.Text = "Manufacturer"
        '
        'manu_t
        '
        Me.manu_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.manu_t.BackColor = System.Drawing.Color.LightGray
        Me.manu_t.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.manu_t.Location = New System.Drawing.Point(394, 161)
        Me.manu_t.Name = "manu_t"
        Me.manu_t.ReadOnly = True
        Me.manu_t.Size = New System.Drawing.Size(237, 31)
        Me.manu_t.TabIndex = 29
        '
        'Label36
        '
        Me.Label36.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label36.AutoSize = True
        Me.Label36.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label36.Location = New System.Drawing.Point(253, 214)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(108, 25)
        Me.Label36.TabIndex = 30
        Me.Label36.Text = "Description"
        '
        'desc_t
        '
        Me.desc_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.desc_t.BackColor = System.Drawing.Color.LightGray
        Me.desc_t.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.desc_t.Location = New System.Drawing.Point(394, 214)
        Me.desc_t.Name = "desc_t"
        Me.desc_t.ReadOnly = True
        Me.desc_t.Size = New System.Drawing.Size(450, 31)
        Me.desc_t.TabIndex = 31
        '
        'not_label
        '
        Me.not_label.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.not_label.AutoSize = True
        Me.not_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.not_label.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.not_label.Location = New System.Drawing.Point(297, 263)
        Me.not_label.Name = "not_label"
        Me.not_label.Size = New System.Drawing.Size(64, 28)
        Me.not_label.TabIndex = 40
        Me.not_label.Text = "Notes"
        '
        'notes_t
        '
        Me.notes_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.notes_t.BackColor = System.Drawing.Color.LightGray
        Me.notes_t.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.notes_t.Location = New System.Drawing.Point(391, 263)
        Me.notes_t.Name = "notes_t"
        Me.notes_t.ReadOnly = True
        Me.notes_t.Size = New System.Drawing.Size(453, 31)
        Me.notes_t.TabIndex = 41
        '
        'mfg_type_t
        '
        Me.mfg_type_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.mfg_type_t.BackColor = System.Drawing.Color.LightGray
        Me.mfg_type_t.Location = New System.Drawing.Point(1159, 266)
        Me.mfg_type_t.Name = "mfg_type_t"
        Me.mfg_type_t.ReadOnly = True
        Me.mfg_type_t.Size = New System.Drawing.Size(289, 30)
        Me.mfg_type_t.TabIndex = 47
        '
        'status_t
        '
        Me.status_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.status_t.BackColor = System.Drawing.Color.LightGray
        Me.status_t.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.status_t.Location = New System.Drawing.Point(394, 107)
        Me.status_t.Name = "status_t"
        Me.status_t.ReadOnly = True
        Me.status_t.Size = New System.Drawing.Size(237, 31)
        Me.status_t.TabIndex = 44
        '
        'type_t
        '
        Me.type_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.type_t.BackColor = System.Drawing.Color.LightGray
        Me.type_t.Location = New System.Drawing.Point(1159, 161)
        Me.type_t.Name = "type_t"
        Me.type_t.ReadOnly = True
        Me.type_t.Size = New System.Drawing.Size(289, 30)
        Me.type_t.TabIndex = 46
        '
        'Label35
        '
        Me.Label35.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label35.AutoSize = True
        Me.Label35.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label35.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label35.Location = New System.Drawing.Point(995, 54)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(133, 25)
        Me.Label35.TabIndex = 32
        Me.Label35.Text = "Min Order Qty"
        '
        'units_t
        '
        Me.units_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.units_t.BackColor = System.Drawing.Color.LightGray
        Me.units_t.Location = New System.Drawing.Point(1159, 107)
        Me.units_t.Name = "units_t"
        Me.units_t.ReadOnly = True
        Me.units_t.Size = New System.Drawing.Size(87, 30)
        Me.units_t.TabIndex = 45
        '
        'Label38
        '
        Me.Label38.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label38.AutoSize = True
        Me.Label38.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label38.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label38.Location = New System.Drawing.Point(1028, 161)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(89, 25)
        Me.Label38.TabIndex = 26
        Me.Label38.Text = "Part Type"
        '
        'Label41
        '
        Me.Label41.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label41.AutoSize = True
        Me.Label41.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label41.Location = New System.Drawing.Point(955, 266)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(171, 25)
        Me.Label41.TabIndex = 42
        Me.Label41.Text = "Manufacturer Type"
        '
        'Min_t
        '
        Me.Min_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Min_t.BackColor = System.Drawing.Color.LightGray
        Me.Min_t.Location = New System.Drawing.Point(1159, 53)
        Me.Min_t.Name = "Min_t"
        Me.Min_t.ReadOnly = True
        Me.Min_t.Size = New System.Drawing.Size(87, 30)
        Me.Min_t.TabIndex = 33
        '
        'legacy_t
        '
        Me.legacy_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.legacy_t.BackColor = System.Drawing.Color.LightGray
        Me.legacy_t.Location = New System.Drawing.Point(1159, 214)
        Me.legacy_t.Name = "legacy_t"
        Me.legacy_t.ReadOnly = True
        Me.legacy_t.Size = New System.Drawing.Size(289, 30)
        Me.legacy_t.TabIndex = 39
        '
        'Label34
        '
        Me.Label34.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label34.AutoSize = True
        Me.Label34.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label34.Location = New System.Drawing.Point(1069, 107)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(57, 28)
        Me.Label34.TabIndex = 34
        Me.Label34.Text = "Units"
        '
        'Label29
        '
        Me.Label29.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label29.Location = New System.Drawing.Point(994, 213)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(123, 25)
        Me.Label29.TabIndex = 38
        Me.Label29.Text = "ADA Number"
        '
        'TabPage3
        '
        Me.TabPage3.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.TabPage3.Controls.Add(Me.alt_grid)
        Me.TabPage3.Location = New System.Drawing.Point(4, 32)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage3.Size = New System.Drawing.Size(1799, 800)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "  Alternatives  "
        '
        'alt_grid
        '
        Me.alt_grid.AllowUserToAddRows = False
        Me.alt_grid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.alt_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.alt_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.alt_grid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.alt_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.alt_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Teal
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.alt_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.alt_grid.ColumnHeadersHeight = 48
        Me.alt_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.alt_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column10, Me.Column2, Me.ADA_number, Me.Column14, Me.Column1})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.alt_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.alt_grid.EnableHeadersVisualStyles = False
        Me.alt_grid.Location = New System.Drawing.Point(21, 37)
        Me.alt_grid.MultiSelect = False
        Me.alt_grid.Name = "alt_grid"
        Me.alt_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.alt_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.alt_grid.RowHeadersWidth = 16
        Me.alt_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.alt_grid.RowTemplate.Height = 54
        Me.alt_grid.Size = New System.Drawing.Size(1752, 653)
        Me.alt_grid.TabIndex = 16
        '
        'Column10
        '
        Me.Column10.HeaderText = "Part No"
        Me.Column10.Name = "Column10"
        Me.Column10.Width = 400
        '
        'Column2
        '
        Me.Column2.HeaderText = "Description"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 400
        '
        'ADA_number
        '
        Me.ADA_number.HeaderText = "ADA Number"
        Me.ADA_number.Name = "ADA_number"
        Me.ADA_number.Width = 330
        '
        'Column14
        '
        Me.Column14.HeaderText = "  Inventory Qty  "
        Me.Column14.Name = "Column14"
        Me.Column14.Width = 200
        '
        'Column1
        '
        Me.Column1.HeaderText = "Manufacturer"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 300
        '
        'name_box
        '
        Me.name_box.BackColor = System.Drawing.Color.Gainsboro
        Me.name_box.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.name_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.name_box.Location = New System.Drawing.Point(1123, 69)
        Me.name_box.Multiline = True
        Me.name_box.Name = "name_box"
        Me.name_box.Size = New System.Drawing.Size(387, 31)
        Me.name_box.TabIndex = 31
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Gainsboro
        Me.Label3.Location = New System.Drawing.Point(948, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(154, 23)
        Me.Label3.TabIndex = 32
        Me.Label3.Text = "Enter Part Number"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(1541, 58)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 49)
        Me.Button2.TabIndex = 33
        Me.Button2.Text = "Search"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label33
        '
        Me.Label33.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label33.BackColor = System.Drawing.Color.Transparent
        Me.Label33.ForeColor = System.Drawing.Color.White
        Me.Label33.Image = CType(resources.GetObject("Label33.Image"), System.Drawing.Image)
        Me.Label33.Location = New System.Drawing.Point(1777, 15)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(45, 39)
        Me.Label33.TabIndex = 62
        '
        'Barc_Search
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1849, 980)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.name_box)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Barc_Search"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Quick Inventory Search"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        CType(Me.alt_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents name_box As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TextBox11 As TextBox
    Friend WithEvents Label46 As Label
    Friend WithEvents TextBox7 As TextBox
    Friend WithEvents Label42 As Label
    Friend WithEvents TextBox9 As TextBox
    Friend WithEvents Label43 As Label
    Friend WithEvents mfg_type_t As TextBox
    Friend WithEvents type_t As TextBox
    Friend WithEvents units_t As TextBox
    Friend WithEvents status_t As TextBox
    Friend WithEvents Label41 As Label
    Friend WithEvents notes_t As TextBox
    Friend WithEvents not_label As Label
    Friend WithEvents legacy_t As TextBox
    Friend WithEvents Label29 As Label
    Friend WithEvents Label34 As Label
    Friend WithEvents Min_t As TextBox
    Friend WithEvents Label35 As Label
    Friend WithEvents desc_t As TextBox
    Friend WithEvents Label36 As Label
    Friend WithEvents manu_t As TextBox
    Friend WithEvents Label37 As Label
    Friend WithEvents Label38 As Label
    Friend WithEvents Label39 As Label
    Friend WithEvents part_t As TextBox
    Friend WithEvents Label40 As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label33 As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents alt_grid As DataGridView
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents ADA_number As DataGridViewTextBoxColumn
    Friend WithEvents Column14 As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
End Class
