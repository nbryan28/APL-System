<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UR_part
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UR_part))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Update_grid = New System.Windows.Forms.DataGridView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.mfg_t = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.notes_t = New System.Windows.Forms.TextBox()
        Me.not_label = New System.Windows.Forms.Label()
        Me.legacy_t = New System.Windows.Forms.TextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.units_t = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Min_t = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.desc_t = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.type_t = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.status_t = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.part_t = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Part_textbox = New System.Windows.Forms.TextBox()
        Me.type_ur = New System.Windows.Forms.ComboBox()
        Me.manu_t = New System.Windows.Forms.ComboBox()
        Me.p_vendor_t = New System.Windows.Forms.ComboBox()
        Me.Panel1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.Update_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.TabControl1)
        Me.Panel1.Location = New System.Drawing.Point(12, 12)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1503, 854)
        Me.Panel1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(624, 27)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(184, 25)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Update/Remove Part"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Image)
        Me.Label1.Location = New System.Drawing.Point(1439, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 39)
        Me.Label1.TabIndex = 2
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(18, 71)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1470, 763)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.LightGray
        Me.TabPage1.Controls.Add(Me.Update_grid)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Location = New System.Drawing.Point(4, 32)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1462, 727)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = " Part Information "
        '
        'Update_grid
        '
        Me.Update_grid.AllowUserToAddRows = False
        Me.Update_grid.AllowUserToDeleteRows = False
        Me.Update_grid.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Update_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Update_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Update_grid.BackgroundColor = System.Drawing.Color.Silver
        Me.Update_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.Update_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Update_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.Update_grid.ColumnHeadersHeight = 48
        Me.Update_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Update_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.Update_grid.EnableHeadersVisualStyles = False
        Me.Update_grid.Location = New System.Drawing.Point(21, 399)
        Me.Update_grid.MultiSelect = False
        Me.Update_grid.Name = "Update_grid"
        Me.Update_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Update_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.Update_grid.RowHeadersWidth = 16
        Me.Update_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Update_grid.RowTemplate.Height = 44
        Me.Update_grid.Size = New System.Drawing.Size(1411, 305)
        Me.Update_grid.TabIndex = 13
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox1.BackColor = System.Drawing.Color.Silver
        Me.GroupBox1.Controls.Add(Me.p_vendor_t)
        Me.GroupBox1.Controls.Add(Me.manu_t)
        Me.GroupBox1.Controls.Add(Me.mfg_t)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Button7)
        Me.GroupBox1.Controls.Add(Me.notes_t)
        Me.GroupBox1.Controls.Add(Me.not_label)
        Me.GroupBox1.Controls.Add(Me.legacy_t)
        Me.GroupBox1.Controls.Add(Me.Label19)
        Me.GroupBox1.Controls.Add(Me.Label18)
        Me.GroupBox1.Controls.Add(Me.units_t)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Min_t)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.desc_t)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.type_t)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.status_t)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.part_t)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(353, 18)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1079, 357)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Part Data"
        '
        'mfg_t
        '
        Me.mfg_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.mfg_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.mfg_t.FormattingEnabled = True
        Me.mfg_t.Items.AddRange(New Object() {"Panel", "Field", "Assembly", "Bulk"})
        Me.mfg_t.Location = New System.Drawing.Point(674, 314)
        Me.mfg_t.Name = "mfg_t"
        Me.mfg_t.Size = New System.Drawing.Size(161, 31)
        Me.mfg_t.TabIndex = 24
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(570, 318)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(85, 23)
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "MFG Type"
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.DarkSlateGray
        Me.Button7.ForeColor = System.Drawing.Color.White
        Me.Button7.Location = New System.Drawing.Point(898, 298)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(150, 43)
        Me.Button7.TabIndex = 22
        Me.Button7.Text = "Create New Part"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'notes_t
        '
        Me.notes_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.notes_t.Location = New System.Drawing.Point(200, 311)
        Me.notes_t.Name = "notes_t"
        Me.notes_t.Size = New System.Drawing.Size(311, 30)
        Me.notes_t.TabIndex = 21
        '
        'not_label
        '
        Me.not_label.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.not_label.AutoSize = True
        Me.not_label.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.not_label.Location = New System.Drawing.Point(115, 311)
        Me.not_label.Name = "not_label"
        Me.not_label.Size = New System.Drawing.Size(55, 23)
        Me.not_label.TabIndex = 20
        Me.not_label.Text = "Notes"
        '
        'legacy_t
        '
        Me.legacy_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.legacy_t.Location = New System.Drawing.Point(610, 209)
        Me.legacy_t.Name = "legacy_t"
        Me.legacy_t.Size = New System.Drawing.Size(228, 30)
        Me.legacy_t.TabIndex = 19
        '
        'Label19
        '
        Me.Label19.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(532, 214)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(72, 23)
        Me.Label19.TabIndex = 18
        Me.Label19.Text = "ADA N#"
        '
        'Label18
        '
        Me.Label18.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(55, 209)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(128, 23)
        Me.Label18.TabIndex = 16
        Me.Label18.Text = "Primary Vendor"
        '
        'units_t
        '
        Me.units_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.units_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.units_t.FormattingEnabled = True
        Me.units_t.Items.AddRange(New Object() {"Each"})
        Me.units_t.Location = New System.Drawing.Point(674, 153)
        Me.units_t.Name = "units_t"
        Me.units_t.Size = New System.Drawing.Size(161, 31)
        Me.units_t.TabIndex = 15
        '
        'Label16
        '
        Me.Label16.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(606, 157)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(49, 23)
        Me.Label16.TabIndex = 14
        Me.Label16.Text = "Units"
        '
        'Button3
        '
        Me.Button3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button3.BackColor = System.Drawing.Color.DarkRed
        Me.Button3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.White
        Me.Button3.Location = New System.Drawing.Point(898, 137)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(129, 42)
        Me.Button3.TabIndex = 13
        Me.Button3.Text = "Remove Part"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(898, 57)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(129, 39)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Update Part"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Min_t
        '
        Me.Min_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Min_t.Location = New System.Drawing.Point(674, 52)
        Me.Min_t.Name = "Min_t"
        Me.Min_t.Size = New System.Drawing.Size(164, 30)
        Me.Min_t.TabIndex = 11
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(535, 55)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 23)
        Me.Label10.TabIndex = 10
        Me.Label10.Text = "Min Order Qty"
        '
        'desc_t
        '
        Me.desc_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.desc_t.Location = New System.Drawing.Point(203, 262)
        Me.desc_t.Name = "desc_t"
        Me.desc_t.Size = New System.Drawing.Size(635, 30)
        Me.desc_t.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(87, 262)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 23)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "Description"
        '
        'Label8
        '
        Me.Label8.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(70, 157)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(113, 23)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "Manufacturer"
        '
        'type_t
        '
        Me.type_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.type_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.type_t.FormattingEnabled = True
        Me.type_t.Location = New System.Drawing.Point(549, 97)
        Me.type_t.Name = "type_t"
        Me.type_t.Size = New System.Drawing.Size(289, 31)
        Me.type_t.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(453, 100)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 23)
        Me.Label7.TabIndex = 4
        Me.Label7.Text = "Part Type"
        '
        'status_t
        '
        Me.status_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.status_t.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.status_t.FormattingEnabled = True
        Me.status_t.Items.AddRange(New Object() {"", "Preferred", "Special Order"})
        Me.status_t.Location = New System.Drawing.Point(203, 100)
        Me.status_t.Name = "status_t"
        Me.status_t.Size = New System.Drawing.Size(161, 31)
        Me.status_t.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(92, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 23)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Part Status"
        '
        'part_t
        '
        Me.part_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.part_t.Location = New System.Drawing.Point(200, 52)
        Me.part_t.Name = "part_t"
        Me.part_t.Size = New System.Drawing.Size(311, 30)
        Me.part_t.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(22, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(161, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Part Number/Name"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox2.BackColor = System.Drawing.Color.Silver
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.Part_textbox)
        Me.GroupBox2.Controls.Add(Me.type_ur)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(21, 41)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(312, 314)
        Me.GroupBox2.TabIndex = 11
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Quick Search"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(20, 140)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 23)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Enter Part Name"
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(20, 50)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(130, 23)
        Me.Label6.TabIndex = 9
        Me.Label6.Text = "Select Part Type"
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(85, 235)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(142, 50)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Search"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Part_textbox
        '
        Me.Part_textbox.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Part_textbox.Location = New System.Drawing.Point(24, 178)
        Me.Part_textbox.Name = "Part_textbox"
        Me.Part_textbox.Size = New System.Drawing.Size(258, 30)
        Me.Part_textbox.TabIndex = 2
        '
        'type_ur
        '
        Me.type_ur.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.type_ur.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.type_ur.FormattingEnabled = True
        Me.type_ur.Items.AddRange(New Object() {"All Types"})
        Me.type_ur.Location = New System.Drawing.Point(24, 78)
        Me.type_ur.Name = "type_ur"
        Me.type_ur.Size = New System.Drawing.Size(258, 31)
        Me.type_ur.TabIndex = 1
        '
        'manu_t
        '
        Me.manu_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.manu_t.FormattingEnabled = True
        Me.manu_t.Items.AddRange(New Object() {"ACME", "ACS", "Advanced Antivibration Componants", "Advantage Industrial", "AIC", "ALLEN BRADLEY", "ARISTA", "Atronix", "Atronix/MHS", "Automation Direct", "AWC", "B&D", "Banner", "Brady", "Bussman", "Eaton", "Fastenal", "FastSigns", "Fibox", "Fuseco", "GA Automation", "GCC", "GOLDX", "GRACEPORT", "Graphic Products", "Heyco", "HOFFMAN", "Honeywell", "HUBBEL-WIEGMANN", "Hytrol", "Ideal", "IDEC", "Integra", "ISLATROL", "LEVITON", "Marathon", "McNaughton-McKay", "Mencom", "Mersen", "METALUX", "Metro Metal Works", "MetroWire", "MHS", "MMW", "Mouser", "Murr", "Omron", "OPTO22", "Panduit", "Pepperl&Fuchs", "PHILIPS", "Phoenix Contact", "Photocraft", "Planet", "PULS", "Rittal", "SherTek", "Siemens", "SquareD", "TECHCONTROLS", "TRENDnet", "Turck", "UL", "Wago"})
        Me.manu_t.Location = New System.Drawing.Point(200, 153)
        Me.manu_t.Name = "manu_t"
        Me.manu_t.Size = New System.Drawing.Size(354, 31)
        Me.manu_t.TabIndex = 25
        '
        'p_vendor_t
        '
        Me.p_vendor_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_vendor_t.FormattingEnabled = True
        Me.p_vendor_t.Items.AddRange(New Object() {"ACME", "ACS", "Advanced Antivibration Componants", "Advantage Industrial", "AIC", "ALLEN BRADLEY", "ARISTA", "Atronix", "Atronix/MHS", "Automation Direct", "AWC", "B&D", "Banner", "Brady", "Bussman", "Eaton", "Fastenal", "FastSigns", "Fibox", "Fuseco", "GA Automation", "GCC", "GOLDX", "GRACEPORT", "Graphic Products", "Heyco", "HOFFMAN", "Honeywell", "HUBBEL-WIEGMANN", "Hytrol", "Ideal", "IDEC", "Integra", "ISLATROL", "LEVITON", "Marathon", "McNaughton-McKay", "Mencom", "Mersen", "METALUX", "Metro Metal Works", "MetroWire", "MHS", "MMW", "Mouser", "Murr", "Omron", "OPTO22", "Panduit", "Pepperl&Fuchs", "PHILIPS", "Phoenix Contact", "Photocraft", "Planet", "PULS", "Rittal", "SherTek", "Siemens", "SquareD", "TECHCONTROLS", "TRENDnet", "Turck", "UL", "Wago"})
        Me.p_vendor_t.Location = New System.Drawing.Point(200, 206)
        Me.p_vendor_t.Name = "p_vendor_t"
        Me.p_vendor_t.Size = New System.Drawing.Size(286, 31)
        Me.p_vendor_t.TabIndex = 26
        '
        'UR_part
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(1527, 878)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "UR_part"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update/Remove part"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.Update_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Part_textbox As TextBox
    Friend WithEvents type_ur As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents type_t As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents status_t As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents part_t As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents desc_t As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Min_t As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents units_t As ComboBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents legacy_t As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents notes_t As TextBox
    Friend WithEvents not_label As Label
    Friend WithEvents Button7 As Button
    Friend WithEvents Update_grid As DataGridView
    Friend WithEvents mfg_t As ComboBox
    Friend WithEvents Label11 As Label
    Friend WithEvents manu_t As ComboBox
    Friend WithEvents p_vendor_t As ComboBox
End Class
