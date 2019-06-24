<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Edit_BOM_Panels
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Edit_BOM_Panels))
        Me.Panel_grid = New System.Windows.Forms.DataGridView()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column12 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Column19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column27 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.options_v = New System.Windows.Forms.CheckedListBox()
        Me.qty_b = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.p_name3 = New System.Windows.Forms.ComboBox()
        Me.p_name2 = New System.Windows.Forms.ComboBox()
        Me.p_name1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.panel_desc = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.assem_grid = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewComboBoxColumn1 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewComboBoxColumn2 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewComboBoxColumn3 = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportBOMToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToExcelToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Part_Picker1 = New Parts_database_system.Part_Picker()
        CType(Me.Panel_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.assem_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel_grid
        '
        Me.Panel_grid.AllowUserToDeleteRows = False
        Me.Panel_grid.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Panel_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Panel_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_grid.BackgroundColor = System.Drawing.Color.Gray
        Me.Panel_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.Panel_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Panel_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.Panel_grid.ColumnHeadersHeight = 48
        Me.Panel_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Panel_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column10, Me.Description, Me.Column1, Me.Column13, Me.Column14, Me.Column15, Me.Column16, Me.Column12, Me.Column3, Me.Column2, Me.Column19, Me.Column27})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Panel_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.Panel_grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.Panel_grid.EnableHeadersVisualStyles = False
        Me.Panel_grid.Location = New System.Drawing.Point(7, 272)
        Me.Panel_grid.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel_grid.Name = "Panel_grid"
        Me.Panel_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Panel_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.Panel_grid.RowHeadersWidth = 16
        Me.Panel_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Panel_grid.RowTemplate.Height = 54
        Me.Panel_grid.Size = New System.Drawing.Size(1797, 624)
        Me.Panel_grid.TabIndex = 16
        '
        'Column10
        '
        Me.Column10.Frozen = True
        Me.Column10.HeaderText = "Part_No"
        Me.Column10.Name = "Column10"
        Me.Column10.Width = 380
        '
        'Description
        '
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        Me.Description.Width = 330
        '
        'Column1
        '
        Me.Column1.HeaderText = "Manufacturer"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 200
        '
        'Column13
        '
        Me.Column13.HeaderText = "Vendor"
        Me.Column13.Name = "Column13"
        Me.Column13.Width = 200
        '
        'Column14
        '
        Me.Column14.HeaderText = "Price"
        Me.Column14.Name = "Column14"
        Me.Column14.Width = 150
        '
        'Column15
        '
        Me.Column15.HeaderText = "Qty"
        Me.Column15.Name = "Column15"
        Me.Column15.Width = 150
        '
        'Column16
        '
        Me.Column16.HeaderText = "Subtotal"
        Me.Column16.Name = "Column16"
        Me.Column16.Width = 150
        '
        'Column12
        '
        Me.Column12.HeaderText = "MFG Type"
        Me.Column12.Items.AddRange(New Object() {"", "Panel", "Field", "Assembly", "Bulk"})
        Me.Column12.Name = "Column12"
        Me.Column12.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column12.Width = 200
        '
        'Column3
        '
        Me.Column3.HeaderText = "(Special Order/Preferred)"
        Me.Column3.Items.AddRange(New Object() {"", "Preferred", "Special Order"})
        Me.Column3.Name = "Column3"
        Me.Column3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column3.Width = 160
        '
        'Column2
        '
        Me.Column2.HeaderText = "Part Type"
        Me.Column2.Items.AddRange(New Object() {"", "Adapters- Eaton SWD", "Adapters- M12", "Adapters- MINIFast", "Assembly", "Cable Glands", "Cable- Bulk", "Cables- M12 4W Gry", "Cables- M12 4W Power Cables", "Cables- M12 5W GRN for SWD", "Cables- M12 5W Gry for Devices", "Cables- M12 5W Shielded for Encoders", "Cables- M12 5W Yel for Estop Bus", "Cables- M12 8W GRN for SWD", "Cables- M12 8W Shielded for Encoders", "Cables- M12 Shielded Ethernet", "Cables- M12 Unshielded Ethernet", "Cables- M12 w/ Flying Leads", "Cables- MINIFast", "Camera & Vision Parts", "Circuit Breakers", "Circuit Breakers- Accessories", "Circuit Breakers- Class-2", "Circuit Breakers- Miniature", "Circuit Breakers- Molded Case", "Connectors", "Connectors- M12", "Connectors- Solenoid", "Control Stations", "Control Stations- Boxes", "Control Stations- Pushbuttons & Lights", "Control Stations- Switches", "Drives", "Drives- Basic VFDs", "Drives- Premium VFDs", "Drives- Soft Starts", "E-Stop- Hardware", "E-Stop- Pull Cords", "E-Stop- Pushbuttons", "Enclosures", "Encoders", "Fuse Holders", "Fuses", "Hardware- General", "Hardware- Mounting", "Hole Plugs", "Indicators", "Indicators- Horns", "Indicators- Panel Buttons", "Indicators- Panel Lights", "Indicators- StackLights & Horns", "KIT", "Labels", "Line Reactors", "Motor Starters", "Motor Starters- Accessories", "Motor Starters- Contactors", "Motor Starters- Protectors", "Motor Starters- Reversing", "Motor Starters- Standard", "Mounting- Panel Brackets", "Mounting- Panel Stands", "Mounting- Universal Plates", "Network Switches", "Other", "PLC- COMM CARDs", "PLC- CPUs", "PLC- IO CARDs", "PLC- Power Supplies", "PLC- Racks", "PLCs", "Power Supplies", "Relays", "RIO Junction Boxes- M12", "RIO Module-IP20", "RIO Module-IP20- Eaton SWD", "RIO Module-IP20- Turck", "RIO Module-IP67", "RIO Module-IP67- Eaton SWD", "Sensors", "Sensors- Inductive", "Sensors- Limit Switches", "Sensors- Photo Eyes", "Sensors- Pressure", "Sensors- Proximity", "Sensors- Temperature", "Sensors- Ultrasonic", "Switches", "Switches- Motor Disconnects", "Switches- Panel", "Terminal Blocks", "Terminal Blocks- Din Rail Anchors", "Terminal Blocks- End Barriers", "Terminal Blocks- Jumpers", "Terminal Blocks- Markers", "Terminal Blocks- Push in Type", "Terminal Blocks- Screw Type", "Terminal Blocks- Spring Cage Type", "Transformers", "Wire", "Wire- Accessories"})
        Me.Column2.Name = "Column2"
        Me.Column2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Column2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Column2.Width = 200
        '
        'Column19
        '
        Me.Column19.HeaderText = "Notes"
        Me.Column19.Name = "Column19"
        Me.Column19.Width = 700
        '
        'Column27
        '
        Me.Column27.HeaderText = "Need by Date"
        Me.Column27.Name = "Column27"
        Me.Column27.Width = 400
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 43)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1819, 940)
        Me.TabControl1.TabIndex = 17
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.DimGray
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.Panel_grid)
        Me.TabPage1.Location = New System.Drawing.Point(4, 32)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1811, 904)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "  Panels  "
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button8)
        Me.GroupBox2.Controls.Add(Me.ComboBox1)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Button5)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 17)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(560, 231)
        Me.GroupBox2.TabIndex = 54
        Me.GroupBox2.TabStop = False
        '
        'Button8
        '
        Me.Button8.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button8.AutoSize = True
        Me.Button8.BackColor = System.Drawing.Color.Gray
        Me.Button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button8.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button8.ForeColor = System.Drawing.Color.Black
        Me.Button8.Location = New System.Drawing.Point(305, 163)
        Me.Button8.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(159, 49)
        Me.Button8.TabIndex = 53
        Me.Button8.Text = "Change Panel Qty"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(202, 29)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(323, 31)
        Me.ComboBox1.TabIndex = 48
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label4.Location = New System.Drawing.Point(20, 29)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(132, 25)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "Select a Panel:"
        '
        'Button5
        '
        Me.Button5.AutoSize = True
        Me.Button5.BackColor = System.Drawing.Color.Brown
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Button5.Location = New System.Drawing.Point(16, 161)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(148, 51)
        Me.Button5.TabIndex = 52
        Me.Button5.Text = "Remove Panel"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.options_v)
        Me.GroupBox1.Controls.Add(Me.qty_b)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.p_name3)
        Me.GroupBox1.Controls.Add(Me.p_name2)
        Me.GroupBox1.Controls.Add(Me.p_name1)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.panel_desc)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Location = New System.Drawing.Point(592, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1197, 240)
        Me.GroupBox1.TabIndex = 53
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter New Panel Info"
        '
        'options_v
        '
        Me.options_v.BackColor = System.Drawing.Color.LightGray
        Me.options_v.CheckOnClick = True
        Me.options_v.FormattingEnabled = True
        Me.options_v.Items.AddRange(New Object() {"(0) None", "(1) Brake", "(2) Motor-Fan", "(3) Braking Resistor", "(4) Brake and Motor-Fan", "(5) Motor-Fan and Braking Resistor", "(6) Brake and Braking Resistor", "(7) Brake, Motor-Fan and Braking Resistor"})
        Me.options_v.Location = New System.Drawing.Point(632, 28)
        Me.options_v.Name = "options_v"
        Me.options_v.Size = New System.Drawing.Size(532, 204)
        Me.options_v.TabIndex = 65
        Me.options_v.Visible = False
        '
        'qty_b
        '
        Me.qty_b.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.qty_b.Location = New System.Drawing.Point(180, 94)
        Me.qty_b.Name = "qty_b"
        Me.qty_b.Size = New System.Drawing.Size(100, 30)
        Me.qty_b.TabIndex = 64
        Me.qty_b.Text = "1"
        Me.qty_b.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(95, 94)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 25)
        Me.Label2.TabIndex = 63
        Me.Label2.Text = "Qty:"
        '
        'Button6
        '
        Me.Button6.AutoSize = True
        Me.Button6.BackColor = System.Drawing.Color.DarkGray
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button6.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.ForeColor = System.Drawing.Color.Black
        Me.Button6.Location = New System.Drawing.Point(434, 186)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(151, 46)
        Me.Button6.TabIndex = 60
        Me.Button6.Text = "Remove Row"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'p_name3
        '
        Me.p_name3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name3.FormattingEnabled = True
        Me.p_name3.Items.AddRange(New Object() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"})
        Me.p_name3.Location = New System.Drawing.Point(409, 47)
        Me.p_name3.Name = "p_name3"
        Me.p_name3.Size = New System.Drawing.Size(90, 31)
        Me.p_name3.TabIndex = 59
        '
        'p_name2
        '
        Me.p_name2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name2.FormattingEnabled = True
        Me.p_name2.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99"})
        Me.p_name2.Location = New System.Drawing.Point(295, 47)
        Me.p_name2.Name = "p_name2"
        Me.p_name2.Size = New System.Drawing.Size(90, 31)
        Me.p_name2.TabIndex = 58
        '
        'p_name1
        '
        Me.p_name1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.p_name1.FormattingEnabled = True
        Me.p_name1.Items.AddRange(New Object() {"ADA", "CP", "VFD"})
        Me.p_name1.Location = New System.Drawing.Point(180, 47)
        Me.p_name1.Name = "p_name1"
        Me.p_name1.Size = New System.Drawing.Size(90, 31)
        Me.p_name1.TabIndex = 57
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(37, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(117, 25)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "Panel Name:"
        '
        'panel_desc
        '
        Me.panel_desc.Location = New System.Drawing.Point(180, 139)
        Me.panel_desc.Name = "panel_desc"
        Me.panel_desc.Size = New System.Drawing.Size(405, 30)
        Me.panel_desc.TabIndex = 55
        Me.panel_desc.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button4
        '
        Me.Button4.AutoSize = True
        Me.Button4.BackColor = System.Drawing.Color.CadetBlue
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(52, 186)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(151, 46)
        Me.Button4.TabIndex = 50
        Me.Button4.Text = "Part Selector"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.AutoSize = True
        Me.Button3.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(252, 186)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(133, 46)
        Me.Button3.TabIndex = 51
        Me.Button3.Text = "Add Panel"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(37, 140)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 25)
        Me.Label3.TabIndex = 52
        Me.Label3.Text = "Description:"
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.DimGray
        Me.TabPage2.Controls.Add(Me.Button7)
        Me.TabPage2.Controls.Add(Me.Button1)
        Me.TabPage2.Controls.Add(Me.Button2)
        Me.TabPage2.Controls.Add(Me.assem_grid)
        Me.TabPage2.Location = New System.Drawing.Point(4, 32)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(1811, 904)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "  Assemblies  "
        '
        'Button7
        '
        Me.Button7.AutoSize = True
        Me.Button7.BackColor = System.Drawing.Color.DarkGray
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button7.ForeColor = System.Drawing.Color.Black
        Me.Button7.Location = New System.Drawing.Point(1592, 8)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(135, 61)
        Me.Button7.TabIndex = 61
        Me.Button7.Text = "Remove Row"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.AutoSize = True
        Me.Button1.BackColor = System.Drawing.Color.DarkGoldenrod
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(1260, 8)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(141, 61)
        Me.Button1.TabIndex = 41
        Me.Button1.Text = "Add Assemblies"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.AutoSize = True
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(1428, 8)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(132, 61)
        Me.Button2.TabIndex = 40
        Me.Button2.Text = "Part Selector"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'assem_grid
        '
        Me.assem_grid.AllowUserToDeleteRows = False
        Me.assem_grid.AllowUserToOrderColumns = True
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.assem_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle5
        Me.assem_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.assem_grid.BackgroundColor = System.Drawing.Color.Gray
        Me.assem_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.assem_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.assem_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle6
        Me.assem_grid.ColumnHeadersHeight = 48
        Me.assem_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.assem_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewComboBoxColumn1, Me.DataGridViewComboBoxColumn2, Me.DataGridViewComboBoxColumn3, Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9})
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.assem_grid.DefaultCellStyle = DataGridViewCellStyle7
        Me.assem_grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.assem_grid.EnableHeadersVisualStyles = False
        Me.assem_grid.Location = New System.Drawing.Point(7, 79)
        Me.assem_grid.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.assem_grid.Name = "assem_grid"
        Me.assem_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.assem_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle8
        Me.assem_grid.RowHeadersWidth = 16
        Me.assem_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.assem_grid.RowTemplate.Height = 54
        Me.assem_grid.Size = New System.Drawing.Size(1797, 817)
        Me.assem_grid.TabIndex = 17
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.Frozen = True
        Me.DataGridViewTextBoxColumn1.HeaderText = "Part_No"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 380
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.HeaderText = "Description"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 330
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.HeaderText = "Manufacturer"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 200
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.HeaderText = "Vendor"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.Width = 200
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.HeaderText = "Price"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 150
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.HeaderText = "Qty"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 150
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.HeaderText = "Subtotal"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.Width = 150
        '
        'DataGridViewComboBoxColumn1
        '
        Me.DataGridViewComboBoxColumn1.HeaderText = "MFG Type"
        Me.DataGridViewComboBoxColumn1.Items.AddRange(New Object() {"", "Panel", "Field", "Assembly", "Bulk"})
        Me.DataGridViewComboBoxColumn1.Name = "DataGridViewComboBoxColumn1"
        Me.DataGridViewComboBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn1.Width = 200
        '
        'DataGridViewComboBoxColumn2
        '
        Me.DataGridViewComboBoxColumn2.HeaderText = "(Special Order/Preferred)"
        Me.DataGridViewComboBoxColumn2.Items.AddRange(New Object() {"", "Preferred", "Special Order"})
        Me.DataGridViewComboBoxColumn2.Name = "DataGridViewComboBoxColumn2"
        Me.DataGridViewComboBoxColumn2.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn2.Width = 160
        '
        'DataGridViewComboBoxColumn3
        '
        Me.DataGridViewComboBoxColumn3.HeaderText = "Part Type"
        Me.DataGridViewComboBoxColumn3.Items.AddRange(New Object() {"", "Adapters- Eaton SWD", "Adapters- M12", "Adapters- MINIFast", "Assembly", "Cable Glands", "Cable- Bulk", "Cables- M12 4W Gry", "Cables- M12 4W Power Cables", "Cables- M12 5W GRN for SWD", "Cables- M12 5W Gry for Devices", "Cables- M12 5W Shielded for Encoders", "Cables- M12 5W Yel for Estop Bus", "Cables- M12 8W GRN for SWD", "Cables- M12 8W Shielded for Encoders", "Cables- M12 Shielded Ethernet", "Cables- M12 Unshielded Ethernet", "Cables- M12 w/ Flying Leads", "Cables- MINIFast", "Camera & Vision Parts", "Circuit Breakers", "Circuit Breakers- Accessories", "Circuit Breakers- Class-2", "Circuit Breakers- Miniature", "Circuit Breakers- Molded Case", "Connectors", "Connectors- M12", "Connectors- Solenoid", "Control Stations", "Control Stations- Boxes", "Control Stations- Pushbuttons & Lights", "Control Stations- Switches", "Drives", "Drives- Basic VFDs", "Drives- Premium VFDs", "Drives- Soft Starts", "E-Stop- Hardware", "E-Stop- Pull Cords", "E-Stop- Pushbuttons", "Enclosures", "Encoders", "Fuse Holders", "Fuses", "Hardware- General", "Hardware- Mounting", "Hole Plugs", "Indicators", "Indicators- Horns", "Indicators- Panel Buttons", "Indicators- Panel Lights", "Indicators- StackLights & Horns", "KIT", "Labels", "Line Reactors", "Motor Starters", "Motor Starters- Accessories", "Motor Starters- Contactors", "Motor Starters- Protectors", "Motor Starters- Reversing", "Motor Starters- Standard", "Mounting- Panel Brackets", "Mounting- Panel Stands", "Mounting- Universal Plates", "Network Switches", "Other", "PLC- COMM CARDs", "PLC- CPUs", "PLC- IO CARDs", "PLC- Power Supplies", "PLC- Racks", "PLCs", "Power Supplies", "Relays", "RIO Junction Boxes- M12", "RIO Module-IP20", "RIO Module-IP20- Eaton SWD", "RIO Module-IP20- Turck", "RIO Module-IP67", "RIO Module-IP67- Eaton SWD", "Sensors", "Sensors- Inductive", "Sensors- Limit Switches", "Sensors- Photo Eyes", "Sensors- Pressure", "Sensors- Proximity", "Sensors- Temperature", "Sensors- Ultrasonic", "Switches", "Switches- Motor Disconnects", "Switches- Panel", "Terminal Blocks", "Terminal Blocks- Din Rail Anchors", "Terminal Blocks- End Barriers", "Terminal Blocks- Jumpers", "Terminal Blocks- Markers", "Terminal Blocks- Push in Type", "Terminal Blocks- Screw Type", "Terminal Blocks- Spring Cage Type", "Transformers", "Wire", "Wire- Accessories"})
        Me.DataGridViewComboBoxColumn3.Name = "DataGridViewComboBoxColumn3"
        Me.DataGridViewComboBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridViewComboBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.DataGridViewComboBoxColumn3.Width = 200
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.HeaderText = "Notes"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.Width = 700
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.HeaderText = "Need by Date"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.Width = 400
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1843, 28)
        Me.MenuStrip1.TabIndex = 55
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportBOMToolStripMenuItem, Me.ExportToExcelToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(56, 24)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'ImportBOMToolStripMenuItem
        '
        Me.ImportBOMToolStripMenuItem.Name = "ImportBOMToolStripMenuItem"
        Me.ImportBOMToolStripMenuItem.Size = New System.Drawing.Size(183, 26)
        Me.ImportBOMToolStripMenuItem.Text = "Import BOM"
        '
        'ExportToExcelToolStripMenuItem
        '
        Me.ExportToExcelToolStripMenuItem.Name = "ExportToExcelToolStripMenuItem"
        Me.ExportToExcelToolStripMenuItem.Size = New System.Drawing.Size(183, 26)
        Me.ExportToExcelToolStripMenuItem.Text = "Export to Excel"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Part_Picker1
        '
        Me.Part_Picker1.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.Part_Picker1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Part_Picker1.ForeColor = System.Drawing.Color.Black
        Me.Part_Picker1.Location = New System.Drawing.Point(524, 31)
        Me.Part_Picker1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Part_Picker1.Name = "Part_Picker1"
        Me.Part_Picker1.Size = New System.Drawing.Size(1, 1)
        Me.Part_Picker1.TabIndex = 54
        '
        'Edit_BOM_Panels
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 23.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1843, 995)
        Me.Controls.Add(Me.Part_Picker1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "Edit_BOM_Panels"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BOM Panels and Assemblies"
        CType(Me.Panel_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.assem_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel_grid As DataGridView
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Description As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column13 As DataGridViewTextBoxColumn
    Friend WithEvents Column14 As DataGridViewTextBoxColumn
    Friend WithEvents Column15 As DataGridViewTextBoxColumn
    Friend WithEvents Column16 As DataGridViewTextBoxColumn
    Friend WithEvents Column12 As DataGridViewComboBoxColumn
    Friend WithEvents Column3 As DataGridViewComboBoxColumn
    Friend WithEvents Column2 As DataGridViewComboBoxColumn
    Friend WithEvents Column19 As DataGridViewTextBoxColumn
    Friend WithEvents Column27 As DataGridViewTextBoxColumn
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents assem_grid As DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn1 As DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn2 As DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewComboBoxColumn3 As DataGridViewComboBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As DataGridViewTextBoxColumn
    Friend WithEvents Label4 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents panel_desc As TextBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Part_Picker1 As Part_Picker
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents p_name3 As ComboBox
    Friend WithEvents p_name2 As ComboBox
    Friend WithEvents p_name1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents qty_b As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportBOMToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportToExcelToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Button8 As Button
    Friend WithEvents options_v As CheckedListBox
End Class
