<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Update_vendors
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Update_vendors))
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Part_textbox = New System.Windows.Forms.TextBox()
        Me.vendor_grid = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Price_box = New System.Windows.Forms.TextBox()
        Me.date_box = New System.Windows.Forms.TextBox()
        Me.vendor_box = New System.Windows.Forms.TextBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.part_n = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.vendor_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label5.Location = New System.Drawing.Point(26, 38)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 23)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Enter Part Name"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(648, 31)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(71, 48)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Go"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Part_textbox
        '
        Me.Part_textbox.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Part_textbox.Location = New System.Drawing.Point(185, 38)
        Me.Part_textbox.Name = "Part_textbox"
        Me.Part_textbox.Size = New System.Drawing.Size(431, 30)
        Me.Part_textbox.TabIndex = 9
        '
        'vendor_grid
        '
        Me.vendor_grid.AllowUserToAddRows = False
        Me.vendor_grid.AllowUserToDeleteRows = False
        Me.vendor_grid.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.vendor_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.vendor_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vendor_grid.BackgroundColor = System.Drawing.Color.LightSlateGray
        Me.vendor_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.vendor_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.vendor_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.vendor_grid.ColumnHeadersHeight = 48
        Me.vendor_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.vendor_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.vendor_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.vendor_grid.EnableHeadersVisualStyles = False
        Me.vendor_grid.GridColor = System.Drawing.SystemColors.ControlDarkDark
        Me.vendor_grid.Location = New System.Drawing.Point(19, 158)
        Me.vendor_grid.MultiSelect = False
        Me.vendor_grid.Name = "vendor_grid"
        Me.vendor_grid.ReadOnly = True
        Me.vendor_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.vendor_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.vendor_grid.RowHeadersWidth = 16
        Me.vendor_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.vendor_grid.RowTemplate.Height = 44
        Me.vendor_grid.Size = New System.Drawing.Size(1385, 307)
        Me.vendor_grid.TabIndex = 17
        '
        'Column1
        '
        Me.Column1.HeaderText = "Vendor Name"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 500
        '
        'Column2
        '
        Me.Column2.HeaderText = "Cost"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 300
        '
        'Column3
        '
        Me.Column3.HeaderText = "Purchase Date"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        Me.Column3.Width = 400
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.BackColor = System.Drawing.Color.Silver
        Me.GroupBox3.Controls.Add(Me.Button6)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.Price_box)
        Me.GroupBox3.Controls.Add(Me.date_box)
        Me.GroupBox3.Controls.Add(Me.vendor_box)
        Me.GroupBox3.Controls.Add(Me.Button5)
        Me.GroupBox3.Controls.Add(Me.Button4)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(19, 483)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1385, 304)
        Me.GroupBox3.TabIndex = 16
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Update/Remove Data"
        '
        'Button6
        '
        Me.Button6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button6.Location = New System.Drawing.Point(507, 231)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(220, 46)
        Me.Button6.TabIndex = 18
        Me.Button6.Text = "Enter New Entry"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Label17
        '
        Me.Label17.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(872, 97)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(118, 25)
        Me.Label17.TabIndex = 17
        Me.Label17.Text = "yyyy-mm-dd"
        '
        'Price_box
        '
        Me.Price_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Price_box.Location = New System.Drawing.Point(213, 144)
        Me.Price_box.Name = "Price_box"
        Me.Price_box.Size = New System.Drawing.Size(181, 30)
        Me.Price_box.TabIndex = 9
        '
        'date_box
        '
        Me.date_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.date_box.Location = New System.Drawing.Point(798, 65)
        Me.date_box.Name = "date_box"
        Me.date_box.Size = New System.Drawing.Size(294, 30)
        Me.date_box.TabIndex = 8
        '
        'vendor_box
        '
        Me.vendor_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.vendor_box.Location = New System.Drawing.Point(209, 65)
        Me.vendor_box.Name = "vendor_box"
        Me.vendor_box.Size = New System.Drawing.Size(411, 30)
        Me.vendor_box.TabIndex = 6
        '
        'Button5
        '
        Me.Button5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button5.BackColor = System.Drawing.Color.Maroon
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(1195, 144)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(138, 45)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "Remove"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button4.Location = New System.Drawing.Point(1195, 55)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(138, 40)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Update"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(659, 64)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(120, 23)
        Me.Label15.TabIndex = 3
        Me.Label15.Text = "Purchase Date"
        '
        'Label14
        '
        Me.Label14.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(131, 151)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(58, 23)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "Cost $"
        '
        'Label12
        '
        Me.Label12.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(102, 68)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(65, 23)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Vendor"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Segoe UI Semibold", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(-34, -24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(207, 23)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "Vendors, Prices and Dates"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.Location = New System.Drawing.Point(1359, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 39)
        Me.Label2.TabIndex = 18
        '
        'part_n
        '
        Me.part_n.AutoSize = True
        Me.part_n.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.part_n.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.part_n.Location = New System.Drawing.Point(803, 40)
        Me.part_n.Name = "part_n"
        Me.part_n.Size = New System.Drawing.Size(126, 28)
        Me.part_n.TabIndex = 19
        Me.part_n.Text = "Part selected:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(25, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(871, 25)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "To Edit an existing entry, Please double click the entry below and enter the new " &
    "values in the textboxes"
        '
        'Update_vendors
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1416, 819)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.part_n)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.vendor_grid)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Part_textbox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Update_vendors"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Vendors and Prices"
        CType(Me.vendor_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label5 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Part_textbox As TextBox
    Friend WithEvents vendor_grid As DataGridView
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Button6 As Button
    Friend WithEvents Label17 As Label
    Friend WithEvents Price_box As TextBox
    Friend WithEvents date_box As TextBox
    Friend WithEvents vendor_box As TextBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents part_n As Label
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
End Class
