<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class add_KIT
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(add_KIT))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.sel_kit = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.kit_part_grid = New System.Windows.Forms.TextBox()
        Me.kit_grid = New System.Windows.Forms.DataGridView()
        Me.qty_box = New System.Windows.Forms.TextBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.legacy = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.p_vendor = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.akt_box = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.desc = New System.Windows.Forms.TextBox()
        Me.kit_name = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.kit_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.DimGray
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(12, 96)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1526, 720)
        Me.Panel1.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(549, 33)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.BackColor = System.Drawing.Color.DimGray
        Me.SplitContainer1.Panel1.Controls.Add(Me.sel_kit)
        Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.BackColor = System.Drawing.Color.DimGray
        Me.SplitContainer1.Panel2.Controls.Add(Me.kit_grid)
        Me.SplitContainer1.Panel2.Controls.Add(Me.qty_box)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label18)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button5)
        Me.SplitContainer1.Size = New System.Drawing.Size(955, 663)
        Me.SplitContainer1.SplitterDistance = 179
        Me.SplitContainer1.TabIndex = 1
        '
        'sel_kit
        '
        Me.sel_kit.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.sel_kit.AutoSize = True
        Me.sel_kit.Font = New System.Drawing.Font("Segoe UI Semibold", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sel_kit.ForeColor = System.Drawing.Color.White
        Me.sel_kit.Location = New System.Drawing.Point(644, 72)
        Me.sel_kit.Name = "sel_kit"
        Me.sel_kit.Size = New System.Drawing.Size(123, 25)
        Me.sel_kit.TabIndex = 3
        Me.sel_kit.Text = "Selected Kit :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.kit_part_grid)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(14, 14)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(576, 136)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Search Component"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(422, 47)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(135, 38)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Search"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'kit_part_grid
        '
        Me.kit_part_grid.Location = New System.Drawing.Point(17, 53)
        Me.kit_part_grid.Name = "kit_part_grid"
        Me.kit_part_grid.Size = New System.Drawing.Size(386, 30)
        Me.kit_part_grid.TabIndex = 0
        '
        'kit_grid
        '
        Me.kit_grid.AllowUserToAddRows = False
        Me.kit_grid.AllowUserToDeleteRows = False
        Me.kit_grid.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.kit_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.kit_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.kit_grid.BackgroundColor = System.Drawing.Color.Silver
        Me.kit_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.kit_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.kit_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.kit_grid.ColumnHeadersHeight = 48
        Me.kit_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.kit_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.kit_grid.EnableHeadersVisualStyles = False
        Me.kit_grid.Location = New System.Drawing.Point(14, 51)
        Me.kit_grid.MultiSelect = False
        Me.kit_grid.Name = "kit_grid"
        Me.kit_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.kit_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.kit_grid.RowHeadersWidth = 16
        Me.kit_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.kit_grid.RowTemplate.Height = 44
        Me.kit_grid.Size = New System.Drawing.Size(753, 410)
        Me.kit_grid.TabIndex = 10
        '
        'qty_box
        '
        Me.qty_box.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.qty_box.Location = New System.Drawing.Point(856, 147)
        Me.qty_box.Name = "qty_box"
        Me.qty_box.Size = New System.Drawing.Size(52, 22)
        Me.qty_box.TabIndex = 2
        Me.qty_box.Text = "1"
        '
        'Label18
        '
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.Color.White
        Me.Label18.Location = New System.Drawing.Point(803, 150)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(37, 17)
        Me.Label18.TabIndex = 1
        Me.Label18.Text = "QTY"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(27, 15)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(372, 23)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Select a Part/Kit/Device member of the new KIT"
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button5.BackColor = System.Drawing.Color.DarkCyan
        Me.Button5.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.White
        Me.Button5.Location = New System.Drawing.Point(794, 39)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(142, 81)
        Me.Button5.TabIndex = 0
        Me.Button5.Text = "ADD PART TO KIT"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.DarkGray
        Me.GroupBox1.Controls.Add(Me.legacy)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.p_vendor)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.akt_box)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.desc)
        Me.GroupBox1.Controls.Add(Me.kit_name)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(19, 33)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(513, 663)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Enter KIT data"
        '
        'legacy
        '
        Me.legacy.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.legacy.Location = New System.Drawing.Point(148, 61)
        Me.legacy.Name = "legacy"
        Me.legacy.Size = New System.Drawing.Size(334, 30)
        Me.legacy.TabIndex = 30
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(54, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = " ADA #"
        '
        'Label5
        '
        Me.Label5.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 204)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Primary Vendor"
        '
        'p_vendor
        '
        Me.p_vendor.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_vendor.Location = New System.Drawing.Point(148, 201)
        Me.p_vendor.Name = "p_vendor"
        Me.p_vendor.Size = New System.Drawing.Size(334, 30)
        Me.p_vendor.TabIndex = 27
        '
        'Label15
        '
        Me.Label15.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(42, 339)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(147, 23)
        Me.Label15.TabIndex = 26
        Me.Label15.Text = "KIT Number (AKT)"
        '
        'akt_box
        '
        Me.akt_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.akt_box.Location = New System.Drawing.Point(206, 336)
        Me.akt_box.Name = "akt_box"
        Me.akt_box.Size = New System.Drawing.Size(276, 30)
        Me.akt_box.TabIndex = 25
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.Location = New System.Drawing.Point(162, 551)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(177, 51)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Add KIT"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(42, 134)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 23)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Description"
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(42, 273)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(85, 23)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "KIT Name"
        '
        'desc
        '
        Me.desc.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.desc.Location = New System.Drawing.Point(148, 134)
        Me.desc.Name = "desc"
        Me.desc.Size = New System.Drawing.Size(334, 30)
        Me.desc.TabIndex = 3
        '
        'kit_name
        '
        Me.kit_name.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.kit_name.Location = New System.Drawing.Point(148, 273)
        Me.kit_name.Name = "kit_name"
        Me.kit_name.Size = New System.Drawing.Size(334, 30)
        Me.kit_name.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.Location = New System.Drawing.Point(1452, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 39)
        Me.Label2.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(769, 41)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 28)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Add KIT"
        '
        'add_KIT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1524, 828)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "add_KIT"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add KIT"
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.kit_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents kit_name As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents desc As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents Label9 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Button2 As Button
    Friend WithEvents kit_part_grid As TextBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Label15 As Label
    Friend WithEvents akt_box As TextBox
    Friend WithEvents qty_box As TextBox
    Friend WithEvents Label18 As Label
    Friend WithEvents sel_kit As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents p_vendor As TextBox
    Friend WithEvents legacy As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents kit_grid As DataGridView
End Class
