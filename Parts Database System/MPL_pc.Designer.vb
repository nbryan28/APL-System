<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MPL_pc
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MPL_pc))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.RemoveShippingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveRowToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PR_grid = New System.Windows.Forms.DataGridView()
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ADA_number = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.p_number = New System.Windows.Forms.Label()
        Me.p_name = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.z_number = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.phone = New System.Windows.Forms.TextBox()
        Me.address_ship = New System.Windows.Forms.TextBox()
        Me.attn = New System.Windows.Forms.TextBox()
        Me.phone_l = New System.Windows.Forms.Label()
        Me.address_l = New System.Windows.Forms.Label()
        Me.attn_l = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.p_number2 = New System.Windows.Forms.Label()
        Me.p_name2 = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.PR_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1799, 33)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator1, Me.RemoveShippingToolStripMenuItem, Me.RemoveRowToolStripMenuItem})
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        Me.MenuToolStripMenuItem.Size = New System.Drawing.Size(73, 29)
        Me.MenuToolStripMenuItem.Text = "Menu"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(259, 6)
        '
        'RemoveShippingToolStripMenuItem
        '
        Me.RemoveShippingToolStripMenuItem.Name = "RemoveShippingToolStripMenuItem"
        Me.RemoveShippingToolStripMenuItem.Size = New System.Drawing.Size(262, 30)
        Me.RemoveShippingToolStripMenuItem.Text = "Export to Excel"
        '
        'RemoveRowToolStripMenuItem
        '
        Me.RemoveRowToolStripMenuItem.Name = "RemoveRowToolStripMenuItem"
        Me.RemoveRowToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.R), System.Windows.Forms.Keys)
        Me.RemoveRowToolStripMenuItem.Size = New System.Drawing.Size(262, 30)
        Me.RemoveRowToolStripMenuItem.Text = "Remove Row"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.Panel1.Controls.Add(Me.PR_grid)
        Me.Panel1.Location = New System.Drawing.Point(0, 288)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1799, 592)
        Me.Panel1.TabIndex = 2
        '
        'PR_grid
        '
        Me.PR_grid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.PR_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.PR_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PR_grid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.PR_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.DimGray
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PR_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.PR_grid.ColumnHeadersHeight = 58
        Me.PR_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.PR_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column13, Me.ADA_number, Me.Description, Me.Column1})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.DimGray
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.PR_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.PR_grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.PR_grid.EnableHeadersVisualStyles = False
        Me.PR_grid.Location = New System.Drawing.Point(13, 0)
        Me.PR_grid.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PR_grid.Name = "PR_grid"
        Me.PR_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PR_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.PR_grid.RowHeadersWidth = 16
        Me.PR_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.PR_grid.RowTemplate.Height = 54
        Me.PR_grid.Size = New System.Drawing.Size(1761, 578)
        Me.PR_grid.TabIndex = 17
        '
        'Column13
        '
        Me.Column13.HeaderText = "  Qty  "
        Me.Column13.Name = "Column13"
        Me.Column13.Width = 200
        '
        'ADA_number
        '
        Me.ADA_number.HeaderText = " Part No "
        Me.ADA_number.Name = "ADA_number"
        Me.ADA_number.Width = 700
        '
        'Description
        '
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        Me.Description.Width = 930
        '
        'Column1
        '
        Me.Column1.HeaderText = "Notes"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 200
        '
        'p_number
        '
        Me.p_number.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_number.AutoSize = True
        Me.p_number.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p_number.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.p_number.Location = New System.Drawing.Point(19, 67)
        Me.p_number.Name = "p_number"
        Me.p_number.Size = New System.Drawing.Size(91, 25)
        Me.p_number.TabIndex = 4
        Me.p_number.Text = "Project #:"
        '
        'p_name
        '
        Me.p_name.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_name.AutoSize = True
        Me.p_name.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p_name.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.p_name.Location = New System.Drawing.Point(19, 114)
        Me.p_name.Name = "p_name"
        Me.p_name.Size = New System.Drawing.Size(77, 28)
        Me.p_name.TabIndex = 21
        Me.p_name.Text = "Project:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox1.BackColor = System.Drawing.Color.DimGray
        Me.GroupBox1.Controls.Add(Me.z_number)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.phone)
        Me.GroupBox1.Controls.Add(Me.address_ship)
        Me.GroupBox1.Controls.Add(Me.attn)
        Me.GroupBox1.Controls.Add(Me.phone_l)
        Me.GroupBox1.Controls.Add(Me.address_l)
        Me.GroupBox1.Controls.Add(Me.attn_l)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Location = New System.Drawing.Point(775, 55)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1012, 190)
        Me.GroupBox1.TabIndex = 23
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Ship to Address"
        '
        'z_number
        '
        Me.z_number.Location = New System.Drawing.Point(198, 141)
        Me.z_number.Name = "z_number"
        Me.z_number.Size = New System.Drawing.Size(657, 31)
        Me.z_number.TabIndex = 16
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(123, 141)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 28)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "1Z:"
        '
        'phone
        '
        Me.phone.Location = New System.Drawing.Point(525, 37)
        Me.phone.Name = "phone"
        Me.phone.Size = New System.Drawing.Size(403, 31)
        Me.phone.TabIndex = 13
        '
        'address_ship
        '
        Me.address_ship.Location = New System.Drawing.Point(198, 87)
        Me.address_ship.Name = "address_ship"
        Me.address_ship.Size = New System.Drawing.Size(788, 31)
        Me.address_ship.TabIndex = 12
        '
        'attn
        '
        Me.attn.Location = New System.Drawing.Point(117, 37)
        Me.attn.Name = "attn"
        Me.attn.Size = New System.Drawing.Size(284, 31)
        Me.attn.TabIndex = 11
        '
        'phone_l
        '
        Me.phone_l.AutoSize = True
        Me.phone_l.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phone_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.phone_l.Location = New System.Drawing.Point(438, 37)
        Me.phone_l.Name = "phone_l"
        Me.phone_l.Size = New System.Drawing.Size(71, 28)
        Me.phone_l.TabIndex = 10
        Me.phone_l.Text = "Phone:"
        '
        'address_l
        '
        Me.address_l.AutoSize = True
        Me.address_l.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.address_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.address_l.Location = New System.Drawing.Point(32, 87)
        Me.address_l.Name = "address_l"
        Me.address_l.Size = New System.Drawing.Size(142, 28)
        Me.address_l.TabIndex = 8
        Me.address_l.Text = "Street Address:"
        '
        'attn_l
        '
        Me.attn_l.AutoSize = True
        Me.attn_l.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.attn_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.attn_l.Location = New System.Drawing.Point(32, 37)
        Me.attn_l.Name = "attn_l"
        Me.attn_l.Size = New System.Drawing.Size(54, 28)
        Me.attn_l.TabIndex = 6
        Me.attn_l.Text = "Attn:"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.DateTimePicker1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Location = New System.Drawing.Point(122, 177)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(299, 30)
        Me.DateTimePicker1.TabIndex = 24
        '
        'Label4
        '
        Me.Label4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label4.Location = New System.Drawing.Point(19, 178)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(57, 28)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Date:"
        '
        'p_number2
        '
        Me.p_number2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_number2.AutoSize = True
        Me.p_number2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p_number2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.p_number2.Location = New System.Drawing.Point(131, 67)
        Me.p_number2.Name = "p_number2"
        Me.p_number2.Size = New System.Drawing.Size(95, 28)
        Me.p_number2.TabIndex = 26
        Me.p_number2.Text = "Unknown"
        '
        'p_name2
        '
        Me.p_name2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.p_name2.AutoSize = True
        Me.p_name2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p_name2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.p_name2.Location = New System.Drawing.Point(131, 114)
        Me.p_name2.Name = "p_name2"
        Me.p_name2.Size = New System.Drawing.Size(45, 28)
        Me.p_name2.TabIndex = 27
        Me.p_name2.Text = "000"
        '
        'MPL_pc
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1799, 880)
        Me.Controls.Add(Me.p_name2)
        Me.Controls.Add(Me.p_number2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.p_name)
        Me.Controls.Add(Me.p_number)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "MPL_pc"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Packing List"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        CType(Me.PR_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents Panel1 As Panel
    Friend WithEvents p_number As Label
    Friend WithEvents PR_grid As DataGridView
    Friend WithEvents RemoveShippingToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents p_name As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents phone_l As Label
    Friend WithEvents address_l As Label
    Friend WithEvents attn_l As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents phone As TextBox
    Friend WithEvents address_ship As TextBox
    Friend WithEvents attn As TextBox
    Friend WithEvents z_number As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents RemoveRowToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label4 As Label
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents p_number2 As Label
    Friend WithEvents p_name2 As Label
    Friend WithEvents Column13 As DataGridViewTextBoxColumn
    Friend WithEvents ADA_number As DataGridViewTextBoxColumn
    Friend WithEvents Description As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
End Class
