<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Purchase_Request
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Purchase_Request))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SavePurcharseRequestToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportToCSVToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PR_grid = New System.Windows.Forms.DataGridView()
        Me.menu_BOM = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.FindAlternativePartToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem30 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem36 = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ADA_number = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PR_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.menu_BOM.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1924, 33)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SavePurcharseRequestToolStripMenuItem, Me.ExportToCSVToolStripMenuItem, Me.ToolStripSeparator1})
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        Me.MenuToolStripMenuItem.Size = New System.Drawing.Size(73, 29)
        Me.MenuToolStripMenuItem.Text = "Menu"
        '
        'SavePurcharseRequestToolStripMenuItem
        '
        Me.SavePurcharseRequestToolStripMenuItem.Name = "SavePurcharseRequestToolStripMenuItem"
        Me.SavePurcharseRequestToolStripMenuItem.Size = New System.Drawing.Size(214, 30)
        Me.SavePurcharseRequestToolStripMenuItem.Text = "Save BOM"
        '
        'ExportToCSVToolStripMenuItem
        '
        Me.ExportToCSVToolStripMenuItem.Name = "ExportToCSVToolStripMenuItem"
        Me.ExportToCSVToolStripMenuItem.Size = New System.Drawing.Size(214, 30)
        Me.ExportToCSVToolStripMenuItem.Text = "Export to Excel"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(211, 6)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.Gainsboro
        Me.Label1.Location = New System.Drawing.Point(22, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 25)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "# Of Parts: "
        '
        'PR_grid
        '
        Me.PR_grid.AllowUserToAddRows = False
        Me.PR_grid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.PR_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.PR_grid.BackgroundColor = System.Drawing.Color.Gray
        Me.PR_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.PR_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.PR_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.PR_grid.ColumnHeadersHeight = 48
        Me.PR_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.PR_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column10, Me.Description, Me.ADA_number, Me.Column1, Me.Column13, Me.Column14, Me.Column15, Me.Column16, Me.Column12})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.PR_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.PR_grid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PR_grid.EnableHeadersVisualStyles = False
        Me.PR_grid.Location = New System.Drawing.Point(3, 3)
        Me.PR_grid.MultiSelect = False
        Me.PR_grid.Name = "PR_grid"
        Me.PR_grid.ReadOnly = True
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
        Me.PR_grid.Size = New System.Drawing.Size(1856, 867)
        Me.PR_grid.TabIndex = 14
        '
        'menu_BOM
        '
        Me.menu_BOM.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.menu_BOM.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.menu_BOM.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FindAlternativePartToolStripMenuItem, Me.ToolStripMenuItem30, Me.ToolStripMenuItem36})
        Me.menu_BOM.Name = "menu_BOM"
        Me.menu_BOM.Size = New System.Drawing.Size(273, 100)
        '
        'FindAlternativePartToolStripMenuItem
        '
        Me.FindAlternativePartToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.FindAlternativePartToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FindAlternativePartToolStripMenuItem.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.FindAlternativePartToolStripMenuItem.Image = CType(resources.GetObject("FindAlternativePartToolStripMenuItem.Image"), System.Drawing.Image)
        Me.FindAlternativePartToolStripMenuItem.Name = "FindAlternativePartToolStripMenuItem"
        Me.FindAlternativePartToolStripMenuItem.Size = New System.Drawing.Size(272, 32)
        Me.FindAlternativePartToolStripMenuItem.Text = "Find Alternative Part"
        '
        'ToolStripMenuItem30
        '
        Me.ToolStripMenuItem30.BackColor = System.Drawing.Color.CadetBlue
        Me.ToolStripMenuItem30.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripMenuItem30.Name = "ToolStripMenuItem30"
        Me.ToolStripMenuItem30.Size = New System.Drawing.Size(272, 32)
        Me.ToolStripMenuItem30.Text = "Remove Part from PR"
        '
        'ToolStripMenuItem36
        '
        Me.ToolStripMenuItem36.BackColor = System.Drawing.Color.Maroon
        Me.ToolStripMenuItem36.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripMenuItem36.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.ToolStripMenuItem36.Name = "ToolStripMenuItem36"
        Me.ToolStripMenuItem36.Size = New System.Drawing.Size(272, 32)
        Me.ToolStripMenuItem36.Text = "Find in inventory"
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(27, 120)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1870, 911)
        Me.TabControl1.TabIndex = 17
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.DimGray
        Me.TabPage1.Controls.Add(Me.PR_grid)
        Me.TabPage1.Location = New System.Drawing.Point(4, 34)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1862, 873)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "  Bill Of Material   "
        '
        'Column10
        '
        Me.Column10.HeaderText = "Part_No"
        Me.Column10.Name = "Column10"
        Me.Column10.ReadOnly = True
        Me.Column10.Width = 380
        '
        'Description
        '
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        Me.Description.ReadOnly = True
        Me.Description.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Description.Width = 430
        '
        'ADA_number
        '
        Me.ADA_number.HeaderText = "ADA Number"
        Me.ADA_number.Name = "ADA_number"
        Me.ADA_number.ReadOnly = True
        Me.ADA_number.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.ADA_number.Width = 300
        '
        'Column1
        '
        Me.Column1.HeaderText = "Manufacturer"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 200
        '
        'Column13
        '
        Me.Column13.HeaderText = "Vendor"
        Me.Column13.Name = "Column13"
        Me.Column13.ReadOnly = True
        Me.Column13.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column13.Width = 200
        '
        'Column14
        '
        Me.Column14.HeaderText = "Price"
        Me.Column14.Name = "Column14"
        Me.Column14.ReadOnly = True
        Me.Column14.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column14.Width = 150
        '
        'Column15
        '
        Me.Column15.HeaderText = "Qty"
        Me.Column15.Name = "Column15"
        Me.Column15.ReadOnly = True
        Me.Column15.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column15.Width = 150
        '
        'Column16
        '
        Me.Column16.HeaderText = "Subtotal"
        Me.Column16.Name = "Column16"
        Me.Column16.ReadOnly = True
        Me.Column16.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column16.Width = 150
        '
        'Column12
        '
        Me.Column12.HeaderText = "Type"
        Me.Column12.Name = "Column12"
        Me.Column12.ReadOnly = True
        Me.Column12.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        Me.Column12.Width = 200
        '
        'Purchase_Request
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(10.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DarkSlateGray
        Me.ClientSize = New System.Drawing.Size(1924, 1055)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Purchase_Request"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Material Request"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PR_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.menu_BOM.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SavePurcharseRequestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExportToCSVToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents Label1 As Label
    Friend WithEvents PR_grid As DataGridView
    Friend WithEvents menu_BOM As ContextMenuStrip
    Friend WithEvents FindAlternativePartToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem30 As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem36 As ToolStripMenuItem
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Description As DataGridViewTextBoxColumn
    Friend WithEvents ADA_number As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column13 As DataGridViewTextBoxColumn
    Friend WithEvents Column14 As DataGridViewTextBoxColumn
    Friend WithEvents Column15 As DataGridViewTextBoxColumn
    Friend WithEvents Column16 As DataGridViewTextBoxColumn
    Friend WithEvents Column12 As DataGridViewTextBoxColumn
End Class
