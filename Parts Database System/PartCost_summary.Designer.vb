<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PartCost_summary
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PartCost_summary))
        Me.Parts_cost_t_grid = New System.Windows.Forms.DataGridView()
        Me.ADA_number = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.type_box_c = New System.Windows.Forms.ComboBox()
        Me.ct_label = New System.Windows.Forms.Label()
        Me.pt_label = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.big_cost = New System.Windows.Forms.Label()
        Me.big_labor = New System.Windows.Forms.Label()
        Me.big_materials = New System.Windows.Forms.Label()
        Me.mt_label = New System.Windows.Forms.Label()
        Me.lt_label = New System.Windows.Forms.Label()
        CType(Me.Parts_cost_t_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Parts_cost_t_grid
        '
        Me.Parts_cost_t_grid.AllowUserToAddRows = False
        Me.Parts_cost_t_grid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Parts_cost_t_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Parts_cost_t_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Parts_cost_t_grid.BackgroundColor = System.Drawing.Color.Gray
        Me.Parts_cost_t_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.Parts_cost_t_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.DarkSlateGray
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Parts_cost_t_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.Parts_cost_t_grid.ColumnHeadersHeight = 48
        Me.Parts_cost_t_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Parts_cost_t_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ADA_number, Me.Column10, Me.Column15, Me.Column1, Me.Column14, Me.Column16})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Parts_cost_t_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.Parts_cost_t_grid.EnableHeadersVisualStyles = False
        Me.Parts_cost_t_grid.GridColor = System.Drawing.Color.DimGray
        Me.Parts_cost_t_grid.Location = New System.Drawing.Point(22, 122)
        Me.Parts_cost_t_grid.MultiSelect = False
        Me.Parts_cost_t_grid.Name = "Parts_cost_t_grid"
        Me.Parts_cost_t_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Parts_cost_t_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.Parts_cost_t_grid.RowHeadersWidth = 16
        Me.Parts_cost_t_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Parts_cost_t_grid.RowTemplate.Height = 54
        Me.Parts_cost_t_grid.Size = New System.Drawing.Size(1742, 619)
        Me.Parts_cost_t_grid.TabIndex = 15
        '
        'ADA_number
        '
        Me.ADA_number.HeaderText = "ADA Number"
        Me.ADA_number.Name = "ADA_number"
        Me.ADA_number.Width = 670
        '
        'Column10
        '
        Me.Column10.HeaderText = "Part No"
        Me.Column10.Name = "Column10"
        Me.Column10.Width = 400
        '
        'Column15
        '
        Me.Column15.HeaderText = "Qty"
        Me.Column15.Name = "Column15"
        '
        'Column1
        '
        Me.Column1.HeaderText = "Vendor"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 300
        '
        'Column14
        '
        Me.Column14.HeaderText = "Price"
        Me.Column14.Name = "Column14"
        '
        'Column16
        '
        Me.Column16.HeaderText = "Subtotal"
        Me.Column16.Name = "Column16"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label3.Location = New System.Drawing.Point(40, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 23)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Select Type:"
        '
        'type_box_c
        '
        Me.type_box_c.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.type_box_c.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.type_box_c.FormattingEnabled = True
        Me.type_box_c.Items.AddRange(New Object() {"Starter Panel", "IO Panel", "PLC Panel", "Scanners", "Field parts", "M12 Cables", "M12 ES Cables"})
        Me.type_box_c.Location = New System.Drawing.Point(156, 53)
        Me.type_box_c.Name = "type_box_c"
        Me.type_box_c.Size = New System.Drawing.Size(281, 33)
        Me.type_box_c.TabIndex = 16
        '
        'ct_label
        '
        Me.ct_label.AutoSize = True
        Me.ct_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ct_label.ForeColor = System.Drawing.Color.DarkSeaGreen
        Me.ct_label.Location = New System.Drawing.Point(1353, 74)
        Me.ct_label.Name = "ct_label"
        Me.ct_label.Size = New System.Drawing.Size(194, 28)
        Me.ct_label.TabIndex = 19
        Me.ct_label.Text = "Materials and Cost: $"
        '
        'pt_label
        '
        Me.pt_label.AutoSize = True
        Me.pt_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pt_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.pt_label.Location = New System.Drawing.Point(1353, 33)
        Me.pt_label.Name = "pt_label"
        Me.pt_label.Size = New System.Drawing.Size(132, 28)
        Me.pt_label.TabIndex = 18
        Me.pt_label.Text = "Parts Needed:"
        Me.pt_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.big_cost)
        Me.GroupBox1.Controls.Add(Me.big_labor)
        Me.GroupBox1.Controls.Add(Me.big_materials)
        Me.GroupBox1.Location = New System.Drawing.Point(22, 757)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1742, 100)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        '
        'big_cost
        '
        Me.big_cost.AutoSize = True
        Me.big_cost.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.big_cost.ForeColor = System.Drawing.Color.DarkSeaGreen
        Me.big_cost.Location = New System.Drawing.Point(1180, 38)
        Me.big_cost.Name = "big_cost"
        Me.big_cost.Size = New System.Drawing.Size(145, 32)
        Me.big_cost.TabIndex = 22
        Me.big_cost.Text = "Total Cost: $"
        Me.big_cost.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'big_labor
        '
        Me.big_labor.AutoSize = True
        Me.big_labor.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.big_labor.ForeColor = System.Drawing.SystemColors.Info
        Me.big_labor.Location = New System.Drawing.Point(599, 46)
        Me.big_labor.Name = "big_labor"
        Me.big_labor.Size = New System.Drawing.Size(151, 23)
        Me.big_labor.TabIndex = 21
        Me.big_labor.Text = "Total Labor Cost: $"
        Me.big_labor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'big_materials
        '
        Me.big_materials.AutoSize = True
        Me.big_materials.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.big_materials.ForeColor = System.Drawing.SystemColors.Info
        Me.big_materials.Location = New System.Drawing.Point(74, 44)
        Me.big_materials.Name = "big_materials"
        Me.big_materials.Size = New System.Drawing.Size(177, 23)
        Me.big_materials.TabIndex = 20
        Me.big_materials.Text = "Total Materials Cost: $"
        Me.big_materials.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'mt_label
        '
        Me.mt_label.AutoSize = True
        Me.mt_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mt_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.mt_label.Location = New System.Drawing.Point(936, 33)
        Me.mt_label.Name = "mt_label"
        Me.mt_label.Size = New System.Drawing.Size(156, 28)
        Me.mt_label.TabIndex = 21
        Me.mt_label.Text = "Materials Cost: $"
        '
        'lt_label
        '
        Me.lt_label.AutoSize = True
        Me.lt_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lt_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lt_label.Location = New System.Drawing.Point(936, 74)
        Me.lt_label.Name = "lt_label"
        Me.lt_label.Size = New System.Drawing.Size(126, 28)
        Me.lt_label.TabIndex = 22
        Me.lt_label.Text = "Labor Cost: $"
        '
        'PartCost_summary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1804, 878)
        Me.Controls.Add(Me.lt_label)
        Me.Controls.Add(Me.mt_label)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ct_label)
        Me.Controls.Add(Me.pt_label)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.type_box_c)
        Me.Controls.Add(Me.Parts_cost_t_grid)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "PartCost_summary"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Part Cost"
        CType(Me.Parts_cost_t_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Parts_cost_t_grid As DataGridView
    Friend WithEvents ADA_number As DataGridViewTextBoxColumn
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Column15 As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column14 As DataGridViewTextBoxColumn
    Friend WithEvents Column16 As DataGridViewTextBoxColumn
    Friend WithEvents Label3 As Label
    Friend WithEvents type_box_c As ComboBox
    Friend WithEvents ct_label As Label
    Friend WithEvents pt_label As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents big_cost As Label
    Friend WithEvents big_labor As Label
    Friend WithEvents big_materials As Label
    Friend WithEvents mt_label As Label
    Friend WithEvents lt_label As Label
End Class
