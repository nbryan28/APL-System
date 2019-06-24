<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Parts_needed
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Parts_needed))
        Me.p_label = New System.Windows.Forms.Label()
        Me.c_label = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Parts_need_grid = New System.Windows.Forms.DataGridView()
        Me.ADA_number = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column14 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.m_label = New System.Windows.Forms.Label()
        Me.l_label = New System.Windows.Forms.Label()
        Me.type_box = New System.Windows.Forms.ComboBox()
        CType(Me.Parts_need_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'p_label
        '
        Me.p_label.AutoSize = True
        Me.p_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.p_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.p_label.Location = New System.Drawing.Point(1091, 19)
        Me.p_label.Name = "p_label"
        Me.p_label.Size = New System.Drawing.Size(179, 28)
        Me.p_label.TabIndex = 2
        Me.p_label.Text = "Total Parts Needed:"
        Me.p_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'c_label
        '
        Me.c_label.AutoSize = True
        Me.c_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.c_label.ForeColor = System.Drawing.Color.DarkSeaGreen
        Me.c_label.Location = New System.Drawing.Point(1091, 55)
        Me.c_label.Name = "c_label"
        Me.c_label.Size = New System.Drawing.Size(102, 28)
        Me.c_label.TabIndex = 3
        Me.c_label.Text = "Total Cost:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label3.Location = New System.Drawing.Point(25, 50)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Select Type:"
        '
        'Parts_need_grid
        '
        Me.Parts_need_grid.AllowUserToAddRows = False
        Me.Parts_need_grid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Parts_need_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.Parts_need_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Parts_need_grid.BackgroundColor = System.Drawing.Color.DimGray
        Me.Parts_need_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.Parts_need_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Parts_need_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.Parts_need_grid.ColumnHeadersHeight = 48
        Me.Parts_need_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.Parts_need_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ADA_number, Me.Column10, Me.Column15, Me.Column1, Me.Column14, Me.Column16})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Parts_need_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.Parts_need_grid.EnableHeadersVisualStyles = False
        Me.Parts_need_grid.Location = New System.Drawing.Point(12, 106)
        Me.Parts_need_grid.MultiSelect = False
        Me.Parts_need_grid.Name = "Parts_need_grid"
        Me.Parts_need_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Parts_need_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.Parts_need_grid.RowHeadersWidth = 16
        Me.Parts_need_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.Parts_need_grid.RowTemplate.Height = 54
        Me.Parts_need_grid.Size = New System.Drawing.Size(1412, 596)
        Me.Parts_need_grid.TabIndex = 14
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
        'm_label
        '
        Me.m_label.AutoSize = True
        Me.m_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.m_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.m_label.Location = New System.Drawing.Point(736, 19)
        Me.m_label.Name = "m_label"
        Me.m_label.Size = New System.Drawing.Size(132, 28)
        Me.m_label.TabIndex = 15
        Me.m_label.Text = "Material Cost:"
        '
        'l_label
        '
        Me.l_label.AutoSize = True
        Me.l_label.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.l_label.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.l_label.Location = New System.Drawing.Point(736, 55)
        Me.l_label.Name = "l_label"
        Me.l_label.Size = New System.Drawing.Size(110, 28)
        Me.l_label.TabIndex = 16
        Me.l_label.Text = "Labor Cost:"
        '
        'type_box
        '
        Me.type_box.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.type_box.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.type_box.FormattingEnabled = True
        Me.type_box.Items.AddRange(New Object() {"Starter Panel", "IO Panel", "PLC Panel", "Scanners", "Field parts"})
        Me.type_box.Location = New System.Drawing.Point(165, 45)
        Me.type_box.Name = "type_box"
        Me.type_box.Size = New System.Drawing.Size(281, 33)
        Me.type_box.TabIndex = 17
        '
        'Parts_needed
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1449, 730)
        Me.Controls.Add(Me.type_box)
        Me.Controls.Add(Me.l_label)
        Me.Controls.Add(Me.m_label)
        Me.Controls.Add(Me.Parts_need_grid)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.c_label)
        Me.Controls.Add(Me.p_label)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "Parts_needed"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Parts needed"
        CType(Me.Parts_need_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents p_label As Label
    Friend WithEvents c_label As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Parts_need_grid As DataGridView
    Friend WithEvents ADA_number As DataGridViewTextBoxColumn
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Column15 As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column14 As DataGridViewTextBoxColumn
    Friend WithEvents Column16 As DataGridViewTextBoxColumn
    Friend WithEvents m_label As Label
    Friend WithEvents l_label As Label
    Friend WithEvents type_box As ComboBox
End Class
