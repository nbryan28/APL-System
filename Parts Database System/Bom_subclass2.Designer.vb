<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Bom_packing
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Bom_packing))
        Me.job_label = New System.Windows.Forms.Label()
        Me.bom_label = New System.Windows.Forms.Label()
        Me.packing_grid = New System.Windows.Forms.DataGridView()
        Me.Column15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Description = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.packing_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'job_label
        '
        Me.job_label.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.job_label.AutoSize = True
        Me.job_label.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.job_label.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.job_label.Location = New System.Drawing.Point(1286, 20)
        Me.job_label.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.job_label.Name = "job_label"
        Me.job_label.Size = New System.Drawing.Size(131, 25)
        Me.job_label.TabIndex = 41
        Me.job_label.Text = "Open Project: "
        '
        'bom_label
        '
        Me.bom_label.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.bom_label.AutoSize = True
        Me.bom_label.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bom_label.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.bom_label.Location = New System.Drawing.Point(27, 20)
        Me.bom_label.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.bom_label.Name = "bom_label"
        Me.bom_label.Size = New System.Drawing.Size(195, 25)
        Me.bom_label.TabIndex = 40
        Me.bom_label.Text = "Packing List for BOM: "
        '
        'packing_grid
        '
        Me.packing_grid.AllowUserToAddRows = False
        Me.packing_grid.AllowUserToDeleteRows = False
        Me.packing_grid.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.packing_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.packing_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.packing_grid.BackgroundColor = System.Drawing.Color.Gray
        Me.packing_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.packing_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.packing_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.packing_grid.ColumnHeadersHeight = 48
        Me.packing_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.packing_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column15, Me.Column10, Me.Description, Me.Column1})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.packing_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.packing_grid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.packing_grid.EnableHeadersVisualStyles = False
        Me.packing_grid.Location = New System.Drawing.Point(13, 80)
        Me.packing_grid.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.packing_grid.Name = "packing_grid"
        Me.packing_grid.ReadOnly = True
        Me.packing_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.packing_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.packing_grid.RowHeadersWidth = 16
        Me.packing_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.packing_grid.RowTemplate.Height = 54
        Me.packing_grid.Size = New System.Drawing.Size(1674, 682)
        Me.packing_grid.TabIndex = 39
        '
        'Column15
        '
        Me.Column15.HeaderText = "Qty"
        Me.Column15.Name = "Column15"
        Me.Column15.ReadOnly = True
        Me.Column15.Width = 180
        '
        'Column10
        '
        Me.Column10.HeaderText = "Part_No"
        Me.Column10.Name = "Column10"
        Me.Column10.ReadOnly = True
        Me.Column10.Width = 480
        '
        'Description
        '
        Me.Description.HeaderText = "Description"
        Me.Description.Name = "Description"
        Me.Description.ReadOnly = True
        Me.Description.Width = 500
        '
        'Column1
        '
        Me.Column1.HeaderText = "Need by Date"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 500
        '
        'Bom_packing
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1700, 783)
        Me.Controls.Add(Me.job_label)
        Me.Controls.Add(Me.bom_label)
        Me.Controls.Add(Me.packing_grid)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Bom_packing"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Project Packing List"
        CType(Me.packing_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents job_label As Label
    Friend WithEvents bom_label As Label
    Friend WithEvents packing_grid As DataGridView
    Friend WithEvents Column15 As DataGridViewTextBoxColumn
    Friend WithEvents Column10 As DataGridViewTextBoxColumn
    Friend WithEvents Description As DataGridViewTextBoxColumn
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
End Class
