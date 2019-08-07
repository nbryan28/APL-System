<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Cables_package
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cables_package))
        Me.feet_table = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.feet_qty = New System.Windows.Forms.Label()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.warning_l = New System.Windows.Forms.Label()
        CType(Me.feet_table, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'feet_table
        '
        Me.feet_table.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.LightGray
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.feet_table.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.feet_table.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.feet_table.BackgroundColor = System.Drawing.Color.CadetBlue
        Me.feet_table.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleVertical
        Me.feet_table.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(135, Byte), Integer), CType(CType(127, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.feet_table.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.feet_table.ColumnHeadersHeight = 48
        Me.feet_table.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.feet_table.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column4})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer), CType(CType(52, Byte), Integer))
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.feet_table.DefaultCellStyle = DataGridViewCellStyle3
        Me.feet_table.EnableHeadersVisualStyles = False
        Me.feet_table.Location = New System.Drawing.Point(13, 88)
        Me.feet_table.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.feet_table.MultiSelect = False
        Me.feet_table.Name = "feet_table"
        Me.feet_table.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.feet_table.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.feet_table.RowHeadersWidth = 16
        Me.feet_table.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.feet_table.RowTemplate.Height = 60
        Me.feet_table.Size = New System.Drawing.Size(1264, 556)
        Me.feet_table.TabIndex = 21
        '
        'Column1
        '
        Me.Column1.HeaderText = "Pieces"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 550
        '
        'Column4
        '
        Me.Column4.HeaderText = "Feet"
        Me.Column4.Name = "Column4"
        Me.Column4.Width = 600
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(28, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(159, 28)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Total Cable Feet: "
        '
        'feet_qty
        '
        Me.feet_qty.AutoSize = True
        Me.feet_qty.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.feet_qty.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.feet_qty.Location = New System.Drawing.Point(193, 32)
        Me.feet_qty.Name = "feet_qty"
        Me.feet_qty.Size = New System.Drawing.Size(23, 28)
        Me.feet_qty.TabIndex = 23
        Me.feet_qty.Text = "0"
        '
        'Button14
        '
        Me.Button14.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button14.AutoSize = True
        Me.Button14.BackColor = System.Drawing.Color.DarkGray
        Me.Button14.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button14.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button14.ForeColor = System.Drawing.Color.Black
        Me.Button14.Location = New System.Drawing.Point(1070, 22)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(196, 57)
        Me.Button14.TabIndex = 24
        Me.Button14.Text = "Apply"
        Me.Button14.UseVisualStyleBackColor = False
        '
        'warning_l
        '
        Me.warning_l.AutoSize = True
        Me.warning_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.warning_l.ForeColor = System.Drawing.Color.Red
        Me.warning_l.Location = New System.Drawing.Point(505, 28)
        Me.warning_l.Name = "warning_l"
        Me.warning_l.Size = New System.Drawing.Size(541, 38)
        Me.warning_l.TabIndex = 25
        Me.warning_l.Text = "You are using more cable than requested !"
        Me.warning_l.Visible = False
        '
        'Cables_package
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1290, 658)
        Me.Controls.Add(Me.warning_l)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.feet_qty)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.feet_table)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(1308, 705)
        Me.Name = "Cables_package"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cables Packing Instructions"
        CType(Me.feet_table, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents feet_table As DataGridView
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column4 As DataGridViewTextBoxColumn
    Friend WithEvents Label1 As Label
    Friend WithEvents feet_qty As Label
    Friend WithEvents Button14 As Button
    Friend WithEvents warning_l As Label
End Class
