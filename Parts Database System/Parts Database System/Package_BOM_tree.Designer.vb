<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Package_BOM_tree
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Package_BOM_tree))
        Me.MBOM = New System.Windows.Forms.Label()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CreateRevisionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.P_BOM = New System.Windows.Forms.Label()
        Me.F_BOM = New System.Windows.Forms.Label()
        Me.A_BOM = New System.Windows.Forms.Label()
        Me.SP_BOM = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.link_f = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.open_grid = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.open_grid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MBOM
        '
        Me.MBOM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.MBOM.BackColor = System.Drawing.Color.Teal
        Me.MBOM.ContextMenuStrip = Me.ContextMenuStrip1
        Me.MBOM.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MBOM.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.MBOM.Location = New System.Drawing.Point(646, 72)
        Me.MBOM.Name = "MBOM"
        Me.MBOM.Size = New System.Drawing.Size(267, 103)
        Me.MBOM.TabIndex = 0
        Me.MBOM.Text = "Master BOM"
        Me.MBOM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreateRevisionToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(123, 32)
        '
        'CreateRevisionToolStripMenuItem
        '
        Me.CreateRevisionToolStripMenuItem.Name = "CreateRevisionToolStripMenuItem"
        Me.CreateRevisionToolStripMenuItem.Size = New System.Drawing.Size(122, 28)
        Me.CreateRevisionToolStripMenuItem.Text = "Open"
        '
        'P_BOM
        '
        Me.P_BOM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.P_BOM.BackColor = System.Drawing.Color.DarkSlateGray
        Me.P_BOM.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.P_BOM.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.P_BOM.Location = New System.Drawing.Point(161, 281)
        Me.P_BOM.Name = "P_BOM"
        Me.P_BOM.Size = New System.Drawing.Size(267, 103)
        Me.P_BOM.TabIndex = 1
        Me.P_BOM.Text = "Panel BOMs"
        Me.P_BOM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'F_BOM
        '
        Me.F_BOM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.F_BOM.BackColor = System.Drawing.Color.Teal
        Me.F_BOM.ContextMenuStrip = Me.ContextMenuStrip1
        Me.F_BOM.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.F_BOM.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.F_BOM.Location = New System.Drawing.Point(482, 281)
        Me.F_BOM.Name = "F_BOM"
        Me.F_BOM.Size = New System.Drawing.Size(267, 103)
        Me.F_BOM.TabIndex = 2
        Me.F_BOM.Text = "Field BOM"
        Me.F_BOM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'A_BOM
        '
        Me.A_BOM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.A_BOM.BackColor = System.Drawing.Color.DarkSlateGray
        Me.A_BOM.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A_BOM.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.A_BOM.Location = New System.Drawing.Point(817, 281)
        Me.A_BOM.Name = "A_BOM"
        Me.A_BOM.Size = New System.Drawing.Size(267, 103)
        Me.A_BOM.TabIndex = 3
        Me.A_BOM.Text = "Assembly BOM"
        Me.A_BOM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'SP_BOM
        '
        Me.SP_BOM.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.SP_BOM.BackColor = System.Drawing.Color.Teal
        Me.SP_BOM.ContextMenuStrip = Me.ContextMenuStrip1
        Me.SP_BOM.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SP_BOM.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.SP_BOM.Location = New System.Drawing.Point(1144, 281)
        Me.SP_BOM.Name = "SP_BOM"
        Me.SP_BOM.Size = New System.Drawing.Size(267, 103)
        Me.SP_BOM.TabIndex = 4
        Me.SP_BOM.Text = "Spare Parts BOM"
        Me.SP_BOM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label6.BackColor = System.Drawing.Color.DimGray
        Me.Label6.Location = New System.Drawing.Point(772, 175)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(10, 74)
        Me.Label6.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label7.BackColor = System.Drawing.Color.DimGray
        Me.Label7.Location = New System.Drawing.Point(286, 249)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(988, 10)
        Me.Label7.TabIndex = 6
        '
        'link_f
        '
        Me.link_f.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.link_f.BackColor = System.Drawing.Color.DimGray
        Me.link_f.Location = New System.Drawing.Point(286, 259)
        Me.link_f.Name = "link_f"
        Me.link_f.Size = New System.Drawing.Size(10, 22)
        Me.link_f.TabIndex = 7
        '
        'Label9
        '
        Me.Label9.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label9.BackColor = System.Drawing.Color.DimGray
        Me.Label9.Location = New System.Drawing.Point(614, 259)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(10, 22)
        Me.Label9.TabIndex = 8
        '
        'Label10
        '
        Me.Label10.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label10.BackColor = System.Drawing.Color.DimGray
        Me.Label10.Location = New System.Drawing.Point(941, 259)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(10, 22)
        Me.Label10.TabIndex = 9
        '
        'Label11
        '
        Me.Label11.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label11.BackColor = System.Drawing.Color.DimGray
        Me.Label11.Location = New System.Drawing.Point(1264, 259)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(10, 22)
        Me.Label11.TabIndex = 10
        '
        'ToolTip1
        '
        Me.ToolTip1.BackColor = System.Drawing.Color.CadetBlue
        Me.ToolTip1.ForeColor = System.Drawing.Color.WhiteSmoke
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.open_grid)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(10, 841)
        Me.Panel1.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.Maroon
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(-37, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 23)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "X"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 10.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(-312, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(150, 25)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Selected Object:"
        '
        'open_grid
        '
        Me.open_grid.AllowUserToAddRows = False
        Me.open_grid.AllowUserToDeleteRows = False
        Me.open_grid.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.open_grid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.open_grid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.open_grid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.open_grid.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.open_grid.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal
        Me.open_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.open_grid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.open_grid.ColumnHeadersHeight = 68
        Me.open_grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.open_grid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column8})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.WhiteSmoke
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.Teal
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.open_grid.DefaultCellStyle = DataGridViewCellStyle3
        Me.open_grid.EnableHeadersVisualStyles = False
        Me.open_grid.Location = New System.Drawing.Point(12, 61)
        Me.open_grid.MultiSelect = False
        Me.open_grid.Name = "open_grid"
        Me.open_grid.ReadOnly = True
        Me.open_grid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.ControlDarkDark
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.open_grid.RowHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.open_grid.RowHeadersVisible = False
        Me.open_grid.RowHeadersWidth = 16
        Me.open_grid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.open_grid.RowTemplate.Height = 64
        Me.open_grid.Size = New System.Drawing.Size(0, 768)
        Me.open_grid.TabIndex = 15
        '
        'Column1
        '
        Me.Column1.HeaderText = "BOM Name"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 410
        '
        'Column8
        '
        Me.Column8.HeaderText = "Job"
        Me.Column8.Name = "Column8"
        Me.Column8.ReadOnly = True
        Me.Column8.Width = 140
        '
        'Package_BOM_tree
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(88, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1596, 841)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.link_f)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.SP_BOM)
        Me.Controls.Add(Me.A_BOM)
        Me.Controls.Add(Me.F_BOM)
        Me.Controls.Add(Me.P_BOM)
        Me.Controls.Add(Me.MBOM)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Package_BOM_tree"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "BOM Manager"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.open_grid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents MBOM As Label
    Friend WithEvents P_BOM As Label
    Friend WithEvents F_BOM As Label
    Friend WithEvents A_BOM As Label
    Friend WithEvents SP_BOM As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents link_f As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents CreateRevisionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Panel1 As Panel
    Friend WithEvents open_grid As DataGridView
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column8 As DataGridViewTextBoxColumn
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
End Class
