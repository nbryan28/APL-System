<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Extra_info
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Extra_info))
        Me.Records_found = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Current_l = New System.Windows.Forms.Label()
        Me.min_l = New System.Windows.Forms.Label()
        Me.max_l = New System.Windows.Forms.Label()
        Me.qti_l = New System.Windows.Forms.Label()
        Me.demand_l = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Records_found
        '
        Me.Records_found.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Records_found.AutoSize = True
        Me.Records_found.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Records_found.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Records_found.Location = New System.Drawing.Point(78, 77)
        Me.Records_found.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Records_found.Name = "Records_found"
        Me.Records_found.Size = New System.Drawing.Size(323, 38)
        Me.Records_found.TabIndex = 11
        Me.Records_found.Text = "Current Qty In Inventory:"
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(78, 166)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 38)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Min Qty:"
        '
        'Label3
        '
        Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(77, 483)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(476, 32)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Total Current Demand from Active Projects:"
        '
        'Label4
        '
        Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label4.Location = New System.Drawing.Point(76, 375)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(187, 38)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Qty In Transit:"
        '
        'Label5
        '
        Me.Label5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label5.Location = New System.Drawing.Point(78, 268)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(127, 38)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Max Qty:"
        '
        'Current_l
        '
        Me.Current_l.AutoSize = True
        Me.Current_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Current_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Current_l.Location = New System.Drawing.Point(440, 77)
        Me.Current_l.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Current_l.Name = "Current_l"
        Me.Current_l.Size = New System.Drawing.Size(32, 38)
        Me.Current_l.TabIndex = 17
        Me.Current_l.Text = "0"
        '
        'min_l
        '
        Me.min_l.AutoSize = True
        Me.min_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.min_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.min_l.Location = New System.Drawing.Point(225, 166)
        Me.min_l.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.min_l.Name = "min_l"
        Me.min_l.Size = New System.Drawing.Size(32, 38)
        Me.min_l.TabIndex = 19
        Me.min_l.Text = "0"
        '
        'max_l
        '
        Me.max_l.AutoSize = True
        Me.max_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.max_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.max_l.Location = New System.Drawing.Point(231, 268)
        Me.max_l.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.max_l.Name = "max_l"
        Me.max_l.Size = New System.Drawing.Size(32, 38)
        Me.max_l.TabIndex = 20
        Me.max_l.Text = "0"
        '
        'qti_l
        '
        Me.qti_l.AutoSize = True
        Me.qti_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.qti_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.qti_l.Location = New System.Drawing.Point(288, 375)
        Me.qti_l.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.qti_l.Name = "qti_l"
        Me.qti_l.Size = New System.Drawing.Size(32, 38)
        Me.qti_l.TabIndex = 21
        Me.qti_l.Text = "0"
        '
        'demand_l
        '
        Me.demand_l.AutoSize = True
        Me.demand_l.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.demand_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.demand_l.Location = New System.Drawing.Point(575, 477)
        Me.demand_l.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.demand_l.Name = "demand_l"
        Me.demand_l.Size = New System.Drawing.Size(32, 38)
        Me.demand_l.TabIndex = 22
        Me.demand_l.Text = "0"
        '
        'Extra_info
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 28.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(735, 727)
        Me.Controls.Add(Me.demand_l)
        Me.Controls.Add(Me.qti_l)
        Me.Controls.Add(Me.max_l)
        Me.Controls.Add(Me.min_l)
        Me.Controls.Add(Me.Current_l)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Records_found)
        Me.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximumSize = New System.Drawing.Size(898, 835)
        Me.Name = "Extra_info"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Extra Info"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Records_found As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Current_l As Label
    Friend WithEvents min_l As Label
    Friend WithEvents max_l As Label
    Friend WithEvents qti_l As Label
    Friend WithEvents demand_l As Label
End Class
