<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyVariables
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MyVariables))
        Me.variables_box = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'variables_box
        '
        Me.variables_box.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.variables_box.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.variables_box.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.variables_box.Location = New System.Drawing.Point(28, 51)
        Me.variables_box.Multiline = True
        Me.variables_box.Name = "variables_box"
        Me.variables_box.ReadOnly = True
        Me.variables_box.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.variables_box.Size = New System.Drawing.Size(778, 692)
        Me.variables_box.TabIndex = 0
        '
        'MyVariables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(844, 784)
        Me.Controls.Add(Me.variables_box)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MyVariables"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Variables"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents variables_box As TextBox
End Class
