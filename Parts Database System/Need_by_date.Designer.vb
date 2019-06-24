<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Need_by_date
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Need_by_date))
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.name_l = New System.Windows.Forms.Label()
        Me.need_date_l = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.date_n = New System.Windows.Forms.Label()
        Me.panel_n = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Location = New System.Drawing.Point(38, 201)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 0
        '
        'name_l
        '
        Me.name_l.AutoSize = True
        Me.name_l.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.name_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.name_l.Location = New System.Drawing.Point(33, 37)
        Me.name_l.Name = "name_l"
        Me.name_l.Size = New System.Drawing.Size(213, 25)
        Me.name_l.TabIndex = 1
        Me.name_l.Text = "Panel/ Assembly Name: "
        '
        'need_date_l
        '
        Me.need_date_l.AutoSize = True
        Me.need_date_l.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.need_date_l.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.need_date_l.Location = New System.Drawing.Point(33, 97)
        Me.need_date_l.Name = "need_date_l"
        Me.need_date_l.Size = New System.Drawing.Size(204, 25)
        Me.need_date_l.TabIndex = 2
        Me.need_date_l.Text = "Current Need by Date: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(33, 157)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(344, 25)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Please select a new need by date below"
        '
        'Button2
        '
        Me.Button2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(454, 278)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(154, 50)
        Me.Button2.TabIndex = 34
        Me.Button2.Text = "Apply Changes"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'date_n
        '
        Me.date_n.AutoSize = True
        Me.date_n.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.date_n.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.date_n.Location = New System.Drawing.Point(259, 97)
        Me.date_n.Name = "date_n"
        Me.date_n.Size = New System.Drawing.Size(86, 25)
        Me.date_n.TabIndex = 35
        Me.date_n.Text = "1/1/2099"
        '
        'panel_n
        '
        Me.panel_n.AutoSize = True
        Me.panel_n.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panel_n.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.panel_n.Location = New System.Drawing.Point(259, 37)
        Me.panel_n.Name = "panel_n"
        Me.panel_n.Size = New System.Drawing.Size(58, 25)
        Me.panel_n.TabIndex = 36
        Me.panel_n.Text = "None"
        '
        'Need_by_date
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(768, 439)
        Me.Controls.Add(Me.panel_n)
        Me.Controls.Add(Me.date_n)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.need_date_l)
        Me.Controls.Add(Me.name_l)
        Me.Controls.Add(Me.MonthCalendar1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximumSize = New System.Drawing.Size(786, 486)
        Me.Name = "Need_by_date"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Need by date"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents MonthCalendar1 As MonthCalendar
    Friend WithEvents name_l As Label
    Friend WithEvents need_date_l As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents date_n As Label
    Friend WithEvents panel_n As Label
End Class
