<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UR_inventory
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(UR_inventory))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Location_t = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Min_qty = New System.Windows.Forms.TextBox()
        Me.max_qty = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.add_box = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.current_qty = New System.Windows.Forms.TextBox()
        Me.remove_box = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.part_name = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label1.Location = New System.Drawing.Point(74, 451)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 38)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Location"
        '
        'Location_t
        '
        Me.Location_t.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Location_t.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Location_t.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Location_t.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location_t.Location = New System.Drawing.Point(231, 456)
        Me.Location_t.Name = "Location_t"
        Me.Location_t.Size = New System.Drawing.Size(243, 31)
        Me.Location_t.TabIndex = 1
        Me.Location_t.Text = "shelf1"
        Me.Location_t.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.Location = New System.Drawing.Point(79, 150)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(125, 41)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Min Qty"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.Location = New System.Drawing.Point(74, 304)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(130, 41)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Max Qty"
        '
        'Min_qty
        '
        Me.Min_qty.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Min_qty.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Min_qty.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Min_qty.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Min_qty.Location = New System.Drawing.Point(231, 160)
        Me.Min_qty.Name = "Min_qty"
        Me.Min_qty.Size = New System.Drawing.Size(173, 31)
        Me.Min_qty.TabIndex = 4
        Me.Min_qty.Text = "0"
        Me.Min_qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'max_qty
        '
        Me.max_qty.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.max_qty.BackColor = System.Drawing.Color.WhiteSmoke
        Me.max_qty.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.max_qty.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.max_qty.Location = New System.Drawing.Point(231, 312)
        Me.max_qty.Name = "max_qty"
        Me.max_qty.Size = New System.Drawing.Size(173, 31)
        Me.max_qty.TabIndex = 5
        Me.max_qty.Text = "0"
        Me.max_qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox1.Controls.Add(Me.add_box)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Location = New System.Drawing.Point(43, 56)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(917, 191)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'add_box
        '
        Me.add_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.add_box.BackColor = System.Drawing.Color.WhiteSmoke
        Me.add_box.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.add_box.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.add_box.Location = New System.Drawing.Point(303, 77)
        Me.add_box.Name = "add_box"
        Me.add_box.Size = New System.Drawing.Size(214, 40)
        Me.add_box.TabIndex = 15
        Me.add_box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button2
        '
        Me.Button2.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button2.BackColor = System.Drawing.Color.CadetBlue
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Black
        Me.Button2.Location = New System.Drawing.Point(627, 61)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(167, 74)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Add"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label5.Location = New System.Drawing.Point(95, 76)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(129, 41)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Add Qty"
        '
        'current_qty
        '
        Me.current_qty.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.current_qty.BackColor = System.Drawing.Color.WhiteSmoke
        Me.current_qty.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.current_qty.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.current_qty.Location = New System.Drawing.Point(355, 43)
        Me.current_qty.Name = "current_qty"
        Me.current_qty.Size = New System.Drawing.Size(200, 40)
        Me.current_qty.TabIndex = 17
        Me.current_qty.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'remove_box
        '
        Me.remove_box.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.remove_box.BackColor = System.Drawing.Color.WhiteSmoke
        Me.remove_box.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.remove_box.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.remove_box.Location = New System.Drawing.Point(335, 57)
        Me.remove_box.Name = "remove_box"
        Me.remove_box.Size = New System.Drawing.Size(200, 40)
        Me.remove_box.TabIndex = 16
        Me.remove_box.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button3
        '
        Me.Button3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button3.BackColor = System.Drawing.Color.Brown
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(667, 25)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(167, 79)
        Me.Button3.TabIndex = 14
        Me.Button3.Text = "Update"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Button1.BackColor = System.Drawing.Color.Peru
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 16.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(641, 42)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(167, 87)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Remove"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label7.Location = New System.Drawing.Point(62, 42)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(244, 41)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Total Current Qty"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.Label6.Location = New System.Drawing.Point(86, 56)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(181, 41)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Remove Qty"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button4)
        Me.GroupBox2.Controls.Add(Me.part_name)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Location_t)
        Me.GroupBox2.Controls.Add(Me.max_qty)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Min_qty)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox2.Location = New System.Drawing.Point(33, 50)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(538, 742)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Levels and Location"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.DarkGray
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(369, 596)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(140, 69)
        Me.Button4.TabIndex = 13
        Me.Button4.Text = "Done"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'part_name
        '
        Me.part_name.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.part_name.AutoSize = True
        Me.part_name.Font = New System.Drawing.Font("Segoe UI", 13.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.part_name.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.part_name.Location = New System.Drawing.Point(181, 58)
        Me.part_name.Name = "part_name"
        Me.part_name.Size = New System.Drawing.Size(126, 32)
        Me.part_name.TabIndex = 6
        Me.part_name.Text = "Part Name"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.current_qty)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Location = New System.Drawing.Point(43, 580)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(917, 122)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.Button1)
        Me.GroupBox4.Controls.Add(Me.remove_box)
        Me.GroupBox4.Location = New System.Drawing.Point(41, 317)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(919, 151)
        Me.GroupBox4.TabIndex = 19
        Me.GroupBox4.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.GroupBox1)
        Me.GroupBox5.Controls.Add(Me.GroupBox3)
        Me.GroupBox5.Controls.Add(Me.GroupBox4)
        Me.GroupBox5.Font = New System.Drawing.Font("Segoe UI", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox5.Location = New System.Drawing.Point(613, 37)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(995, 755)
        Me.GroupBox5.TabIndex = 20
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Inventory Adjustment"
        '
        'UR_inventory
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(61, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(78, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1620, 867)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "UR_inventory"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Update Inventory"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Location_t As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Min_qty As TextBox
    Friend WithEvents max_qty As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents current_qty As TextBox
    Friend WithEvents remove_box As TextBox
    Friend WithEvents add_box As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Button4 As Button
    Friend WithEvents part_name As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox5 As GroupBox
End Class
