<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class princ_browse
    'Inherits System.Windows.Forms.Form
    Inherits MaterialSkin.Controls.MaterialForm

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(princ_browse))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Label13 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(14, 254)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(1519, 459)
        Me.DataGridView1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Chartreuse
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.RoyalBlue
        Me.Button1.Location = New System.Drawing.Point(1308, 736)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(199, 112)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "EXPORT DATA"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(420, 67)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(611, 181)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 27
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(11, 736)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 18)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "Select type:"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"Register", "First Name", "Semester", "Department", "Quota"})
        Me.ComboBox1.Location = New System.Drawing.Point(113, 732)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(195, 24)
        Me.ComboBox1.TabIndex = 29
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Crimson
        Me.Button2.ForeColor = System.Drawing.Color.Yellow
        Me.Button2.Location = New System.Drawing.Point(511, 718)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 50)
        Me.Button2.TabIndex = 30
        Me.Button2.Text = "FILTER"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(315, 732)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(192, 22)
        Me.TextBox1.TabIndex = 31
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Highlight
        Me.Button3.ForeColor = System.Drawing.Color.Chartreuse
        Me.Button3.Location = New System.Drawing.Point(1308, 854)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(147, 50)
        Me.Button3.TabIndex = 32
        Me.Button3.Text = "RESET"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.Coral
        Me.Button4.ForeColor = System.Drawing.Color.Crimson
        Me.Button4.Location = New System.Drawing.Point(14, 810)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(153, 63)
        Me.Button4.TabIndex = 33
        Me.Button4.Text = "PROFILE"
        Me.Button4.UseVisualStyleBackColor = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Coral
        Me.Button5.ForeColor = System.Drawing.Color.Crimson
        Me.Button5.Location = New System.Drawing.Point(14, 894)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(153, 63)
        Me.Button5.TabIndex = 34
        Me.Button5.Text = "PHYSICAL CERT"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Coral
        Me.Button6.ForeColor = System.Drawing.Color.Crimson
        Me.Button6.Location = New System.Drawing.Point(14, 977)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(153, 63)
        Me.Button6.TabIndex = 35
        Me.Button6.Text = "MEDICAL CERT"
        Me.Button6.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.Coral
        Me.Button7.ForeColor = System.Drawing.Color.Crimson
        Me.Button7.Location = New System.Drawing.Point(337, 810)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(153, 63)
        Me.Button7.TabIndex = 36
        Me.Button7.Text = "FEES: YEAR 1"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.Coral
        Me.Button8.ForeColor = System.Drawing.Color.Crimson
        Me.Button8.Location = New System.Drawing.Point(337, 894)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(153, 63)
        Me.Button8.TabIndex = 37
        Me.Button8.Text = "FEES: YEAR 2"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'Button9
        '
        Me.Button9.BackColor = System.Drawing.Color.Coral
        Me.Button9.ForeColor = System.Drawing.Color.Crimson
        Me.Button9.Location = New System.Drawing.Point(337, 977)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(153, 63)
        Me.Button9.TabIndex = 38
        Me.Button9.Text = "FEES: YEAR 3"
        Me.Button9.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.Coral
        Me.Button10.ForeColor = System.Drawing.Color.Crimson
        Me.Button10.Location = New System.Drawing.Point(653, 810)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(153, 63)
        Me.Button10.TabIndex = 39
        Me.Button10.Text = "STUDY CERT"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'Button11
        '
        Me.Button11.BackColor = System.Drawing.Color.Coral
        Me.Button11.ForeColor = System.Drawing.Color.Crimson
        Me.Button11.Location = New System.Drawing.Point(653, 894)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(153, 63)
        Me.Button11.TabIndex = 40
        Me.Button11.Text = "CHARACTER CERT"
        Me.Button11.UseVisualStyleBackColor = False
        '
        'Button12
        '
        Me.Button12.BackColor = System.Drawing.Color.Coral
        Me.Button12.ForeColor = System.Drawing.Color.Crimson
        Me.Button12.Location = New System.Drawing.Point(653, 977)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(153, 63)
        Me.Button12.TabIndex = 41
        Me.Button12.Text = "MIGRATION CERT"
        Me.Button12.UseVisualStyleBackColor = False
        '
        'Button13
        '
        Me.Button13.BackColor = System.Drawing.Color.Coral
        Me.Button13.ForeColor = System.Drawing.Color.Crimson
        Me.Button13.Location = New System.Drawing.Point(951, 977)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(153, 63)
        Me.Button13.TabIndex = 42
        Me.Button13.Text = "10TH MARKS SHEET"
        Me.Button13.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(183, 833)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(81, 17)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "AVAILABLE"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(183, 917)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 17)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "AVAILABLE"
        Me.Label3.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(183, 1000)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 17)
        Me.Label4.TabIndex = 45
        Me.Label4.Text = "AVAILABLE"
        Me.Label4.Visible = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(508, 833)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(81, 17)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "AVAILABLE"
        Me.Label5.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(508, 917)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(81, 17)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "AVAILABLE"
        Me.Label6.Visible = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(508, 1000)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(81, 17)
        Me.Label7.TabIndex = 48
        Me.Label7.Text = "AVAILABLE"
        Me.Label7.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(823, 833)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(81, 17)
        Me.Label8.TabIndex = 49
        Me.Label8.Text = "AVAILABLE"
        Me.Label8.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(823, 917)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(81, 17)
        Me.Label9.TabIndex = 50
        Me.Label9.Text = "AVAILABLE"
        Me.Label9.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(823, 1000)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(81, 17)
        Me.Label10.TabIndex = 51
        Me.Label10.Text = "AVAILABLE"
        Me.Label10.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(1119, 1000)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(81, 17)
        Me.Label11.TabIndex = 52
        Me.Label11.Text = "AVAILABLE"
        Me.Label11.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(650, 736)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(98, 17)
        Me.Label12.TabIndex = 53
        Me.Label12.Text = "Admission No:"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(754, 733)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(192, 22)
        Me.TextBox2.TabIndex = 54
        '
        'Button14
        '
        Me.Button14.BackColor = System.Drawing.Color.Crimson
        Me.Button14.ForeColor = System.Drawing.Color.Yellow
        Me.Button14.Location = New System.Drawing.Point(963, 728)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(96, 36)
        Me.Button14.TabIndex = 55
        Me.Button14.Text = "CHECK"
        Me.Button14.UseVisualStyleBackColor = False
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(315, 761)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(19, 20)
        Me.Label13.TabIndex = 56
        Me.Label13.Text = "?"
        '
        'princ_browse
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(1545, 1084)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "princ_browse"
        Me.Text = "BROWSE"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents Button1 As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Button2 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents Button9 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents Button11 As Button
    Friend WithEvents Button12 As Button
    Friend WithEvents Button13 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Button14 As Button
    Friend WithEvents Label13 As Label
End Class
