<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class main
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(main))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(404, 427)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select the Database file "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(407, 588)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Selected: "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(510, 588)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(0, 17)
        Me.Label3.TabIndex = 3
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(430, 482)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(121, 64)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(430, 638)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(121, 43)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "CONTINUE"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(181, 383)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(243, 186)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "INSERT"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(557, 383)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(243, 186)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "DERIVATIVE"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.CollegeManagement.My.Resources.Resources.mvjp_logo_color
        Me.PictureBox1.Location = New System.Drawing.Point(258, 106)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(476, 204)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button5.Location = New System.Drawing.Point(799, 106)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(151, 50)
        Me.Button5.TabIndex = 8
        Me.Button5.Text = "LOGOUT"
        Me.Button5.UseVisualStyleBackColor = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(823, 701)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(192, 20)
        Me.Label16.TabIndex = 63
        Me.Label16.Text = "Computer Science Dept."
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(823, 681)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(116, 20)
        Me.Label15.TabIndex = 62
        Me.Label15.Text = "(366CS15019)"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(823, 661)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(157, 20)
        Me.Label14.TabIndex = 61
        Me.Label14.Text = "Pradyumna Rahul K"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(823, 638)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(115, 20)
        Me.Label13.TabIndex = 60
        Me.Label13.Text = "Developed by:"
        '
        'main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(1020, 744)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Menu"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents Button5 As Button
    Friend WithEvents Label16 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label13 As Label
End Class
