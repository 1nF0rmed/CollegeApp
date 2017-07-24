Imports MaterialSkin
Imports System.IO

Public Class princ_main

    Public cs, ce, meg, ec, eee As String

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            meg = Path.GetFullPath(OpenFileDialog1.FileName)
            Label9.Visible = True
            Label9.Text = meg
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ec = Path.GetFullPath(OpenFileDialog1.FileName)
            Label10.Visible = True
            Label10.Text = ec
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            eee = Path.GetFullPath(OpenFileDialog1.FileName)
            Label11.Visible = True
            Label11.Text = eee
        End If
        Button6.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        princ_browse.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        about.Show()
    End Sub

    Private Sub princ_main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Form1.Show()
        princ_browse.Close()
        about.Close()
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            cs = Path.GetFullPath(OpenFileDialog1.FileName)
            Label7.Visible = True
            Label7.Text = cs
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ce = Path.GetFullPath(OpenFileDialog1.FileName)
            Label8.Visible = True
            Label8.Text = ce
        End If
    End Sub
End Class