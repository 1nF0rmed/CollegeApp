Imports MaterialSkin
Imports System.Data.OleDb

Public Class Form1

    Dim conString As String
    Dim myCon As OleDbConnection = New OleDbConnection

    Public naame As String
    Public dept As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)

        TextBox2.PasswordChar = "x"

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        MessageBox.Show("Enter your Username")
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        ResetPassword.Show()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        MessageBox.Show("Enter your Password")
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =|%USERPROFILE%|\Documents\College.accdb".Replace("|%USERPROFILE%|", src)
        myCon.ConnectionString = conString
        myCon.Open()

        ' Query
        Dim query As OleDbCommand = New OleDbCommand("SELECT * FROM login WHERE user = '" & TextBox1.Text & "' AND pass = '" & TextBox2.Text & "'", myCon)
        Dim dr As OleDbDataReader = query.ExecuteReader

        Dim userFound As Boolean = False

        While dr.Read
            userFound = True
        End While

        If userFound = True Then
            MessageBox.Show("Login Success")
            If String.Compare(TextBox1.Text, "principal") = 0 Then
                naame = TextBox1.Text
                princ_main.Show()
            Else
                naame = TextBox1.Text
                dept = naame.Substring(4)
                'MessageBox.Show(dept)
                main.Show()
            End If
            TextBox1.Clear()
            TextBox2.Clear()
            Me.Hide()
        Else
            MessageBox.Show("Invalid Username or Password")
        End If
        myCon.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) 
        about.Show()
    End Sub
End Class
