Imports MaterialSkin
Imports System.Data.OleDb

Public Class ResetPassword

    Dim conString As String
    Dim myCon As OleDbConnection = New OleDbConnection

    Private Sub ResetPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)

        TextBox2.PasswordChar = "x"
        TextBox1.PasswordChar = "x"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If String.Compare(TextBox1.Text, TextBox2.Text) = 0 Then
            conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =C:\Users\1nf0rmed\Documents\College.accdb"
            myCon.ConnectionString = conString
            myCon.Open()

            ' Get the user data
            Dim user As String = TextBox3.Text


            ' Query
            Dim query As OleDbCommand = New OleDbCommand("UPDATE login SET pass='" & TextBox2.Text & "' WHERE user='" & TextBox3.Text & "'", myCon)
            query.ExecuteNonQuery()

            MessageBox.Show("Password Change Successfull")

            myCon.Close()
        Else

            MessageBox.Show("Passwords don't match!")

        End If


    End Sub
End Class