Imports MaterialSkin
Imports System.IO

Public Class main

    Public file As String
    Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
    Private Sub main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileDialog1.Filter = "Access Database File (*.accdb)|*.accdb"
        Try
            My.Computer.FileSystem.CopyFile("style.css", "|%USERPROFILE%|\Pictures\style.css".Replace("|%USERPROFILE%|", src))
        Catch ex As Exception
            Console.WriteLine("Copying style.css file")
        End Try
        Try
            My.Computer.FileSystem.CopyFile("sheets-of-paper.css", "|%USERPROFILE%|\Pictures\sheets-of-paper.css".Replace("|%USERPROFILE%|", src))
        Catch ex As Exception
            Console.WriteLine("Copying sheets-of-paper.css file")
        End Try
        Button2.Hide()
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)

        If Not System.IO.Directory.Exists("|%USERPROFILE%|\Documents\".Replace("|%USERPROFILE%|", src) & Form1.naame) Then
            System.IO.Directory.CreateDirectory("|%USERPROFILE%|\Documents\".Replace("|%USERPROFILE%|", src) & Form1.naame)
            System.IO.Directory.CreateDirectory("|%USERPROFILE%|\Documents\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img")
        End If

        Dim path As String = "|%USERPROFILE%|\Documents\".Replace("|%USERPROFILE%|", src)
        ' MessageBox.Show(path)
        If System.IO.File.Exists(path & "app.config") Then
            Hide_browse()
            file = getPath(path & "app.config")
            Button3.Show()
            Button4.Show()
        End If
    End Sub

    Private Function getPath(conf As String)
        Dim file_data As String = My.Computer.FileSystem.ReadAllText(conf)

        Dim data() As String
        data = file_data.Split(":")

        ' MessageBox.Show(data(1) & ":" & data(2))
        Return data(1) & ":" & data(2)
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            file = Path.GetFullPath(OpenFileDialog1.FileName)
            Label3.Text = file
            Button2.Show()
        End If

        Dim conf_path As String = "|%USERPROFILE%|\Documents\app.config".Replace("|%USERPROFILE%|", src)

        My.Computer.FileSystem.WriteAllText(conf_path, "path:" & file, True)

    End Sub

    Private Sub Hide_browse()
        Label1.Hide()
        Label2.Hide()
        Label3.Hide()

        Button1.Hide()
        Button2.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Hide_browse()

        Button3.Show()
        Button4.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        insert.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        browse.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form1.Show()
        changeNo.Close()
        insert.Close()
        browse.Close()
        about.Close()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) 
        about.Show()
    End Sub
End Class