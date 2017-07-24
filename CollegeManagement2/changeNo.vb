Imports MaterialSkin
Imports System.Data.OleDb

Public Class changeNo

    Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
    Dim loca As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\"

    Private Sub changeNo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)
    End Sub

    Private Sub renameDirectory(no As String, newNo As String)
        Try
            My.Computer.FileSystem.RenameDirectory(loca & no, newNo)
        Catch ex As Exception
            MessageBox.Show("Check if the admission no exists")
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub renameFiles(no As String, newNo As String)
        Dim types As New List(Of String) _
            From {"_profile", "_f_profile", "_g_profile", "_m_profile", "_medical", "_physical_med", "_fee1", "_fee2", "_fee3", "_study_cert", "_character", "_migration", "_marks_sheet"}
        Dim exts As New List(Of String) _
            From {".jpg", ".png", ".jpeg", ".pdf"}
        Dim path As String
        Try
            For Each type As String In types
                For Each ext As String In exts
                    path = loca & newNo & "\" & no & type & ext
                    If System.IO.File.Exists(path) Then
                        MessageBox.Show(path)
                        My.Computer.FileSystem.RenameFile(path, newNo & type & ext)
                    End If
                Next
            Next
        Catch ex As Exception
            MessageBox.Show("Check if the admission no exists")
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub updateDB(no As String, newNo As String)
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + main.file
        Dim connection As New OleDbConnection(connectionString)
        Dim query As String = "UPDATE student SET admission_no='" & newNo & "' WHERE admission_no='" & no & "'"

        Dim cmd = New OleDbCommand(query, connection)

        connection.Open()
        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        connection.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrEmpty(TextBox1.Text) Then
            MessageBox.Show("Please Enter OLD Admission No")
        ElseIf String.IsNullOrEmpty(TextBox2.Text) Then
            MessageBox.Show("Please Enter NEW Admission No")
        Else
            MessageBox.Show("Changing No.....")
            renameDirectory(TextBox1.Text, TextBox2.Text)
            renameFiles(TextBox1.Text, TextBox2.Text)
            updateDB(TextBox1.Text, TextBox2.Text)
        End If
    End Sub
End Class