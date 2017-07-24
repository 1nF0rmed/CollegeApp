Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports MaterialSkin
Imports System.IO

Public Class browse
    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + main.file
    Dim connection As New OleDbConnection(connectionString)
    Dim query As String = "SELECT * FROM student"
    Dim dataAdp As New OleDbDataAdapter(query, connection)
    Dim ds As New DataSet()

    Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")

    Dim file, ext As String

    Private Sub browse_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)

        connection.Open()
        dataAdp.Fill(ds, "student")
        connection.Close()

        Dim builder As OleDbCommandBuilder = New OleDbCommandBuilder(dataAdp)

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "student"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Validate()
        Me.dataAdp.Update(Me.ds.Tables("student"))
        Me.ds.AcceptChanges()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label13.Visible = False
        Label14.Visible = False


        Dim base As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img"
        'MessageBox.Show(base)
        Dim path As String = base & "\" & TextBox1.Text & "\" & TextBox1.Text
        If System.IO.File.Exists(path & "_profile.jpg") Then
            Label1.Visible = True
            Label1.Text = "Available"
            'MessageBox.Show("Yeah mate")
        Else
            Label1.Visible = True
            Label1.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_physical_med.jpg") Or System.IO.File.Exists(path & "_physical_med.png") Or System.IO.File.Exists(path & "_physical_med.pdf") Then
            Label2.Visible = True
            Label2.Text = "Available"
            'MessageBox.Show("Yeah mate")
        Else
            Label2.Visible = True
            Label2.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_medical.jpg") Or System.IO.File.Exists(path & "_medical.png") Or System.IO.File.Exists(path & "_medical.pdf") Then
            Label3.Visible = True
            Label3.Text = "Available"
        Else
            Label3.Visible = True
            Label3.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_fee1.jpg") Or System.IO.File.Exists(path & "_fee1.png") Or System.IO.File.Exists(path & "_fee1.pdf") Then
            Label6.Visible = True
            Label6.Text = "Available"
        Else
            Label6.Visible = True
            Label6.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_fee2.jpg") Or System.IO.File.Exists(path & "_fee2.png") Or System.IO.File.Exists(path & "_fee2.pdf") Then
            Label10.Visible = True
            Label10.Text = "Available"
        Else
            Label10.Visible = True
            Label10.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_fee3.jpg") Or System.IO.File.Exists(path & "_fee3.png") Or System.IO.File.Exists(path & "_fee3.pdf") Then
            Label11.Visible = True
            Label11.Text = "Available"
        Else
            Label11.Visible = True
            Label11.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_study_cert.jpg") Or System.IO.File.Exists(path & "_study_cert.png") Or System.IO.File.Exists(path & "_study_cert.pdf") Then
            Label5.Visible = True
            Label5.Text = "Available"
        Else
            Label5.Visible = True
            Label5.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_character.jpg") Or System.IO.File.Exists(path & "_character.png") Or System.IO.File.Exists(path & "_character.pdf") Then
            Label4.Visible = True
            Label4.Text = "Available"
        Else
            Label4.Visible = True
            Label4.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_migration.jpg") Or System.IO.File.Exists(path & "_migration.png") Or System.IO.File.Exists(path & "_migration.pdf") Then
            Label7.Visible = True
            Label7.Text = "Available"
        Else
            Label7.Visible = True
            Label7.Text = "Not Available"
        End If
        If System.IO.File.Exists(path & "_marks_sheet.jpg") Or System.IO.File.Exists(path & "_marks_sheet.png") Or System.IO.File.Exists(path & "_marks_sheet.pdf") Then
            Label9.Visible = True
            Label9.Text = "Available"
        Else
            Label9.Visible = True
            Label9.Text = "Not Available"
        End If

        If System.IO.File.Exists(path & "_m_profile.jpg") Or System.IO.File.Exists(path & "_m_profile.png") Or System.IO.File.Exists(path & "_m_profile.pdf") Then
            Label14.Visible = True
            Label14.Text = "Available"
        Else
            Label14.Visible = True
            Label14.Text = "Not Available"
        End If

        If System.IO.File.Exists(path & "_f_profile.jpg") Or System.IO.File.Exists(path & "_f_profile.png") Or System.IO.File.Exists(path & "_f_profile.pdf") Then
            Label13.Visible = True
            Label13.Text = "Available"
        Else
            Label13.Visible = True
            Label13.Text = "Not Available"
        End If

    End Sub

    Private Sub save_files(file1 As String, type As String, ext As String)
        Dim path As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\" & TextBox1.Text
        Dim dest As String = path & "\" & TextBox1.Text & "_" & type & ext

        'MessageBox.Show(dest)

        Try
            FileCopy(file1, dest)
        Catch ex As Exception
            MessageBox.Show("No such folder")
        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "profile", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "physical_med", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "medical", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "fee1", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "study_cert", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "character", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "migration", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "marks_sheet", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub getFile()
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                file = .FileName
            End If
        End With
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim base As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\"
        Dim admin, path As String

        Dim i As Int16, j As Int16

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        ' Add the Column Names
        For k As Integer = 0 To DataGridView1.Columns.Count - 1
            'MessageBox.Show("Time lah")
            xlWorkSheet.Cells(1, k + 1) = DataGridView1.Columns(k).HeaderText
        Next

        ' Adding the picture column
        xlWorkSheet.Cells(1, (DataGridView1.Columns.Count + 1)) = "Profile"

        ' Adding the data for Columns
        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
            Next
        Next

        ' Adding pictures
        For k As Integer = 0 To DataGridView1.Rows.Count - 2
            admin = DataGridView1(1, k).Value
            path = base & admin & "\" & admin & "_profile.jpg"

            'MessageBox.Show(path)

            If System.IO.File.Exists(path) Then
                With xlApp.ActiveSheet.Pictures.Insert(path)
                    With .ShapeRange
                        .LockAspectRatio = False
                        .Left = 53.8 * (DataGridView1.ColumnCount - 6)
                        .Top = 19 * (k + 1)
                        .Width = 45
                        .Height = 10
                    End With
                    .Placement = 1
                    .PrintObject = True
                End With
            End If
        Next

        xlWorkSheet.Range("A:BE").ColumnWidth = 15
        xlWorkSheet.Range("BF:BG").ColumnWidth = 19
        xlWorkSheet.Rows.RowHeight = 100

        Dim loca As String = ""

        SaveFileDialog1.Filter = "Excel(*.xls)| *xls"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            loca = SaveFileDialog1.FileName
        End If

        Try
            xlWorkBook.SaveAs(loca, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            MessageBox.Show("Exported file to location: " & loca)
            xlApp.Quit()
        Catch ex As Exception
            MessageBox.Show("Not saving the file!")
            xlWorkBook.Close(False, misValue, misValue)
            xlApp.Quit()
        End Try
        'xlWorkBook.Close(True, misValue, misValue)
        'xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.Clear()
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Label11.Visible = False
        Label13.Visible = False
        Label14.Visible = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'Dim query1 As String = "SELECT first_name, last_name, reg_no, dob, year_of, valid_opto, blood_grp, address_line1, address_line2, address_line3, mobile_no, current_sem FROM student"
        Dim query1 As String = "SELECT admission_no, first_name, last_name, reg_no, dob, year_of, valid_opto, blood_grp, address_line1, address_line2, address_line3, mobile_no, current_sem FROM student"
        Dim dataAdp1 As New OleDbDataAdapter(query1, connection)
        Dim ds1 As New DataSet()

        Try
            connection.Open()
            dataAdp1.Fill(ds1, "student")
            connection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim builder1 As OleDbCommandBuilder = New OleDbCommandBuilder(dataAdp)

        DataGridView1.DataSource = ds1
        DataGridView1.DataMember = "student"
        DataGridView1.Refresh()

        Button11.Visible = True
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim query1 As String = "SELECT * FROM student"
        Dim dataAdp1 As New OleDbDataAdapter(query1, connection)
        Dim ds1 As New DataSet()

        Try
            connection.Open()
            dataAdp1.Fill(ds1, "student")
            connection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        Dim builder1 As OleDbCommandBuilder = New OleDbCommandBuilder(dataAdp)

        DataGridView1.DataSource = ds1
        DataGridView1.DataMember = "student"
        DataGridView1.Refresh()

        Button11.Visible = False
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim base As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\"
        Dim admin, path As String

        Dim i As Int16, j As Int16

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        ' Add the Column Names
        For k As Integer = 1 To DataGridView1.Columns.Count
            'MessageBox.Show("Time lah")
            xlWorkSheet.Cells(1, k) = DataGridView1.Columns(k - 1).HeaderText
        Next

        ' Adding the data for Columns
        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
            Next
        Next

        ' Adding the picture column
        xlWorkSheet.Cells(1, DataGridView1.Columns.Count + 2) = "Profile"

        ' Adding pictures
        For k As Integer = 0 To DataGridView1.Rows.Count - 1
            admin = DataGridView1(0, k).Value
            path = base & admin & "\" & admin & "_profile.jpg"

            'MessageBox.Show(path)

            If System.IO.File.Exists(path) Then
                With xlApp.ActiveSheet.Pictures.Insert(path)
                    With .ShapeRange
                        .LockAspectRatio = False
                        .Left = 50 * (DataGridView1.ColumnCount)
                        .Top = 19 * (k + 1)
                        .Width = 45
                        .Height = 10
                    End With
                    .Placement = 1
                    .PrintObject = True
                End With
            End If
        Next

        xlWorkSheet.Rows.RowHeight = 100

        Dim loca As String = ""

        SaveFileDialog1.Filter = "Excel(*.xls)| *xls"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            loca = SaveFileDialog1.FileName
        End If

        xlWorkSheet.Cells(1, 1) = "Department"

        For i = 2 To DataGridView1.RowCount
            xlWorkSheet.Cells(i, 1) = Form1.dept
        Next

        xlWorkSheet.Range("A:M").ColumnWidth = 12
        xlWorkSheet.Range("N:P").ColumnWidth = 19

        Try
            xlWorkBook.SaveAs(loca, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
        Catch ex As Exception
            MessageBox.Show("Not saving the file!")
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
        End Try
        'xlWorkBook.Close(True, misValue, misValue)
        'xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)


        MessageBox.Show("Exported file to location: " & loca)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "fee2", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "fee3", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Search(term As String, type As String)
        Dim query As String = "SELECT * From student WHERE " & type & " like '%" & term & "%'"

        Dim connection As New OleDbConnection(connectionString)
        Dim dataAdp As New OleDbDataAdapter(query, connection)
        Dim ds1 As New DataSet

        connection.Open()
        dataAdp.Fill(ds1, "student")
        connection.Close()

        DataGridView1.DataSource = ds1
        DataGridView1.DataMember = "student"
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If ComboBox1.SelectedIndex = 0 Then
            Search(TextBox2.Text, "reg_no")
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Search(TextBox2.Text, "first_name")
        ElseIf ComboBox1.SelectedIndex = 2 Then
            Search(TextBox2.Text, "current_sem")
        Else
            MessageBox.Show("Please Select Type of Search Term!")
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Dim query As String = "SELECT * From student"

        Dim connection As New OleDbConnection(connectionString)
        Dim dataAdp As New OleDbDataAdapter(query, connection)

        Dim ds2 As New DataSet

        connection.Open()
        dataAdp.Fill(ds2, "student")
        connection.Close()

        DataGridView1.ClearSelection()
        DataGridView1.DataSource = ds2
        'DataGridView1.DataMember = "student"
    End Sub

    Private Sub viewFile(type As String)
        Dim base As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img"
        'MessageBox.Show(base)
        Dim path As String = base & "\" & TextBox1.Text & "\" & TextBox1.Text
        Dim exts As New List(Of String) _
            From {".png", ".jpg", ".pdf"}
        For Each ext As String In exts
            If System.IO.File.Exists(path & type & ext) Then
                Process.Start(path & type & ext)
            End If
        Next
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_fee1")
        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_fee2")
        End If
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_fee3")
        End If
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_migration")
        End If
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_marks_sheet")
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_character")
        End If
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_study_cert")
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_profile")
        End If
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_physical_med")
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_medical")
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim quota As String = ComboBox2.GetItemText(ComboBox2.SelectedItem)

        Dim query As String = "SELECT * From student WHERE quota='" & quota & "'"

        Dim connection As New OleDbConnection(connectionString)
        Dim dataAdp As New OleDbDataAdapter(query, connection)
        Dim ds1 As New DataSet

        connection.Open()
        dataAdp.Fill(ds1, "student")
        connection.Close()

        DataGridView1.DataSource = ds1
        DataGridView1.DataMember = "student"
    End Sub

    Private Function Check_img(ad_no As String, type As String)
        Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        Dim loca As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\" & ad_no & "\" & ad_no & "_" & type

        Dim exts() As String = {".png", ".jpg", ".bmp"}

        For Each cat As String In exts
            'MessageBox.Show(loca & cat)
            If System.IO.File.Exists(loca & cat) Then
                Return (loca & cat)
            End If
        Next

        Return ""

    End Function

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click

        Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        Dim loca As String = "|%USERPROFILE%|\Pictures\index.html".Replace("|%USERPROFILE%|", src)
        Dim i As Integer
        ' Try to get the current selected row index
        Try
            i = DataGridView1.CurrentRow.Index
        Catch ex As Exception
            MessageBox.Show("Please Select the row to be exported!")
            Return
        End Try

        ' Get the data and store it in an array
        ' Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        ' Dim loca As String = "|%USERPROFILE%|\Pictures\index.html".Replace("|%USERPROFILE%|", src)

        ' Get the admission number so we can access the picture files
        Dim num As String = DataGridView1.Item(1, i).Value
        'MessageBox.Show(num)
        Dim profile, m_profile, f_profile, g_profile As String

        ' Check if the file exists
        profile = Check_img(num, "profile")
        m_profile = Check_img(num, "m_profile")
        f_profile = Check_img(num, "f_profile")
        g_profile = Check_img(num, "g_profile")

        Dim fs As FileStream = Nothing
        'If (Not File.Exists(loca)) Then
        'File.Create(loca)
        'End If
        Dim j As Integer = 0

        Dim edu As New List(Of String) _
            From {"Board Name", "State", "Year Of Grad", "10th Total Marks", "10th Percentage"}
        Dim per As Double = (Integer.Parse(DataGridView1.Item(46, i).Value) / Integer.Parse(DataGridView1.Item(47, i).Value)) * 100
        'MessageBox.Show("The value: " & per.ToString)
        Dim edu_data As New List(Of String) _
            From {DataGridView1.Item(40, i).Value, DataGridView1.Item(41, i).Value, DataGridView1.Item(42, i).Value, DataGridView1.Item(46, i).Value, per.ToString}

        Dim hostel As New List(Of String) _
            From {"Hostel Name:", "Hostel Block:", "Hostel Room No:", "", ""}
        Dim hostel_data As New List(Of String) _
            From {DataGridView1.Item(51, i).Value, DataGridView1.Item(52, i).Value, DataGridView1.Item(53, i).Value, "", ""}
        ' xo
        Dim father As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Office Address"}
        Dim father_data As New List(Of String) _
            From {DataGridView1.Item(21, i).Value, DataGridView1.Item(22, i).Value, DataGridView1.Item(23, i).Value, DataGridView1.Item(24, i).Value, DataGridView1.Item(25, i).Value, DataGridView1.Item(26, i).Value}
        ' xo
        Dim mother As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Office Address"}
        Dim mother_data As New List(Of String) _
            From {DataGridView1.Item(27, i).Value, DataGridView1.Item(28, i).Value, DataGridView1.Item(29, i).Value, DataGridView1.Item(30, i).Value, DataGridView1.Item(31, i).Value, DataGridView1.Item(32, i).Value}
        ' xo
        Dim guard As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Address", "Office Address"}
        Dim guard_data As New List(Of String) _
            From {DataGridView1.Item(33, i).Value, DataGridView1.Item(34, i).Value, DataGridView1.Item(35, i).Value, DataGridView1.Item(36, i).Value, DataGridView1.Item(37, i).Value, DataGridView1.Item(38, i).Value, DataGridView1.Item(39, i).Value}
        ' xo 
        Dim basic As New List(Of String) _
            From {"First Name", "Middle Name", "Last Name", "DOB", "Gender", "E-mail", " Addr Line 1", "Addr Line 2", "Addr Line 3", "Mother Tongue", "Blood Group", "Religion"}
        ' xo
        Dim basic_data As New List(Of String) _
            From {DataGridView1.Item(5, i).Value, DataGridView1.Item(6, i).Value, DataGridView1.Item(7, i).Value, DataGridView1.Item(8, i).Value, DataGridView1.Item(9, i).Value, DataGridView1.Item(10, i).Value, DataGridView1.Item(11, i).Value, DataGridView1.Item(12, i).Value, DataGridView1.Item(13, i).Value, DataGridView1.Item(14, i).Value, DataGridView1.Item(15, i).Value, DataGridView1.Item(16, i).Value}
        ' xo
        Dim college As New List(Of String) _
            From {"Admission Number", "Admission Date", "Academic Year", "Merit Number", "Register No", "Year", "Valid Upto", "Current Sem", "Mobile No:", "Caste", "Category", "Transfer Cert No"}
        ' xo                                                                                                                                                                    "Year",                             "Valid Upto",                   "Current Sem",                              "Mobile No:",                   "Caste",                            "Category",                 "Transfer Cert No"
        Dim college_data As New List(Of String) _
            From {DataGridView1.Item(1, i).Value, DataGridView1.Item(2, i).Value, DataGridView1.Item(3, i).Value, DataGridView1.Item(4, i).Value, DataGridView1.Item(0, i).Value, DataGridView1.Item(54, i).Value, DataGridView1.Item(55, i).Value, DataGridView1.Item(56, i).Value.ToString, DataGridView1.Item(20, i).Value, DataGridView1.Item(17, i).Value, DataGridView1.Item(18, i).Value, DataGridView1.Item(48, i).Value}

        Using sw As StreamWriter = New StreamWriter(loca)
            sw.Write("<!DOCTYPE html>
                <html>

                <head>
                    <meta charset='utf-8'>
                    <meta http-equiv='X-UA-Compatible' content='IE=edge'>
                    <meta name='description' content='Emulating real sheets of paper in web documents (using HTML and CSS)'>
                    <title>Student Details</title>
                    <link rel='stylesheet' type='text/css' href='style.css'>
                    <style>span{font-weight:bold;} img{border: 2px solid black} td{width:50%;word-break: break-all;}</style>
                </head>")
            sw.Write("<body class='document'>")

            ' Page 1 xD

            sw.Write("<div class='page' contenteditable='false'>")

            sw.Write("<p><img src='" & profile & "' width='200px' height='200px'></p>")
            j = 0
            sw.Write("<p>")
            For Each cat As String In basic
                sw.Write("<table style='width:100%;'><tbody><tr><td style='width:50%;'><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & basic_data(j) & "</td><td style='max-width:50%;'><span>" & college(j) & ":&nbsp;&nbsp;&nbsp;</span>" & college_data(j) & "</td></tr></tbody></table>")
                j = j + 1
            Next
            sw.Write("</p>")
            sw.Write("<p><table><tr><td style='width:77.5%;'>Father Info: </td><td style='width:80%;'>Mother Info: </td></tr></table></p>")
            sw.Write("<p><table><tr><td style='width:70%;'><img src='" & f_profile & "' width='150px' height='150px'></td><td style='width:80%;'><img src='" & m_profile & "' width='150px' height='150px'></td></tr></table></p>")
            j = 0
            sw.Write("<p>")
            For Each cat As String In father
                sw.Write("<table style='width:100%;'><tbody><tr><td ><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & father_data(j) & "</td><td><span>" & mother(j) & ":&nbsp;&nbsp;&nbsp;</span>" & mother_data(j) & "</td></tr></tbody></table>")
                j = j + 1
            Next
            sw.Write("</p>")
            sw.Write("</div>")

            ' Page 2 xD

            sw.Write("<div class='page' contenteditable='false'>")
            sw.Write("<p>Guard Info(OPTIONAL): </p><p><img src='" & g_profile & "' width='200px' height='200px'></p>")
            j = 0
            sw.Write("<p>")
            For Each cat As String In guard
                sw.Write("<table style='width:100%;'><tbody><tr><td style='width:50%;'><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & guard_data(j) & "</td></tr></tbody></table>")
                j = j + 1
            Next
            sw.Write("</p>")
            sw.Write("<p><table><tr><td style='width:88%;'>Education: </td><td style='width:86%;'>Hostel: </td></tr></table></p>")
            j = 0
            sw.Write("<p>")
            For Each cat As String In edu
                sw.Write("<table style='width:100%;'><tbody><tr><td><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & edu_data(j) & "</td><td><span>" & hostel(j) & "&nbsp;&nbsp;&nbsp;</span>" & hostel_data(j) & "</td></tr></tbody></table>")
                j = j + 1
            Next
            sw.Write("</p>")
            sw.Write("<p>Other Info: </p>")
            sw.Write("<table>")
            sw.Write("<tr><td style='width:56%;'><span>Physically Disabled: </span>" & If(DataGridView1.Item(19, i).Value = -1, "Yes", "No") & "</td>")
            sw.Write("<td><span>Kannada Medium: </span>" & If(DataGridView1.Item(43, i).Value = -1, "Yes", "No") & "</td></tr>")
            sw.Write("<tr><td style='width:56%;'><span>Rural: </span>" & If(DataGridView1.Item(44, i).Value = -1, "Yes", "No") & "</td>")
            sw.Write("<td><span>Hydr-Karnataka: </span>" & If(DataGridView1.Item(45, i).Value = -1, "Yes", "No") & "</td></tr>")
            sw.Write("</table>")

            sw.Write("<br><br><br><br><br><br><br><table cellspacing=8><tr><th style='width:24%;border-top:2px solid;'>Sign of Student</th><th style='width:26%;border-top:2px solid;'>Sign of Parent</th><th style='width:28%;border-top:2px solid;'>Sign of HOD</th><th style='width:26%;border-top:2px solid;'>Sign of Principal</th></tr></table>")

            sw.Write("</div>")

            sw.Write("</body>")
            sw.Write("</html>")

        End Using

        ' Just run the browser to display the document
        Process.Start("cmd", "/c explorer " & loca)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        changeNo.Show()
        DataGridView1.Refresh()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "f_profile", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png| PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                save_files(.FileName, "m_profile", Path.GetExtension(.FileName))
                Button13_Click(sender, e)
            End If
        End With
        Button13_Click(sender, e)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_f_profile")
        End If
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        If String.Compare(Label6.Text, "Available") = 0 Then
            viewFile("_m_profile")
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        about.Show()
    End Sub
End Class