Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel
Imports MaterialSkin
Imports System.IO

Public Class princ_browse

    Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
    Dim source1 As New BindingSource()

    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source ="
    Dim ds As New DataSet()

    Dim paths() As String = {princ_main.cs, princ_main.ce, princ_main.meg, princ_main.ec, princ_main.eee}

    Private Sub princ_browse_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)


        For Each path In paths
            Dim connection As New OleDbConnection(connectionString + path)
            Dim query As String = "SELECT * FROM student"
            Dim dataAdp As New OleDbDataAdapter(query, connection)

            connection.Open()
            dataAdp.Fill(ds, "student")
            connection.Close()
        Next

        source1.DataSource = ds.Tables(0)

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "student"
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Dim i As Int16, j As Int16

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        ' Add the Column Names
        For k As Integer = 0 To DataGridView1.Columns.Count - 1
            'MessageBox.Show("Time lah")
            xlWorkSheet.Cells(1, k + 1) = DataGridView1.Columns(k).HeaderText
        Next

        ' Adding the data for Columns
        For i = 0 To DataGridView1.RowCount - 2
            For j = 0 To DataGridView1.ColumnCount - 1
                xlWorkSheet.Cells(i + 2, j + 1) = DataGridView1(j, i).Value.ToString()
            Next
        Next
        xlWorkSheet.Rows.RowHeight = 50
        xlWorkSheet.Columns.ColumnWidth = 15

        Dim loca As String = ""

        SaveFileDialog1.Filter = "Excel(*.xls)| *xls"
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            loca = SaveFileDialog1.FileName
        Else
            MessageBox.Show("File is not saved")
            loca = ""
        End If

        Try
            xlWorkBook.SaveAs(loca, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
            Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            MessageBox.Show("Exported file to location: " & loca)
        Catch ex As Exception
            xlWorkBook.Close(False, misValue, misValue)
            xlApp.Quit()
        End Try
        'xlWorkBook.Close(True, misValue, misValue)
        'xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.Compare(ComboBox1.GetItemText(ComboBox1.SelectedItem), "Register") = 0 Then
            source1.Filter = "[reg_no] = '" & TextBox1.Text & "'"
            'MessageBox.Show("Regno")
        ElseIf String.Compare(ComboBox1.GetItemText(ComboBox1.SelectedItem), "First Name") = 0 Then
            source1.Filter = "[first_name] LIKE '*" & TextBox1.Text & "*'"
            'MessageBox.Show("F_name")
        ElseIf String.Compare(ComboBox1.GetItemText(ComboBox1.SelectedItem), "Department") = 0 Then
            source1.Filter = "[reg_no] LIKE '*" & TextBox1.Text & "*'"
            'MessageBox.Show("Dept")
        ElseIf String.Compare(ComboBox1.GetItemText(ComboBox1.SelectedItem), "Semester") = 0 Then
            source1.Filter = "[current_sem] = " & TextBox1.Text
            'MessageBox.Show("Sem")
        ElseIf String.Compare(ComboBox1.GetItemText(ComboBox1.SelectedItem), "Quota") = 0 Then
            source1.Filter = "[quota] = '" & TextBox1.Text & "'"
            'MessageBox.Show("Hello")
            'MessageBox.Show("Sem")
        Else
            MessageBox.Show("Please select a search type from the combo box")
            'MessageBox.Show("Selected: " & ComboBox1.SelectedText)
            Return
        End If

        DataGridView1.DataSource = source1
        DataGridView1.Refresh()
    End Sub

    Private Sub Check(path As String, no As String)
        Dim i As Integer = 0
        Dim j As Integer = 2
        Dim dest As String
        Dim types As New List(Of String) _
            From {"_profile.jpg", "_physical_med.jpg", "_medical.jpg", "_fee1.jpg", "_fee2.jpg", "_fee3.jpg", "_study_cert.jpg", "_character.jpg", "_migration.jpg", "_marks_sheet.jpg",
            "_profile.png", "_physical_med.png", "_medical.png", "_fee1.png", "_fee2.png", "_fee3.png", "_study_cert.png", "_character.png", "_migration.png", "_marks_sheet.png",
            "_physical_med.pdf", "_medical.pdf", "_fee1.pdf", "_fee2.pdf", "_fee3.pdf", "_study_cert.pdf", "_character.pdf", "_migration.pdf", "_marks_sheet.pdf"
        }
        Dim labels As New List(Of Label) _
            From {Label2, Label3, Label4, Label5, Label6, Label7, Label8, Label9, Label10, Label11}
        For Each type As String In types
            dest = path & "\" & no & type
            'MessageBox.Show(dest)
            If File.Exists(dest) Then
                labels(i).Visible = True
            End If
            Try
                i = i + 1
            Catch ex As Exception
                j = 0
            End Try
        Next
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim admin_no, base, path As String
        admin_no = TextBox2.Text
        Dim depts As New List(Of String) _
            From {"hod_cs\", "hod_ce\", "hod_me\", "hod_ec\", "hod_eee\"}

        base = "|%USERPROFILE%|\Pictures\"

        For Each dept As String In depts
            path = base.Replace("|%USERPROFILE%|", src) & dept & "img\" & admin_no
            'MessageBox.Show(path)
            If Directory.Exists(path) Then
                Check(path, admin_no)
                Return
            End If
        Next

        MessageBox.Show("Admission Number is not found!")

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)
        about.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ds As New DataSet()
        For Each path In paths
            Dim connection As New OleDbConnection(connectionString + path)
            Dim query As String = "SELECT * FROM student"
            Dim dataAdp As New OleDbDataAdapter(query, connection)

            connection.Open()
            dataAdp.Fill(ds, "student")
            connection.Close()
        Next

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "student"
    End Sub

    Private Sub Check_it(path As String, no As String, data As String)
        Dim i As Integer = 0
        Dim dest As String
        Dim types As New List(Of String) _
            From {
            data & ".jpg", data & ".png", data & ".pdf"
        }
        For Each type As String In types
            dest = path & "\" & no & "_" & type
            'MessageBox.Show(dest)
            If File.Exists(dest) Then
                Process.Start(dest)
            End If
        Next
    End Sub

    Private Sub getDirect(type As String)
        Dim admin_no, base, path As String
        admin_no = TextBox2.Text
        Dim depts As New List(Of String) _
            From {"hod_cs\", "hod_ce\", "hod_me\", "hod_ec\", "hod_eee\"}

        base = "|%USERPROFILE%|\Pictures\"

        For Each dept As String In depts
            path = base.Replace("|%USERPROFILE%|", src) & dept & "img\" & admin_no
            'MessageBox.Show(path)
            If Directory.Exists(path) Then
                Check_it(path, admin_no, type)
                Return
            End If
        Next
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        If Label2.Visible = True Then
            getDirect("profile")
        End If
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        If Label3.Visible = True Then
            getDirect("physical_med")
        End If
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        If Label4.Visible = True Then
            getDirect("medical")
        End If
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        If Label5.Visible = True Then
            getDirect("fee1")
        End If
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        If Label6.Visible = True Then
            getDirect("fee2")
        End If
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        If Label7.Visible = True Then
            getDirect("fee3")
        End If
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        If Label8.Visible = True Then
            getDirect("study_cert")
        End If
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        If Label9.Visible = True Then
            getDirect("character")
        End If
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        If Label10.Visible = True Then
            getDirect("migration")
        End If
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        If Label11.Visible = True Then
            getDirect("marks_sheet")
        End If
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        MessageBox.Show("Management Karnataka - 'MNGT KAR', Management Non-Karnataka - 'MNGT Non-KAR', Government - 'GOV', SNQ - 'SNQ'")
    End Sub
End Class