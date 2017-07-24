Imports MaterialSkin
Imports System.Data.OleDb
Imports System.IO

Public Class insert
    Public total, max, percent As Decimal
    Dim profile, med, physical, fee1, fee2, fee3, study, character, mig, marks, f_profile, m_profile, g_profile As String

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                med = .FileName
            End If
        End With
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                fee1 = .FileName
            End If
        End With
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                study = .FileName
            End If
        End With
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                character = .FileName
            End If
        End With
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                marks = .FileName
            End If
        End With
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                fee3 = .FileName
            End If
        End With
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        Dim loca As String = "|%USERPROFILE%|\Pictures\index.html".Replace("|%USERPROFILE%|", src)

        Dim fs As FileStream = Nothing
        'If (Not File.Exists(loca)) Then
        'File.Create(loca)
        'End If
        Dim i As Integer = 0

        Dim edu As New List(Of String) _
            From {"Board Name", "State", "Year Of Grad", "10th Total Marks", "10th Percentage"}
        Dim edu_data As New List(Of String) _
            From {TextBox38.Text, TextBox39.Text, TextBox40.Text, Label70.Text, Label73.Text}

        Dim hostel As New List(Of String) _
            From {"Hostel Name:", "Hostel Block:", "Hostel Room No:", "", ""}
        Dim hostel_data As New List(Of String) _
            From {TextBox61.Text, TextBox62.Text, TextBox63.Text, "", ""}
        ' xo
        Dim father As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Office Address"}
        Dim father_data As New List(Of String) _
            From {TextBox19.Text, TextBox20.Text, TextBox21.Text, TextBox22.Text, TextBox23.Text, TextBox24.Text}
        ' xo
        Dim mother As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Office Address"}
        Dim mother_data As New List(Of String) _
            From {TextBox30.Text, TextBox29.Text, TextBox28.Text, TextBox27.Text, TextBox26.Text, TextBox25.Text}
        ' xo
        Dim guard As New List(Of String) _
            From {"Name", "Occupation", "Mobile No", "E-mail", "Landline", "Address", "Office Address"}
        Dim guard_data As New List(Of String) _
            From {TextBox35.Text, TextBox34.Text, TextBox33.Text, TextBox32.Text, TextBox31.Text, TextBox36.Text, TextBox37.Text}
        ' xo 
        Dim basic As New List(Of String) _
            From {"First Name", "Middle Name", "Last Name", "DOB", "Gender", "E-mail", " Addr Line 1", "Addr Line 2", "Addr Line 3", "Mother Tongue", "Blood Group", "Religion"}
        ' xo
        Dim basic_data As New List(Of String) _
            From {TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, ComboBox1.Text, TextBox5.Text, TextBox8.Text, TextBox7.Text, TextBox6.Text, TextBox9.Text, TextBox10.Text, TextBox11.Text, TextBox12.Text, TextBox13.Text}
        ' xo
        Dim college As New List(Of String) _
            From {"Admission Number", "Admission Date", "Academic Year", "Merit Number", "Register No", "Year", "Valid Upto", "Current Sem", "Mobile No:", "Caste", "Category", "Transfer Cert No"}
        ' xo
        Dim college_data As New List(Of String) _
            From {TextBox17.Text, TextBox16.Text, TextBox15.Text, TextBox18.Text, TextBox60.Text, TextBox64.Text, TextBox65.Text, TextBox66.Text, TextBox14.Text, TextBox12.Text, TextBox13.Text, TextBox59.Text}

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
            i = 0
            sw.Write("<p>")
            For Each cat As String In basic
                sw.Write("<table style='width:100%;'><tbody><tr><td style='width:50%;'><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & basic_data(i) & "</td><td style='max-width:50%;'><span>" & college(i) & ":&nbsp;&nbsp;&nbsp;</span>" & college_data(i) & "</td></tr></tbody></table>")
                i = i + 1
            Next
            sw.Write("</p>")
            sw.Write("<p><table><tr><td style='width:77.5%;'>Father Info: </td><td style='width:80%;'>Mother Info: </td></tr></table></p>")
            sw.Write("<p><table><tr><td style='width:70%;'><img src='" & f_profile & "' width='150px' height='150px'></td><td style='width:80%;'><img src='" & m_profile & "' width='150px' height='150px'></td></tr></table></p>")
            i = 0
            sw.Write("<p>")
            For Each cat As String In father
                sw.Write("<table style='width:100%;'><tbody><tr><td ><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & father_data(i) & "</td><td><span>" & mother(i) & ":&nbsp;&nbsp;&nbsp;</span>" & mother_data(i) & "</td></tr></tbody></table>")
                i = i + 1
            Next
            sw.Write("</p>")
            sw.Write("</div>")

            ' Page 2 xD

            sw.Write("<div class='page' contenteditable='false'>")
            sw.Write("<p>Guard Info(OPTIONAL): </p><p><img src='" & g_profile & "' width='200px' height='200px'></p>")
            i = 0
            sw.Write("<p>")
            For Each cat As String In guard
                sw.Write("<table style='width:100%;'><tbody><tr><td style='width:50%;'><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & guard_data(i) & "</td></tr></tbody></table>")
                i = i + 1
            Next
            sw.Write("</p>")
            sw.Write("<p><table><tr><td style='width:88%;'>Education: </td><td style='width:86%;'>Hostel: </td></tr></table></p>")
            i = 0
            sw.Write("<p>")
            For Each cat As String In edu
                sw.Write("<table style='width:100%;'><tbody><tr><td><span>" & cat & ":&nbsp;&nbsp;&nbsp;</span>" & edu_data(i) & "</td><td><span>" & hostel(i) & "&nbsp;&nbsp;&nbsp;</span>" & hostel_data(i) & "</td></tr></tbody></table>")
                i = i + 1
            Next
            sw.Write("</p>")
            sw.Write("<p>Other Info: </p>")
            sw.Write("<table>")
            sw.Write("<tr><td style='width:56%;'><span>Physically Disabled: </span>" & If(RadioButton1.Checked, "Yes", "No") & "</td>")
            sw.Write("<td><span>Kannada Medium: </span>" & If(RadioButton3.Checked, "Yes", "No") & "</td></tr>")
            sw.Write("<tr><td style='width:56%;'><span>Rural: </span>" & If(RadioButton6.Checked, "Yes", "No") & "</td>")
            sw.Write("<td><span>Hydr-Karnataka: </span>" & If(RadioButton8.Checked, "Yes", "No") & "</td></tr>")
            sw.Write("</table>")

            sw.Write("<br><br><br><br><br><br><br><table cellspacing=8><tr><th style='width:24%;border-top:2px solid;'>Sign of Student</th><th style='width:26%;border-top:2px solid;'>Sign of Parent</th><th style='width:28%;border-top:2px solid;'>Sign of HOD</th><th style='width:26%;border-top:2px solid;'>Sign of Principal</th></tr></table>")

            sw.Write("</div>")

            sw.Write("</body>")
            sw.Write("</html>")

        End Using

        ' Just run the browser to display the document
        Process.Start("cmd", "/c explorer " & loca)

        'SaveFileDialog1.Filter = "PDF(*.pdf)| *.pdf"
        'If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
        'loca = SaveFileDialog1.FileName
        'End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                m_profile = .FileName
                PictureBox4.Image = Image.FromFile(m_profile)
            End If
        End With
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                g_profile = .FileName
                PictureBox5.Image = Image.FromFile(g_profile)
            End If
        End With
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                mig = .FileName
            End If
        End With
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                f_profile = .FileName
                PictureBox3.Image = Image.FromFile(f_profile)
            End If
        End With
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                fee2 = .FileName
            End If
        End With
    End Sub

    ' Handles clicking the physcial disability button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png|PDF (*.pdf)|*.pdf"
            If .ShowDialog = DialogResult.OK Then
                physical = .FileName
            End If
        End With
    End Sub

    Dim con As New OleDbConnection("Provider= Microsoft.ACE.OLEDB.12.0;Data Source =" + main.file)
    Dim cmd As OleDbCommand

    Private Sub insert_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SkinManager As MaterialSkinManager = MaterialSkinManager.Instance
        SkinManager.AddFormToManage(Me)
        SkinManager.Theme = MaterialSkinManager.Themes.LIGHT
        'SkinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE)
        SkinManager.ColorScheme = New ColorScheme(Primary.Red800, Primary.Orange800, Primary.Teal300, Accent.DeepOrange200, TextShade.WHITE)
    End Sub

    Private Sub save_files(file1 As String, type As String)
        Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
        Dim path As String = "|%USERPROFILE%|\Pictures\".Replace("|%USERPROFILE%|", src) & Form1.naame & "\img\" & TextBox17.Text
        Dim dest As String = path & "\" & TextBox17.Text & "_" & type & ".jpg"

        ' MessageBox.Show(dest)

        FileCopy(file1, dest)

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        With OpenFileDialog1
            .CheckFileExists = True
            .ShowReadOnly = True
            .Filter = "Jpeg Files (*.jpg)|*.jpg|Png Files (*.png)|*.png"
            If .ShowDialog = DialogResult.OK Then
                profile = .FileName
                PictureBox1.Image = Image.FromFile(profile)
            End If
        End With
    End Sub

    Private Sub send_query(query As String)
        con.Open()
        cmd = New OleDbCommand(query, con)
        cmd.ExecuteNonQuery()
        con.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        total = Convert.ToDecimal(TextBox43.Text) + Convert.ToDecimal(TextBox44.Text) + Convert.ToDecimal(TextBox47.Text) + Convert.ToDecimal(TextBox50.Text) + Convert.ToDecimal(TextBox53.Text) + Convert.ToDecimal(TextBox56.Text)
        max = Convert.ToDecimal(TextBox42.Text) + Convert.ToDecimal(TextBox45.Text) + Convert.ToDecimal(TextBox48.Text) + Convert.ToDecimal(TextBox51.Text) + Convert.ToDecimal(TextBox54.Text) + Convert.ToDecimal(TextBox57.Text)

        percent = Math.Round((total / max) * 100.0, 2)

        Label73.Text = Convert.ToString(percent)
        Label70.Text = Convert.ToString(total) + " / " + Convert.ToString(max)

        Label70.Show()
        Label73.Show()
    End Sub

    Public Sub ClearTextBox(parent As Control)

        For Each child As Control In parent.Controls
            ClearTextBox(child)
        Next

        If TryCast(parent, TextBox) IsNot Nothing Then
            TryCast(parent, TextBox).Text = [String].Empty
        End If

        PictureBox1.Image = PictureBox1.InitialImage
        'PictureBox2.Image = PictureBox2.InitialImage
        PictureBox3.Image = PictureBox3.InitialImage
        PictureBox4.Image = PictureBox4.InitialImage
        PictureBox5.Image = PictureBox5.InitialImage

    End Sub

    Private Sub Submit_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim query As String
        Dim phys, kar_med, rural, hyd_kar, non_kar, hotel As Integer
        phys = kar_med = rural = hyd_kar = non_kar = hotel = 0

        'med, physical, fee, study, character, mig, marks

        If RadioButton1.Checked Then
            phys = -1
        End If

        If RadioButton3.Checked Then
            kar_med = -1
        End If

        If RadioButton6.Checked Then
            rural = -1
        End If

        If RadioButton8.Checked Then
            hyd_kar = -1
        End If

        If RadioButton9.Checked Then
            non_kar = -1
        End If

        If RadioButton11.Checked Then
            hotel = -1
        End If

        'MessageBox.Show(ComboBox1.GetItemText(ComboBox1.SelectedItem))
        'MessageBox.Show(ComboBox2.GetItemText(ComboBox2.SelectedItem))

        Try
            query = "INSERT INTO student VALUES('" & TextBox60.Text & "', '" & TextBox17.Text & "', '" & TextBox16.Text & "', '" & TextBox15.Text & "', '" & TextBox18.Text & "', '" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & ComboBox1.GetItemText(ComboBox1.SelectedItem) & "', '" & TextBox5.Text & "', '" & TextBox8.Text & "', '" & TextBox7.Text & "', '" & TextBox6.Text & "', '" & TextBox9.Text & "', '" & TextBox10.Text & "', '" & TextBox11.Text & "', '" & TextBox12.Text & "', '" & TextBox13.Text & "', " & phys & ", '" & TextBox14.Text & "', '" & TextBox19.Text & "', '" & TextBox20.Text & "', '" & TextBox21.Text & "', '" & TextBox22.Text & "', '" & TextBox23.Text & "', '" & TextBox24.Text & "', '" & TextBox30.Text & "', '" & TextBox29.Text & "', '" & TextBox28.Text & "', '" & TextBox27.Text & "', '" & TextBox26.Text & "', '" & TextBox25.Text & "', '" & TextBox35.Text & "', '" & TextBox34.Text & "', '" & TextBox33.Text & "', '" & TextBox32.Text & "', '" & TextBox31.Text & "', '" & TextBox36.Text & "', '" & TextBox37.Text & "', '" & TextBox38.Text & "', '" & TextBox39.Text & "', '" & TextBox40.Text & "', " & kar_med & ", " & rural & ", " & hyd_kar & ", " & total & ", " & max & ", " & non_kar & ", '" & TextBox59.Text & "', " & hotel & ", '" & TextBox61.Text & "', '" & TextBox62.Text & "', '" & TextBox63.Text & "', '" & TextBox64.Text & "', '" & TextBox65.Text & "', " & TextBox66.Text & ", '" & ComboBox2.GetItemText(ComboBox2.SelectedItem) & "')"
            send_query(query)

            Dim src As String = Environment.ExpandEnvironmentVariables("%USERPROFILE%")
            ' Create the directory
            System.IO.Directory.CreateDirectory("|%UserProfile%|\Pictures\".Replace("|%UserProfile%|", src) & Form1.naame & "\img\" & TextBox17.Text)
            ' Create the files
            If Not String.IsNullOrEmpty(profile) Then
                save_files(profile, "profile")
            End If

            If Not String.IsNullOrEmpty(f_profile) Then
                save_files(f_profile, "f_profile")
            End If

            If Not String.IsNullOrEmpty(g_profile) Then
                save_files(f_profile, "g_profile")
            End If

            If Not String.IsNullOrEmpty(m_profile) Then
                save_files(m_profile, "m_profile")
            End If

            If Not String.IsNullOrEmpty(med) Then
                save_files(med, "medical")
            End If

            If Not String.IsNullOrEmpty(physical) Then
                save_files(physical, "physical_med")
            End If

            If Not String.IsNullOrEmpty(fee1) Then
                save_files(fee1, "fee1")
            End If

            If Not String.IsNullOrEmpty(fee2) Then
                save_files(fee2, "fee2")
            End If

            If Not String.IsNullOrEmpty(fee3) Then
                save_files(fee3, "fee3")
            End If

            If Not String.IsNullOrEmpty(study) Then
                save_files(study, "study_cert")
            End If

            If Not String.IsNullOrEmpty(character) Then
                save_files(character, "character")
            End If

            If Not String.IsNullOrEmpty(mig) Then
                save_files(mig, "migration")
            End If

            If Not String.IsNullOrEmpty(marks) Then
                save_files(marks, "marks_sheet")
            End If

            ClearTextBox(Me)

        Catch ex As Exception
            MessageBox.Show("Mandatory fields are empty.")
            MessageBox.Show(ex.Message)
            con.Close()
        End Try



    End Sub
End Class