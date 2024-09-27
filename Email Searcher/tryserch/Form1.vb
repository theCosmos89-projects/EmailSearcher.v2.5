
Imports System.Drawing.Text
Imports System.IO

Public Class Form1
    Private RunOnce As Boolean = False
    Private bycosmos As String = "Email Searcher by cosmos89"
    Public doms As String = Path.Combine(Path.GetTempPath(), "All-TLD.txt")
    Private textBefore As String = ""
    Private lastDirectory As String = ""
    Private newEmails As String
    Private folder As String
    Private installFont As Boolean = False
    Private selectAll As Boolean = False
    Private isDeleting As Boolean = False
    Private isFilter As Boolean = False
    Private isFirst As Boolean = True
    Private customFontCollection As New System.Drawing.Text.PrivateFontCollection
    Public stopSearch As Boolean = False
    Public isOn As Boolean = False
    Public gCount As Integer = 0
    Private myToolTip As New ToolTip
    Private timer As New Stopwatch
    Private originalItems As List(Of String)
    Private WithEvents Ex As New MailExtractor
    Delegate Sub processFileDelegate(ByVal path As String)
    Declare Function SendMessageLB Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Int32, ByVal wParam As UInt16, ByVal lParam As UInt16) As Int32
    Const LB_SELITEMRANGEEX As Int32 = &H183


    Private Sub LoadAndApplyCustomFont()

        Dim fontPath As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "EmailSearcher_Font.ttf")

        If Not File.Exists(fontPath) Then
            IO.File.WriteAllBytes(fontPath, My.Resources.font)
        End If

        customFontCollection.Dispose()
        customFontCollection = New PrivateFontCollection()
        customFontCollection.AddFontFile(fontPath)

        Label1.Font = New Font(customFontCollection.Families(0), 38)
        Button1.Font = New Font(customFontCollection.Families(0), 36)
        Button2.Font = New Font(customFontCollection.Families(0), 11)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            LoadAndApplyCustomFont()

        Catch ex As Exception
            MessageBox.Show("Error to loading and applying custom font", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try


        Try

            If Not File.Exists(doms) Then File.WriteAllText(doms, My.Resources.domains)

        Catch ex As Exception
            MessageBox.Show("Error to generate domains file", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try



        ContextMenuStrip1.Enabled = False
        ContextMenuStrip2.Enabled = False
        Label11.Text = "Filter Intensity:  " & TrackBar1.Value.ToString()
    End Sub
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        MessageBox.Show("This is a free tool. Its function is to find and list all the email addresses stored in the files of a folder and its sub-folders. I hope you find it useful.", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Button2_Click(sender, e)

    End Sub

    Private Sub Applyallfiles(ByVal folder As String, ByVal extension As String, ByVal fileaction As processFileDelegate)

        If stopSearch Or Not isOn Then Exit Sub


        For Each File In Directory.GetFiles(folder, extension)
            Try

                fileaction.Invoke(File)

            Catch ex As Exception
                Exit Sub
            End Try
        Next

        For Each subdir In Directory.GetDirectories(folder)

            Try

                Applyallfiles(subdir, extension, fileaction)

            Catch ex As Exception
                Continue For
            End Try
        Next

    End Sub


    Private Sub ProcessFile(ByVal filePath As String)

        If stopSearch Or Not isOn Then Exit Sub


        Try
            Dim fileInfo As New System.IO.FileInfo(filePath)
            If fileInfo.Length < Integer.Parse(TextBox3.Text) * 1000 Then
                Dim currentDirectory As String = System.IO.Path.GetDirectoryName(filePath)

                If currentDirectory <> lastDirectory Then
                    lastDirectory = currentDirectory
                    TextBox1.Text = currentDirectory
                End If

                Dim fileName As String = System.IO.Path.GetFileName(filePath)
                Label4.Text = "Searching in " & fileName

                Label4.Update()
                TextBox1.Update()

                Ex.ExtractEmails(filePath)

            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Function GetElapsedTime(milliseconds As Long) As String
        Dim timeSpan As TimeSpan = TimeSpan.FromMilliseconds(milliseconds)

        Dim days As Integer = timeSpan.Days
        Dim hours As Integer = timeSpan.Hours
        Dim minutes As Integer = timeSpan.Minutes
        Dim seconds As Integer = timeSpan.Seconds
        Dim millisecondsRemaining As Integer = timeSpan.Milliseconds

        Dim result As String = ""

        If days > 0 Then result &= days & " days "
        If hours > 0 Then result &= hours & " hr "
        If minutes > 0 Then result &= minutes & " min "

        If seconds > 0 Then
            result &= seconds & " sec"
        ElseIf millisecondsRemaining >= 0 Then
            result &= millisecondsRemaining & " ms"
        End If

        Return result.Trim()
    End Function
    Private Sub Finish()
        Try
            ison = False
            stopSearch = True
            Button1.Text = "SEARCH"
            Button1.Enabled = False
            TextBox1.Text = folder
            If ListBox2.Items.Count = 0 Then
                Label7.Visible = False
                Label4.Visible = False
            Else
                Label4.Text = "Double click to open - Right click to options"
                Label7.Visible = True
            End If
            If CheckBox2.Checked = True And gcount > 0 And Not isfirst Then
                newemails = " new"
            End If

            Dim elapsedTime As String = GetElapsedTime(timer.ElapsedMilliseconds)
            Dim message As String = gcount & newemails & " emails found in " & elapsedTime

            MessageBox.Show(message, bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)


            originalItems = ListBox2.Items.Cast(Of String)().ToList()

            timer.Reset()
            Button1.Enabled = True
            Button2.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Label2.Enabled = True
            Label3.Enabled = True
            Label11.Enabled = True
            TrackBar1.Enabled = True
            If isfirst Then isfirst = False

            If ListBox1.Items.Count > 0 Then
                Label8.Enabled = True
                TextBox4.Enabled = True
            End If

            Button1.Focus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Button2.Font = New Font(customFontCollection.Families(0), 11)
        Catch es As Exception
        End Try


        Try

            Using dialog As New FolderBrowserDialog()
                dialog.Description = "Choose a folder to search for email addresses..."
                dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                dialog.ShowNewFolderButton = True

                If dialog.ShowDialog() = DialogResult.OK Then
                    TextBox1.Text = dialog.SelectedPath
                    Button1.Enabled = True
                    Button1.Enabled = True
                    Button2.Enabled = True
                    CheckBox1.Enabled = True
                    CheckBox2.Enabled = True
                    TextBox2.Enabled = True
                    TextBox3.Enabled = True
                    Label2.Enabled = True
                    Label3.Enabled = True
                    Label11.Enabled = True
                    TrackBar1.Enabled = True
                    Button1.Select()
                Else

                    If TextBox1.Text = Nothing Then
                        MessageBox.Show("To search for emails, just select a folder", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)
                        'Button2_Click(sender, e)
                    End If
                End If




            End Using
        Catch ex As Exception
            MessageBox.Show("Error while selecting the folder.", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Button1.Font = New Font(customFontCollection.Families(0), 36)
        Catch es As Exception
        End Try

        If isOn Then
            Try
                stopSearch = True
                Finish()
                Exit Sub

            Catch ex As Exception
                Exit Sub
            End Try

        End If

        If Not isOn Then


            Dim num As Integer = TextBox3.Text

            If num <= 0 Then TextBox3.Text = "1"

            TextBox3.Text = TextBox3.Text.TrimStart("0"c)
            TextBox3.Refresh()
            TextBox4.Clear()

            If CheckBox2.Checked = False Then
                ListBox2.Items.Clear()
            End If


            ListBox1.Items.Clear()
            ListBox2.ClearSelected()

            Aleft1()
            Aleft2()

            Label7.Visible = False
            newEmails = Nothing

            gCount = 0


            Label5.Text = "Emails found: " & ListBox2.Items.Count
            Label6.Text = "Files with emails: " & ListBox1.Items.Count

            Dim epath As String = TextBox1.Text
            Dim ext As String = "*." & TextBox2.Text

            folder = TextBox1.Text
            stopSearch = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            TextBox2.Enabled = False
            TextBox3.Enabled = False
            Button2.Enabled = False
            Label2.Enabled = False
            Label3.Enabled = False
            Label4.Visible = True
            Label8.Enabled = False
            TextBox4.Enabled = False
            Label11.Enabled = False
            TrackBar1.Enabled = False
            ContextMenuStrip1.Enabled = False
            ContextMenuStrip2.Enabled = False

            Button1.Text = "STOP"

            isOn = True

            Label9.Update()
            Label10.Update()
            Button1.Update()

            Me.Focus()



            Try

                Dim runprocess As processFileDelegate = AddressOf ProcessFile
                timer.Start()

                Applyallfiles(epath, ext, runprocess)

                timer.Stop()

            Catch ex As Exception
                MessageBox.Show("An unexpected error occurred while searching", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finish()
                Exit Sub
            End Try

            If stopSearch = False Then
                Finish()
            End If
        End If

    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        Try
            Dim DataObject As New DataObject
            Dim tempFileArray As New ArrayList
            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1


                tempFileArray.Add(ListBox1.SelectedItems(i).ToString)

            Next
            DataObject.SetData(DataFormats.FileDrop, False, DirectCast(tempFileArray.ToArray(GetType(String)), String()))
            Clipboard.SetDataObject(DataObject)


            MessageBox.Show("File(s) copied to clipboard. You can paste on any folder/desktop", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Unable to copy the file(s)", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try
    End Sub

    Private Sub OpenFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFolderToolStripMenuItem.Click
        Try

            Process.Start("explorer.exe", Path.GetDirectoryName(ListBox1.SelectedItems(0).ToString))

        Catch ex As Exception
            MessageBox.Show("Unable to open the folder", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub
    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        TextBox2.Text = ""
    End Sub
    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If TextBox2.Text = "" Then TextBox2.Text = "*"
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Try
            Dim webAddress As String = "https://chathispano.com/webchat?theme=embed&style=orange&title=Canalcosmos&logo=https%3A%2F%2Fcdn.chathispano.com%2Fnews%2Fesquina.jpg&autojoin=false&autoload=false&nick=cosmonauta&chan=cosmos"
            Process.Start(webAddress)
        Catch Ex As Exception
        End Try
    End Sub

    Private Sub CopyToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem1.Click
        Try

            Clipboard.SetText(String.Join(Environment.NewLine, ListBox2.SelectedItems.Cast(Of String).ToArray))
            MessageBox.Show("Email(s) strings copied to clipboard successfully", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)


        Catch ex As Exception
            MessageBox.Show("Unable to copy the email strings to clipboard", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End Try
    End Sub

    Private Async Sub DeleteFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteFilesToolStripMenuItem.Click
        Try
            Dim result As DialogResult = MessageBox.Show("Are you certain you wish to delete the file(s)?", bycosmos, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                isDeleting = True
                Dim archivosAEliminar As New List(Of String)
                Dim archivosEliminados As Boolean = False
                Dim erroresAlEliminar As New List(Of String)


                For Each selectedIndex As Integer In ListBox1.SelectedIndices
                    If selectedIndex >= 0 AndAlso selectedIndex < ListBox1.Items.Count Then
                        Dim archivoSeleccionado As String = ListBox1.Items(selectedIndex).ToString()
                        archivosAEliminar.Add(archivoSeleccionado)
                    End If
                Next


                Dim tasks As List(Of Task) = archivosAEliminar.Select(Function(archivo) Task.Run(Sub()
                                                                                                     Try
                                                                                                         File.Delete(archivo)
                                                                                                         archivosEliminados = True
                                                                                                     Catch ex As Exception
                                                                                                         ' Almacenar el archivo en la lista de errores
                                                                                                         erroresAlEliminar.Add(archivo)
                                                                                                     End Try
                                                                                                 End Sub)).ToList()


                Await Task.WhenAll(tasks)


                Me.Invoke(Sub()

                              If erroresAlEliminar.Count > 0 Then
                                  MessageBox.Show($"Unable to delete the following files: {String.Join(", ", erroresAlEliminar)}", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                              End If


                              For Each archivo As String In archivosAEliminar
                                  If Not File.Exists(archivo) Then
                                      ListBox1.Items.Remove(archivo)
                                  End If
                              Next


                              Label6.Text = "Files with emails: " & ListBox1.Items.Count


                              If archivosEliminados Then
                                  MessageBox.Show("File(s) deleted successfully", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)
                              End If
                          End Sub)
            End If
            isDeleting = False
            SelectEmails()
        Catch ex As Exception

            MessageBox.Show("The file(s) cannot be deleted at the moment", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Try


            Process.Start(ListBox1.SelectedItems(0).ToString)

        Catch ex As Exception
            MessageBox.Show("Unable to open this file", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub


    Private Sub CheckBox1_CheckStateChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckStateChanged
        If Not isFilter And CheckBox1.Checked = False Then

            MessageBox.Show("If unchecked this option may result in the scan including false email addresses", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Information)

            isFilter = True

        End If
    End Sub

    Private Sub CopyToNotepadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToNotepadToolStripMenuItem.Click
        Try
            Dim dateString = DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss")
            Dim filepath As String = System.IO.Path.GetTempPath() & "\Emails(" & dateString & ").txt"

            Dim objWriter As New System.IO.StreamWriter(filepath, True)
            For Each Item As Object In ListBox2.SelectedItems
                objWriter.WriteLine(Item.ToString)
            Next

            objWriter.Close()

            Process.Start(filepath)
        Catch ex As Exception
            MessageBox.Show("Unable to copy the email strings to Notepad", bycosmos, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub
    Private Sub Clean()
        ListBox1.ClearSelected()
        ListBox2.ClearSelected()
        ContextMenuStrip1.Enabled = False
        ContextMenuStrip2.Enabled = False
    End Sub
    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click
        Clean()
    End Sub
    Private Sub Tooltip(ltb As ListBox)
        Dim maxElementos As Integer = 25
        Dim elementosSeleccionados = ltb.SelectedItems.
        Cast(Of Object)().
        Select(Function(item)

                   Return System.IO.Path.GetFileName(item.ToString())
               End Function).
        Take(maxElementos).
        ToList()

        Dim cadenaElementos As String = String.Join(Environment.NewLine, elementosSeleccionados)

        If ltb.SelectedItems.Count > maxElementos Then
            cadenaElementos &= Environment.NewLine & "..."
        End If

        mytooltip.SetToolTip(ltb, cadenaElementos)
    End Sub
    Private Sub SelectFiles()

        Dim archivosSeleccionados As New HashSet(Of String)(ListBox2.SelectedItems.Cast(Of String)())


        Dim correosRelacionados As New HashSet(Of String)

        For Each archivoSeleccionado As String In archivosSeleccionados
            If Ex.EmailFileRelation.ContainsKey(archivoSeleccionado) Then
                correosRelacionados.UnionWith(Ex.EmailFileRelation(archivoSeleccionado))
            End If
        Next


        ListBox1.ClearSelected()


        For Each correoRelacionado As String In correosRelacionados
            Dim index As Integer = ListBox1.FindStringExact(correoRelacionado)
            If index >= 0 Then
                ListBox1.SetSelected(index, True)
            End If
        Next
    End Sub



    Private Sub SelectEmails()

        Dim archivosSeleccionados As New HashSet(Of String)(ListBox1.SelectedItems.Cast(Of String)())


        Dim correosRelacionados As New HashSet(Of String)

        For Each archivoSeleccionado As String In archivosSeleccionados
            If Ex.FileEmailRelation.ContainsKey(archivoSeleccionado) Then
                correosRelacionados.UnionWith(Ex.FileEmailRelation(archivoSeleccionado))
            End If
        Next


        ListBox2.ClearSelected()


        For Each correoRelacionado As String In correosRelacionados
            Dim index As Integer = ListBox2.FindStringExact(correoRelacionado)
            If index >= 0 Then
                ListBox2.SetSelected(index, True)
            End If
        Next
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        Try

            If Me.ListBox1.SelectedIndex >= 0 Then

                ContextMenuStrip1.Enabled = True



                If ListBox1.Focused Then
                    If Not isDeleting Then
                        SelectEmails()
                    End If
                End If
                Else

                ContextMenuStrip1.Enabled = False

            End If
            Aleft1()
        Catch Ex As Exception
        End Try

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

        Try
            If Me.ListBox2.SelectedIndex >= 0 Then


                ContextMenuStrip2.Enabled = True



                If ListBox2.Focused Then

                    SelectFiles()

                End If


                Else
                ContextMenuStrip2.Enabled = False

            End If
            Aleft2()
        Catch Ex As Exception
        End Try
    End Sub

    Private Sub Filter_result()

        Try
            Dim searchTerm As String = TextBox4.Text()

            If searchTerm <> textBefore Then

                ListBox2.Items.Clear()

                Dim filteredItems = originalItems.Where(Function(item) item.Contains(searchTerm))

                ListBox2.Items.AddRange(filteredItems.ToArray())
                textBefore = searchTerm
            End If
        Catch Ex As Exception
                    End Try

    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

        Try

            ListBox1.ClearSelected()
            ListBox2.ClearSelected()
            filter_result()
            Label5.Text = "Emails found: " & ListBox2.Items.Count

        Catch ex As Exception
        End Try

    End Sub


    Private Sub ListBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles ListBox1.MouseMove

        Try
            If Not selectAll Then
                Dim index As Integer = ListBox1.IndexFromPoint(e.Location)
                If ListBox1.SelectedIndices.Contains(index) Then
                    Tooltip(ListBox1)
                Else
                    myToolTip.SetToolTip(ListBox1, "")
                End If
            End If
        Catch Ex As Exception
        End Try

    End Sub

    Private Sub ListBox2_MouseMove(sender As Object, e As MouseEventArgs) Handles ListBox2.MouseMove

        Try
            If Not selectAll Then
                Dim index As Integer = ListBox2.IndexFromPoint(e.Location)

                If ListBox2.SelectedIndices.Contains(index) Then
                    Tooltip(ListBox2)
                Else
                    myToolTip.SetToolTip(ListBox2, "")
                End If
            End If
        Catch Ex As Exception
        End Try

    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress

        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Button1.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Button1.Focus()
        End If
    End Sub
    Private Sub Aleft1()
        Dim rightAlign As Integer = Label9.Left + Label9.Width
        Label9.Text = "Sel: " & ListBox1.SelectedIndices.Count()
        Label9.Left = rightAlign - Label9.Width
    End Sub
    Private Sub Aleft2()
        Dim rightAlign As Integer = Label10.Left + Label10.Width
        Label10.Text = "Sel: " & ListBox2.SelectedIndices.Count()
        Label10.Left = rightAlign - Label10.Width
    End Sub
    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click

        Try
            selectall = True
            ListBox1.ClearSelected()
            SendMessageLB(Me.ListBox2.Handle, LB_SELITEMRANGEEX, 0, CUShort(Me.ListBox2.Items.Count - 1))
            Aleft2()
            selectall = False
        Catch Ex As Exception
        End Try

    End Sub
    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label11.Text = "Filter Intensity:  " & TrackBar1.Value.ToString()
    End Sub
    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        Clean()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Clean()
    End Sub


End Class
