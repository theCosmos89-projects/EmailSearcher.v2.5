Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Class MailExtractor


    Public Shared FileEmailRelation As New Dictionary(Of String, List(Of String))
    Public Shared EmailFileRelation As New Dictionary(Of String, List(Of String))

    Public Shared Sub ExtractEmails(ByVal inFilePath As String)
        Try
            Dim bufferLength As Integer = SizeBuffer(inFilePath)
            Dim buffer(bufferLength - 1) As Byte

            Dim emailRegex As New Regex(RegexString, RegexOptions.IgnoreCase)
            Dim domainExtensions As HashSet(Of String) = GetDomainExtensions()

            Dim emailsFound As New HashSet(Of String)()

            Dim stringBuilder As New StringBuilder()

            Using fs As FileStream = File.OpenRead(inFilePath)
                Dim bytesRead As Integer

                Do
                    bytesRead = fs.Read(buffer, 0, bufferLength)
                    If bytesRead <= 0 Then Exit Do

                    stringBuilder.Append(Encoding.UTF8.GetString(buffer, 0, bytesRead))
                    Dim data As String = stringBuilder.ToString()

                    Dim emailMatch As Match = emailRegex.Match(data)
                    While emailMatch.Success
                        Dim email As String = emailMatch.Value
                        If Not Form1.CheckBox1.Checked OrElse IsValidDomainEmail(email, domainExtensions) Then
                            emailsFound.Add(email)


                            UpdateUI(email, inFilePath)
                        End If
                        emailMatch = emailMatch.NextMatch()
                    End While

                    stringBuilder.Clear()


                    If Form1.stopSearch Then Exit Sub
                    Application.DoEvents()
                Loop
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Shared Sub UpdateUI(email As String, inFilePath As String)
        If Not Form1.ListBox2.Items.Contains(email) Then
            Form1.ListBox2.Items.Add(email)
            Form1.gCount += 1
        End If

        If Not Form1.ListBox1.Items.Contains(inFilePath) Then
            Form1.ListBox1.Items.Add(inFilePath)
        End If

        AddRelation(inFilePath, email, FileEmailRelation)
        AddRelation(email, inFilePath, EmailFileRelation)

        Form1.Label5.Text = "Emails found: " & Form1.ListBox2.Items.Count
        Form1.Label6.Text = "Files with emails: " & Form1.ListBox1.Items.Count
    End Sub

    Private Shared Function IsValidDomainEmail(email As String, domainExtensions As HashSet(Of String)) As Boolean
        Dim emailParts As String() = email.Split("."c)
        Dim emailExtension As String = emailParts.LastOrDefault()
        Return emailExtension IsNot Nothing AndAlso domainExtensions.Contains(emailExtension.ToLower())
    End Function
    Private Shared Sub AddRelation(inFilePath As String, email As String, relationDictionary As Dictionary(Of String, List(Of String)))
        If Not relationDictionary.ContainsKey(inFilePath) Then
            relationDictionary.Add(inFilePath, New List(Of String))
        End If
        relationDictionary(inFilePath).Add(email)
    End Sub

    Private Shared Function RegexString() As String
        Dim valor As Integer = Form1.TrackBar1.Value
        Dim v1 As String

        Select Case valor
            Case 0
                v1 = "{0,48}"
            Case 1
                v1 = "{1,38}"
            Case 2
                v1 = "{2,30}"
            Case 3
                v1 = "{3,24}"
            Case 4
                v1 = "{4,18}"
        End Select

        RegexString = $"\b\w{v1}([-+.]\w{v1})*@\w{v1}([-.]\w{v1})*\.\w{v1}([-.]\w{v1})*"

        Return RegexString

    End Function

    Private Shared Function GetDomainExtensions() As HashSet(Of String)
        Return New HashSet(Of String)(File.ReadAllLines(Form1.doms), StringComparer.OrdinalIgnoreCase)
    End Function
    Private Shared Function SizeBuffer(filePath As String) As Integer

        Dim fileInfo As New FileInfo(filePath)
        Dim fileSize As Long = fileInfo.Length

        Dim bufferSize As Integer = 4096

        Dim limiteBuffer As Integer = 1024 * 1024

        If fileSize < bufferSize Then

            bufferSize = CInt(fileSize)
        ElseIf bufferSize > limiteBuffer Then

            bufferSize = limiteBuffer
        End If

        Return bufferSize
    End Function



End Class

