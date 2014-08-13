Public Class Main_Screen

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If My.Computer.FileSystem.FileExists((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\")) = True Then
                My.Computer.Audio.Play((Application.StartupPath & "\Sounds\UHOH.WAV").Replace("\\", "\"), AudioPlayMode.Background)
            End If
            Dim Display_Message1 As New Display_Message()
            Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString
            Display_Message1.Timer1.Interval = 1000
            Display_Message1.ShowDialog()
            Display_Message1.Dispose()
            Display_Message1 = Nothing
            If My.Computer.FileSystem.DirectoryExists((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs") = False Then
                My.Computer.FileSystem.CreateDirectory((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            End If
            Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
            filewriter.Flush()
            filewriter.Close()
            filewriter = Nothing
            ex = Nothing
            identifier_msg = Nothing
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

        
            If Textbox1.Text.Length > 0 Then
                Label1.Text = "Checking spelling..."
                ' Make a Word server object.
                Dim word_server As New Word.Application

                ' Hide the server.
                word_server.Visible = False

                ' Make a Word Document.
                Dim doc As Word.Document = _
                    word_server.Documents.Add()
                Dim rng As Word.Range

                ' Make a Range to represent the Document.
                rng = doc.Range()

                ' Copy the text into the Document.
                rng.Text = Textbox1.Text

                ' Activate the Document and call its CheckSpelling
                ' method.
                doc.Activate()
                doc.CheckSpelling()

                ' Copy the results back into the TextBox,
                ' trimming off trailing CR and LF characters.
                Dim chars() As Char = {CType(vbCr, Char), _
                    CType(vbLf, Char)}
                Textbox1.Text = doc.Range().Text.Trim(chars)

                ' Close the Document, not saving changes.
                doc.Close(SaveChanges:=False)

                ' Close the Word server.
                word_server.Quit()
                Label1.Text = "Spell check complete"
            End If
        Catch ex As Exception
            Error_Handler(ex, "Spell Check")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If Textbox1.Text.Length > 0 Then
                Label1.Text = "Checking grammar..."
                ' Make a Word server object.
                Dim word_server As New Word.Application

                ' Hide the server.
                word_server.Visible = False

                ' Make a Word Document.
                Dim doc As Word.Document = _
                    word_server.Documents.Add()
                Dim rng As Word.Range

                ' Make a Range to represent the Document.
                rng = doc.Range()

                ' Copy the text into the Document.
                rng.Text = Textbox1.Text

                ' Activate the Document and call its CheckSpelling
                ' method.
                doc.Activate()
                doc.CheckGrammar()

                ' Copy the results back into the TextBox,
                ' trimming off trailing CR and LF characters.
                Dim chars() As Char = {CType(vbCr, Char), _
                    CType(vbLf, Char)}
                Textbox1.Text = doc.Range().Text.Trim(chars)

                ' Close the Document, not saving changes.
                doc.Close(SaveChanges:=False)

                ' Close the Word server.
                word_server.Quit()
                Label1.Text = "Grammar check complete"
            End If
        Catch ex As Exception
            Error_Handler(ex, "Spell Check")
        End Try
    End Sub
End Class
