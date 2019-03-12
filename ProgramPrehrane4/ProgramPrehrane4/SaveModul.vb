Module SaveModul
    Sub SavePrikazi()
        With Form1
            On Error GoTo err
            ' .TextBox69.Text = .RichTextBoxPrintCtrl1.Text
            '         .SaveFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            '        Dim result As System.Nullable(Of Boolean) = .SaveFileDialog1.ShowDialog()
            '       If result = True Then
            'Dim fileStream As System.IO.Stream = .SaveFileDialog1.OpenFile()
            'Dim sw As New System.IO.StreamWriter(fileStream)
            'sw.WriteLine(.TextBox69.Text)
            '' sw.WriteLine(.RichTextBoxPrintCtrl1.Text)
            'sw.Flush()
            'sw.Close()
            'End If

            ' Create a SaveFileDialog to request a path and file name to save to. 
            Dim saveFile1 As New SaveFileDialog()

            ' Initialize the SaveFileDialog to specify the RTF extention for the file.
            saveFile1.DefaultExt = "*.rtf"
            saveFile1.Filter = "RTF Files|*.rtf"

            ' Determine whether the user selected a file name from the saveFileDialog. 
            If (saveFile1.ShowDialog() = System.Windows.Forms.DialogResult.OK) _
                And (saveFile1.FileName.Length > 0) Then

                ' Save the contents of the RichTextBox into the file.
                .RichTextBoxPrintCtrl1.SaveFile(saveFile1.FileName)
            End If


            Exit Sub
err:
            MsgBox(ErrorToString)
        End With
    End Sub
End Module
