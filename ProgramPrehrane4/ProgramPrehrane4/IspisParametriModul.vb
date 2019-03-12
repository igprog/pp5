Module IspisParametriModul
    Sub IspisParametri()
        On Error Resume Next
        Parametri()
        ParametriPreporuceneVrijednosti()
        ParametriProgressBar()
        NParametri()
        With Form1
            .TextBox69.Text = "UKUPNE VRIJEDNOSTI:" _
             & vbCrLf & "_______________________________________________________________________" _
                & vbCrLf & vbCrLf & "Klijent: " & .TextBox1.Text & " " & .TextBox2.Text _
                & vbCrLf & .Label21.Text & vbCrLf & "Jelovnik za " & .TextBox13.Text _
                & vbCrLf & "_______________________________________________________________________"

            .TextBox74.Text = ""
            Dim a As Integer
            For a = 0 To .ListBox14.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox14.Items(a) & " " & .ListBox15.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            For a = 0 To .ListBox17.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox17.Items(a) & " " & .ListBox41.Items(a) _
                    & " " & .ListBox22.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            For a = 0 To .ListBox18.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox18.Items(a) & " " & .ListBox42.Items(a) _
                    & " " & .ListBox23.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            For a = 0 To .ListBox20.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox20.Items(a) & " " & .ListBox43.Items(a) _
                    & " " & .ListBox21.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            'ostale namirnice
            For a = 0 To .ListBox34.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox34.Items(a) & " " & .ListBox35.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            For a = 0 To .ListBox45.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox45.Items(a) & " " & .ListBox47.Items(a) _
                    & " " & .ListBox46.Items(a)
            Next
            .TextBox74.Text = .TextBox74.Text & vbCrLf 'prazan red
            For a = 0 To .ListBox19.Items.Count - 1
                .TextBox74.Text = .TextBox74.Text & vbCrLf & .ListBox19.Items(a) & " " & .ListBox44.Items(a) _
                    & " " & .ListBox24.Items(a)
            Next

            .TextBox69.Text = .TextBox69.Text & vbCrLf & .TextBox74.Text  'UKUPNO
            .RichTextBoxPrintCtrl1.Text = .TextBox69.Text

        End With
    End Sub
End Module
