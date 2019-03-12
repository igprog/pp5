Module PracenjeStanjaModul
    Sub PracenjeStanja()
        With Form1
            .TextBox14.Text = .TextBox1.Text   'Ime
            .TextBox15.Text = .TextBox2.Text    'Prezime
            .Label308.Text = .TextBox14.Text & " " & .TextBox15.Text  'Korisnik
            .TextBox76.Text = .ComboBox1.Text   'Dob
            .TextBox16.Text = .ComboBox2.Text    'Visina
            .TextBox17.Text = .ComboBox3.Text     'Masa
            .TextBox18.Text = .ComboBox6.Text       'Opseg struka
            .TextBox63.Text = .ComboBox7.Text         'Opseg bokova
        End With
    End Sub
End Module
