Module BazaNazivaJelovnikaSpremiModul
    Sub BazaNazivaJelovnikaSpremi()
        On Error Resume Next
        With Form1
            Dim BS As BindingSource = .BazaNazivaJelovnikaBindingSource
            If .DataGridView7.RowCount < 1 Then BS.AddNew()

            BS.MoveLast()

            .Label181.Text = .TextBox1.Text & " " & .TextBox2.Text  'Korisnik
            .Label182.Text = Val(.Label179.Text) 'BrojDijete
            .Label183.Text = .Label21.Text 'NazivDijete
            .Label184.Text = Date.Today             'DatumIzradeJelovnika
            .Label185.Text = .DateTimePicker1.Value.Date 'DatumJelovnika
            .Label186.Text = .TextBox13.Text  'NazivJelovnika
            .Label187.Text = Val(.Label175.Text)  'Energetska vrijednost jelovnika

            BS.AddNew()

        End With
    End Sub
End Module
