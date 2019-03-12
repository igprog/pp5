Module BazaKorisnikaBindingSourceModul
    Sub BazaKorisnika_BindingSource()
        With Form1
            Dim BS As BindingSource = .BazaKorisnikaBindingSource

            'Korisnici
            .Label210.DataBindings.Add(New Binding("Text", BS, "Korisnik"))
            .Label211.DataBindings.Add(New Binding("Text", BS, "Ime"))
            .Label212.DataBindings.Add(New Binding("Text", BS, "Prezime"))
            .Label213.DataBindings.Add(New Binding("Text", BS, "Dob"))
            .Label178.DataBindings.Add(New Binding("Text", BS, "Spol"))
            .Label250.DataBindings.Add(New Binding("Text", BS, "Visina"))
            .Label251.DataBindings.Add(New Binding("Text", BS, "Masa"))
            .Label252.DataBindings.Add(New Binding("Text", BS, "OpsegStruka"))
            .Label253.DataBindings.Add(New Binding("Text", BS, "OpsegBokova"))
            .Label289.DataBindings.Add(New Binding("Text", BS, "WHR"))
            .Label290.DataBindings.Add(New Binding("Text", BS, "IntenzitetAktivnostiNaPoslu"))
            .Label291.DataBindings.Add(New Binding("Text", BS, "IntenzitetAktivnostiIzvanPosla"))
            .Label292.DataBindings.Add(New Binding("Text", BS, "IntenzitetAktivnostiDjeca"))
            .Label293.DataBindings.Add(New Binding("Text", BS, "BMI"))
            .Label294.DataBindings.Add(New Binding("Text", BS, "TEE"))
            .Label295.DataBindings.Add(New Binding("Text", BS, "DodatnaEnergetskaPotrosnja"))
            .Label296.DataBindings.Add(New Binding("Text", BS, "Napomena"))
            .Label297.DataBindings.Add(New Binding("Text", BS, "Datum"))



        End With
    End Sub
End Module
