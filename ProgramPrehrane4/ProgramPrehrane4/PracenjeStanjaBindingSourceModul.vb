Module PracenjeStanjaBindingSourceModul
    Sub PracenjeStanja_BindingSource()
        On Error Resume Next
        With Form1
            Dim BS As BindingSource = .KorisniciPracenjeStanjaBindingSource

            'KorisniciPracenjeStanja
            .Label308.DataBindings.Add(New Binding("Text", BS, "Korisnik"))
            .Label219.DataBindings.Add(New Binding("Text", BS, "Ime"))
            .Label220.DataBindings.Add(New Binding("Text", BS, "Prezime"))
            .Label221.DataBindings.Add(New Binding("Text", BS, "Visina"))
            .Label307.DataBindings.Add(New Binding("Text", BS, "Dob"))
            .Label222.DataBindings.Add(New Binding("Text", BS, "Masa"))
            .Label223.DataBindings.Add(New Binding("Text", BS, "OpsegStruka"))
            .Label227.DataBindings.Add(New Binding("Text", BS, "OpsegBokova"))
            .Label302.DataBindings.Add(New Binding("Text", BS, "WHR"))
            .Label303.DataBindings.Add(New Binding("Text", BS, "BMI"))
            .Label304.DataBindings.Add(New Binding("Text", BS, "PrimjerenaTjelesnaMasaMin"))
            .Label305.DataBindings.Add(New Binding("Text", BS, "PrimjerenaTjelesnaMasaMax"))
            .Label224.DataBindings.Add(New Binding("Text", BS, "Datum"))

        End With
    End Sub
End Module
