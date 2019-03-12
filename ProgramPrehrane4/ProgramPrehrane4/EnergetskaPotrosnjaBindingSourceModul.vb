Module EnergetskaPotrosnjaBindingSourceModul
    Sub EnergetskaPotrosnjaBindingSource()
        With Form1
            .Label334.DataBindings.Clear()   'klijent
            .Label355.DataBindings.Clear()   'dan
            .Label346.DataBindings.Clear()   'tjelesna aktivnost
            .Label354.DataBindings.Clear()   'minuta (trajanje aktivnosti)
            .Label353.DataBindings.Clear()   'energetska potrosnja (kcal)

            Dim BS As BindingSource = .BazaEnergetskePotrosnjeBindingSource
            .Label334.DataBindings.Add(New Binding("Text", BS, "Korisnik"))
            .Label355.DataBindings.Add(New Binding("Text", BS, "Dan"))
            .Label346.DataBindings.Add(New Binding("Text", BS, "OpisTjelesneAktivnosti"))
            .Label354.DataBindings.Add(New Binding("Text", BS, "Minuta"))
            .Label353.DataBindings.Add(New Binding("Text", BS, "EnergetskaPotrosnja_kcal"))


        End With
    End Sub
End Module
