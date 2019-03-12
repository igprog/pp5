Module AktivnostiBindingSourceModul
    Sub AktivnostiBindingSource()
        With Form1

            .Label20.DataBindings.Clear()
            .Label94.DataBindings.Clear()
            .Label96.DataBindings.Clear()

            Dim BS As BindingSource = .SportskeAktivnostiBindingSource
            .Label20.DataBindings.Add(New Binding("Text", BS, "OpisTjelesneAktivnosti"))
            .Label96.DataBindings.Add(New Binding("Text", BS, "FaktorTjelesneAktivnostikJ"))
            .Label94.DataBindings.Add(New Binding("Text", BS, "FaktorTjelesneAktivnostiKcal"))

            Dim BS1 As BindingSource = .OdabranaDodatnaTjelesnaAktivnostBindingSource
            .Label98.DataBindings.Add(New Binding("Text", BS1, "OpisTjelesneAktivnosti"))
            .Label99.DataBindings.Add(New Binding("Text", BS1, "Minuta"))
            .Label100.DataBindings.Add(New Binding("Text", BS1, "EnergetskaPotrosnja_kcal"))

        End With
    End Sub
End Module
