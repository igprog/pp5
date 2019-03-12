Module DijetaBindingSourceModul
    Sub Dijeta_BindingSource()
        With Form1
            Dim BS As BindingSource = .VrstaDijeteBindingSource

            .Label179.DataBindings.Add(New Binding("Text", BS, "BrojDijete"))
            .Label21.DataBindings.Add(New Binding("Text", BS, "NazivDijete"))
            .Label22.DataBindings.Add(New Binding("Text", BS, "NamjenaDijete"))
            .Label83.DataBindings.Add(New Binding("Text", BS, "UgljPost"))
            .Label84.DataBindings.Add(New Binding("Text", BS, "BjelPost"))
            .Label85.DataBindings.Add(New Binding("Text", BS, "MastiPost"))
            .Label88.DataBindings.Add(New Binding("Text", BS, "UgljikohidratiMinPost"))
            .Label89.DataBindings.Add(New Binding("Text", BS, "BjelancevineMinPost"))
            .Label90.DataBindings.Add(New Binding("Text", BS, "MastiMinPost"))
            .Label91.DataBindings.Add(New Binding("Text", BS, "UgljikohidratiMaxPost"))
            .Label92.DataBindings.Add(New Binding("Text", BS, "BjelancevineMaxPost"))
            .Label93.DataBindings.Add(New Binding("Text", BS, "MastiMaxPost"))
            .Label87.DataBindings.Add(New Binding("Text", BS, "Napomena"))
            .Label286.DataBindings.Add(New Binding("Text", BS, "MastiMaxZMK"))

        End With
    End Sub
End Module
