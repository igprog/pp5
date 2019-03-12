﻿Module MojeNamirniceBindingSourceModul
    Sub MojeNamirnice_BindingSource()
        With Form1
            Dim BS As BindingSource = .MojeNamirniceBindingSource
            'Baza Naziva Jelovnika
            .TextBox19.DataBindings.Add(New Binding("Text", BS, "NazivNamirnice"))
            .TextBox20.DataBindings.Add(New Binding("Text", BS, "Kolicina"))
            .ComboBox8.DataBindings.Add(New Binding("Text", BS, "Mjera"))
            .TextBox21.DataBindings.Add(New Binding("Text", BS, "Masa_g"))
            .TextBox22.DataBindings.Add(New Binding("Text", BS, "Energija_kcal"))
            .Label26.DataBindings.Add(New Binding("Text", BS, "Energija_kJ"))
            .TextBox23.DataBindings.Add(New Binding("Text", BS, "Ugljikohidrati_g"))
            .TextBox24.DataBindings.Add(New Binding("Text", BS, "Bjelancevine_g"))
            .TextBox25.DataBindings.Add(New Binding("Text", BS, "Masti_g"))
            .TextBox26.DataBindings.Add(New Binding("Text", BS, "Skrob_g"))
            .TextBox27.DataBindings.Add(New Binding("Text", BS, "UkupniSeceri_g"))
            .TextBox28.DataBindings.Add(New Binding("Text", BS, "Glukoza_g"))
            .TextBox29.DataBindings.Add(New Binding("Text", BS, "Fruktoza_g"))
            .TextBox30.DataBindings.Add(New Binding("Text", BS, "Saharoza_g"))
            .TextBox31.DataBindings.Add(New Binding("Text", BS, "Maltoza_g"))
            .TextBox32.DataBindings.Add(New Binding("Text", BS, "Laktoza_g"))
            .TextBox33.DataBindings.Add(New Binding("Text", BS, "Vlakna_g"))
            .TextBox34.DataBindings.Add(New Binding("Text", BS, "ZasiceneMasti_g"))
            .TextBox35.DataBindings.Add(New Binding("Text", BS, "JednostrukoNezasiceneMasti_g"))
            .TextBox36.DataBindings.Add(New Binding("Text", BS, "VisestrukoNezasiceneMasti_g"))
            .TextBox37.DataBindings.Add(New Binding("Text", BS, "TransMasneKiseline_g"))
            .TextBox38.DataBindings.Add(New Binding("Text", BS, "Kolesterol_mg"))
            .TextBox39.DataBindings.Add(New Binding("Text", BS, "Natrij_mg"))
            .TextBox40.DataBindings.Add(New Binding("Text", BS, "Kalij_mg"))
            .TextBox41.DataBindings.Add(New Binding("Text", BS, "Kalcij_mg"))
            .TextBox42.DataBindings.Add(New Binding("Text", BS, "Magnezij_mg"))
            .TextBox43.DataBindings.Add(New Binding("Text", BS, "Fosfor_mg"))
            .TextBox44.DataBindings.Add(New Binding("Text", BS, "Zeljezo_mg"))
            .TextBox45.DataBindings.Add(New Binding("Text", BS, "Bakar_mg"))
            .TextBox46.DataBindings.Add(New Binding("Text", BS, "Cink_mg"))
            .TextBox47.DataBindings.Add(New Binding("Text", BS, "Klor_mg"))
            .TextBox48.DataBindings.Add(New Binding("Text", BS, "Mangan_mg"))
            .TextBox49.DataBindings.Add(New Binding("Text", BS, "Selen_mikro_g"))
            .TextBox50.DataBindings.Add(New Binding("Text", BS, "Jod_mikro_g"))
            .TextBox51.DataBindings.Add(New Binding("Text", BS, "Retinol_mikro_g"))
            .TextBox52.DataBindings.Add(New Binding("Text", BS, "Karoten_mikro_g"))
            .TextBox53.DataBindings.Add(New Binding("Text", BS, "VitaminD_mikro_g"))
            .TextBox54.DataBindings.Add(New Binding("Text", BS, "VitaminE_mg"))
            .TextBox55.DataBindings.Add(New Binding("Text", BS, "VitaminB1_mg"))
            .TextBox56.DataBindings.Add(New Binding("Text", BS, "VitaminB2_mg"))
            .TextBox57.DataBindings.Add(New Binding("Text", BS, "VitaminB3_mg"))
            .TextBox58.DataBindings.Add(New Binding("Text", BS, "VitaminB6_mg"))
            .TextBox59.DataBindings.Add(New Binding("Text", BS, "VitaminB12_mikro_g"))
            .TextBox60.DataBindings.Add(New Binding("Text", BS, "Folat_mikro_g"))
            .TextBox61.DataBindings.Add(New Binding("Text", BS, "PantotenskaKiselina_mg"))
            .TextBox62.DataBindings.Add(New Binding("Text", BS, "Biotin_mikro_g"))
            .TextBox71.DataBindings.Add(New Binding("Text", BS, "VitaminC_mg"))
            .TextBox72.DataBindings.Add(New Binding("Text", BS, "VitaminK_mikro_g"))
            .Label367.DataBindings.Add(New Binding("Text", BS, "Zitarice"))
            .Label368.DataBindings.Add(New Binding("Text", BS, "Povrce"))
            .Label369.DataBindings.Add(New Binding("Text", BS, "Voce"))
            .Label370.DataBindings.Add(New Binding("Text", BS, "Meso"))
            .Label371.DataBindings.Add(New Binding("Text", BS, "Mlijeko"))
            .Label372.DataBindings.Add(New Binding("Text", BS, "Masti"))
            .Label373.DataBindings.Add(New Binding("Text", BS, "OstaleNamirnice"))

            .Label374.DataBindings.Add(New Binding("Text", BS, "SkupinaNamirnica"))
            .Label375.DataBindings.Add(New Binding("Text", BS, "SkupinaNamirnicaGubiciVitamina"))

        End With
    End Sub
End Module