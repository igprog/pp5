Module BazaJelovnikaBindingSourceModul
    Sub BazaJelovnika_BindingSource()
        On Error Resume Next
        With Form1
            Dim BS As BindingSource = .BazaNazivaJelovnikaBindingSource
            Dim BS1 As BindingSource = .BazaJelovnikaBindingSource

            'Moji jelovnici
            If .RadioButton20.Checked = True Then
                BS = .BazaNazivaJelovnikaBindingSource
                BS1 = .BazaJelovnikaBindingSource
            End If
            'Primjeri jelovnika
            If .RadioButton21.Checked = True Then
                BS = .BazaNazivaPrimjeraJelovnikaBindingSource
                BS1 = .BazaPrimjeraJelovnikaBindingSource
            End If

            'Baza Naziva Jelovnika
            .Label180.DataBindings.Add(New Binding("Text", BS, "ID"))
            .Label181.DataBindings.Add(New Binding("Text", BS, "Korisnik"))
            .Label182.DataBindings.Add(New Binding("Text", BS, "BrojDijete"))
            .Label183.DataBindings.Add(New Binding("Text", BS, "NazivDijete"))
            .Label184.DataBindings.Add(New Binding("Text", BS, "DatumIzradeJelovnika"))
            .Label185.DataBindings.Add(New Binding("Text", BS, "DatumJelovnika"))
            .Label186.DataBindings.Add(New Binding("Text", BS, "NazivJelovnika"))
            .Label187.DataBindings.Add(New Binding("Text", BS, "EnergetskaVrijednostJelovnika_kcal"))

            'Baza Jelovnika
            .Label107.DataBindings.Add(New Binding("Text", BS1, "ID"))
            .Label108.DataBindings.Add(New Binding("Text", BS1, "Korisnik"))
            .Label109.DataBindings.Add(New Binding("Text", BS1, "BrojDijete"))
            .Label110.DataBindings.Add(New Binding("Text", BS1, "NazivDijete"))
            .Label111.DataBindings.Add(New Binding("Text", BS1, "DatumIzradeJelovnika"))
            .Label112.DataBindings.Add(New Binding("Text", BS1, "DatumJelovnika"))
            .Label113.DataBindings.Add(New Binding("Text", BS1, "NazivJelovnika"))
            .Label114.DataBindings.Add(New Binding("Text", BS1, "EnergetskaVrijednostJelovnika_kcal"))
            .Label115.DataBindings.Add(New Binding("Text", BS1, "Obrok"))
            .Label358.DataBindings.Add(New Binding("Text", BS1, "ObrokBroj"))
            .Label116.DataBindings.Add(New Binding("Text", BS1, "NazivPripremaJela"))
            .Label119.DataBindings.Add(New Binding("Text", BS1, "NazivNamirnice"))
            .Label359.DataBindings.Add(New Binding("Text", BS1, "TermickaObrada"))
            .Label120.DataBindings.Add(New Binding("Text", BS1, "Serviranja"))
            .Label121.DataBindings.Add(New Binding("Text", BS1, "Kolicina"))
            .Label122.DataBindings.Add(New Binding("Text", BS1, "Mjera"))
            .Label123.DataBindings.Add(New Binding("Text", BS1, "Masa_g"))
            .Label124.DataBindings.Add(New Binding("Text", BS1, "Energija_kcal"))
            .Label125.DataBindings.Add(New Binding("Text", BS1, "Energija_kJ"))
            .Label126.DataBindings.Add(New Binding("Text", BS1, "Ugljikohidrati_g"))
            .Label127.DataBindings.Add(New Binding("Text", BS1, "Bjelancevine_g"))
            .Label128.DataBindings.Add(New Binding("Text", BS1, "Masti_g"))
            .Label129.DataBindings.Add(New Binding("Text", BS1, "Skrob_g"))
            .Label130.DataBindings.Add(New Binding("Text", BS1, "UkupniSeceri_g"))
            .Label131.DataBindings.Add(New Binding("Text", BS1, "Glukoza_g"))
            .Label132.DataBindings.Add(New Binding("Text", BS1, "Fruktoza_g"))
            .Label133.DataBindings.Add(New Binding("Text", BS1, "Saharoza_g"))
            .Label134.DataBindings.Add(New Binding("Text", BS1, "Maltoza_g"))
            .Label135.DataBindings.Add(New Binding("Text", BS1, "Laktoza_g"))
            .Label136.DataBindings.Add(New Binding("Text", BS1, "Vlakna_g"))
            .Label137.DataBindings.Add(New Binding("Text", BS1, "ZasiceneMasti_g"))
            .Label138.DataBindings.Add(New Binding("Text", BS1, "JednostrukoNezasiceneMasti_g"))
            .Label139.DataBindings.Add(New Binding("Text", BS1, "VisestrukoNezasiceneMasti_g"))
            .Label140.DataBindings.Add(New Binding("Text", BS1, "TransMasneKiseline_g"))
            .Label141.DataBindings.Add(New Binding("Text", BS1, "Kolesterol_mg"))
            .Label142.DataBindings.Add(New Binding("Text", BS1, "Natrij_mg"))
            .Label143.DataBindings.Add(New Binding("Text", BS1, "Kalij_mg"))
            .Label144.DataBindings.Add(New Binding("Text", BS1, "Kalcij_mg"))
            .Label145.DataBindings.Add(New Binding("Text", BS1, "Magnezij_mg"))
            .Label146.DataBindings.Add(New Binding("Text", BS1, "Fosfor_mg"))
            .Label147.DataBindings.Add(New Binding("Text", BS1, "Zeljezo_mg"))
            .Label148.DataBindings.Add(New Binding("Text", BS1, "Bakar_mg"))
            .Label149.DataBindings.Add(New Binding("Text", BS1, "Cink_mg"))
            .Label150.DataBindings.Add(New Binding("Text", BS1, "Klor_mg"))
            .Label151.DataBindings.Add(New Binding("Text", BS1, "Mangan_mg"))
            .Label152.DataBindings.Add(New Binding("Text", BS1, "Selen_mikro_g"))
            .Label153.DataBindings.Add(New Binding("Text", BS1, "Jod_mikro_g"))
            .Label154.DataBindings.Add(New Binding("Text", BS1, "Retinol_mikro_g"))
            .Label155.DataBindings.Add(New Binding("Text", BS1, "Karoten_mikro_g"))
            .Label156.DataBindings.Add(New Binding("Text", BS1, "VitaminD_mikro_g"))
            .Label157.DataBindings.Add(New Binding("Text", BS1, "VitaminE_mg"))
            .Label158.DataBindings.Add(New Binding("Text", BS1, "VitaminB1_mg"))
            .Label159.DataBindings.Add(New Binding("Text", BS1, "VitaminB2_mg"))
            .Label160.DataBindings.Add(New Binding("Text", BS1, "VitaminB3_mg"))
            .Label161.DataBindings.Add(New Binding("Text", BS1, "VitaminB6_mg"))
            .Label162.DataBindings.Add(New Binding("Text", BS1, "VitaminB12_mikro_g"))
            .Label163.DataBindings.Add(New Binding("Text", BS1, "Folat_mikro_g"))
            .Label164.DataBindings.Add(New Binding("Text", BS1, "PantotenskaKiselina_mg"))
            .Label165.DataBindings.Add(New Binding("Text", BS1, "Biotin_mikro_g"))
            .Label166.DataBindings.Add(New Binding("Text", BS1, "VitaminC_mg"))
            .Label167.DataBindings.Add(New Binding("Text", BS1, "VitaminK_mikro_g"))
            .Label168.DataBindings.Add(New Binding("Text", BS1, "Zitarice"))
            .Label169.DataBindings.Add(New Binding("Text", BS1, "Povrce"))
            .Label170.DataBindings.Add(New Binding("Text", BS1, "Voce"))
            .Label171.DataBindings.Add(New Binding("Text", BS1, "Meso"))
            .Label172.DataBindings.Add(New Binding("Text", BS1, "Mlijeko"))
            .Label173.DataBindings.Add(New Binding("Text", BS1, "Masti"))
            .Label174.DataBindings.Add(New Binding("Text", BS1, "OstaleNamirnice"))
            .Label360.DataBindings.Add(New Binding("Text", BS1, "Cijena"))



        End With
    End Sub
End Module
