Module UzmiBazaJelovnikaModul
    Sub UzmiBazaJelovnika()
        On Error Resume Next
        With Form1
            ' Dim K As Integer = .ComboBox10.Text    'broj korisnika jelovnika
            .TabControl1.SelectedIndex = 5   'Izrada jelovnika
            NoviJelovnik()
            .Label27.Visible = True
            .Label31.Visible = True
            .Label32.Visible = True
            .Label33.Visible = True
            .Label34.Visible = True
            .Label35.Visible = True
            .Label36.Visible = True
            .Label37.Visible = True
            .Label38.Visible = True
            .Label39.Visible = True
            .Label40.Visible = True
            .Label41.Visible = True
            .Label42.Visible = True
            .Label43.Visible = True
            .Label44.Visible = True
            .Label45.Visible = True
            .Label46.Visible = True
            .Label47.Visible = True
            .Label48.Visible = True
            .Label49.Visible = True
            .Label50.Visible = True
            .Label51.Visible = True
            .Label52.Visible = True
            .Label53.Visible = True
            .Label54.Visible = True
            .Label55.Visible = True
            .Label56.Visible = True
            .Label57.Visible = True
            .Label58.Visible = True
            .Label59.Visible = True
            .Label60.Visible = True
            .Label61.Visible = True
            .Label62.Visible = True
            .Label63.Visible = True
            .Label64.Visible = True
            .Label65.Visible = True
            .Label66.Visible = True
            .Label67.Visible = True
            .Label68.Visible = True
            .Label69.Visible = True
            .Label70.Visible = True
            .Label71.Visible = True
            .Label72.Visible = True
            .Label73.Visible = True
            .Label74.Visible = True
            .Label75.Visible = True
            .Label76.Visible = True
            .Label77.Visible = True
            .Label78.Visible = True
            .Label79.Visible = True
            .Label80.Visible = True
            .Label81.Visible = True
            .Label82.Visible = True
            .Label104.Visible = True
            .Label105.Visible = True
            .Label106.Visible = True
            .Label228.Visible = True
            .Label229.Visible = True
            .Label230.Visible = True

            .Label358.Visible = True
            .Label359.Visible = True
            .Label360.Visible = True


            Dim DGV As DataGridView
            DGV = .DataGridView8
            Dim BS As BindingSource

            .TextBox13.Text = .DataGridView8.Rows(0).Cells(6).Value    'naziv jelovnika

            For i = 0 To DGV.RowCount - 1
                If (DGV.Rows(i).Cells(9).Value) = 1 Then   'doručak
                    BS = .DorucakBindingSource
                    .TextBox11.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If
                If (DGV.Rows(i).Cells(9).Value) = 2 Then   'jutarnja uzina
                    BS = .JutarnjaUzinaBindingSource
                    .TextBox64.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If
                If (DGV.Rows(i).Cells(9).Value) = 3 Then    'rucak
                    BS = .RucakBindingSource
                    .TextBox65.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If
                If (DGV.Rows(i).Cells(9).Value) = 4 Then    'popodnevna uzi
                    BS = .PopodnevnaUzinaBindingSource
                    .TextBox66.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If
                If (DGV.Rows(i).Cells(9).Value) = 5 Then   'vecera
                    BS = .VeceraBindingSource
                    .TextBox67.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If
                If (DGV.Rows(i).Cells(9).Value) = 6 Then    'obrok pred spavanje
                    BS = .ObrokPredSpavanjeBindingSource
                    .TextBox68.Text = (DGV.Rows(i).Cells(10).Value)  'NacinPripremaJela
                End If

                '   .TextBox13.DataBindings.Clear()     'naziv jelovnika
                '  .TextBox13.DataBindings.Add(New Binding("Text", BS, "NazivJelovnika"))

                .Label27.DataBindings.Clear()
                .Label27.DataBindings.Add(New Binding("Text", BS, "NazivNamirnice"))
                .Label27.Text = .DataGridView8.Rows(i).Cells(11).Value.ToString

                .Label359.DataBindings.Clear()
                .Label359.DataBindings.Add(New Binding("Text", BS, "TermickaObrada"))
                .Label359.Text = .DataGridView8.Rows(i).Cells(12).Value.ToString

                .Label106.DataBindings.Clear()
                .Label106.DataBindings.Add(New Binding("Text", BS, "Serviranja"))
                .Label106.Text = .DataGridView8.Rows(i).Cells(13).Value.ToString

                .Label31.DataBindings.Clear()
                .Label31.DataBindings.Add(New Binding("Text", BS, "Kolicina"))
                .Label31.Text = .DataGridView8.Rows(i).Cells(14).Value.ToString

                .Label32.DataBindings.Clear()
                .Label32.DataBindings.Add(New Binding("Text", BS, "Mjera"))
                .Label32.Text = .DataGridView8.Rows(i).Cells(15).Value.ToString

                .Label33.DataBindings.Clear()
                .Label33.DataBindings.Add(New Binding("Text", BS, "Masa_g"))
                .Label33.Text = .DataGridView8.Rows(i).Cells(16).Value.ToString

                .Label34.DataBindings.Clear()
                .Label34.DataBindings.Add(New Binding("Text", BS, "Energija_kcal"))
                .Label34.Text = .DataGridView8.Rows(i).Cells(17).Value.ToString

                .Label35.DataBindings.Clear()
                .Label35.DataBindings.Add(New Binding("Text", BS, "Energija_kJ"))
                .Label35.Text = .DataGridView8.Rows(i).Cells(18).Value.ToString

                .Label36.DataBindings.Clear()
                .Label36.DataBindings.Add(New Binding("Text", BS, "Ugljikohidrati_g"))
                .Label36.Text = .DataGridView8.Rows(i).Cells(19).Value.ToString

                .Label37.DataBindings.Clear()
                .Label37.DataBindings.Add(New Binding("Text", BS, "Bjelancevine_g"))
                .Label37.Text = .DataGridView8.Rows(i).Cells(20).Value.ToString

                .Label38.DataBindings.Clear()
                .Label38.DataBindings.Add(New Binding("Text", BS, "Masti_g"))
                .Label38.Text = .DataGridView8.Rows(i).Cells(21).Value.ToString

                .Label39.DataBindings.Clear()
                .Label39.DataBindings.Add(New Binding("Text", BS, "Skrob_g"))
                .Label39.Text = .DataGridView8.Rows(i).Cells(22).Value.ToString

                .Label40.DataBindings.Clear()
                .Label40.DataBindings.Add(New Binding("Text", BS, "UkupniSeceri_g"))
                .Label40.Text = .DataGridView8.Rows(i).Cells(23).Value.ToString

                .Label41.DataBindings.Clear()
                .Label41.DataBindings.Add(New Binding("Text", BS, "Glukoza_g"))
                .Label41.Text = .DataGridView8.Rows(i).Cells(24).Value.ToString

                .Label42.DataBindings.Clear()
                .Label42.DataBindings.Add(New Binding("Text", BS, "Fruktoza_g"))
                .Label42.Text = .DataGridView8.Rows(i).Cells(25).Value.ToString

                .Label43.DataBindings.Clear()
                .Label43.DataBindings.Add(New Binding("Text", BS, "Saharoza_g"))
                .Label43.Text = .DataGridView8.Rows(i).Cells(26).Value.ToString

                .Label44.DataBindings.Clear()
                .Label44.DataBindings.Add(New Binding("Text", BS, "Maltoza_g"))
                .Label44.Text = .DataGridView8.Rows(i).Cells(27).Value.ToString

                .Label45.DataBindings.Clear()
                .Label45.DataBindings.Add(New Binding("Text", BS, "Laktoza_g"))
                .Label45.Text = .DataGridView8.Rows(i).Cells(28).Value.ToString

                .Label46.DataBindings.Clear()
                .Label46.DataBindings.Add(New Binding("Text", BS, "Vlakna_g"))
                .Label46.Text = .DataGridView8.Rows(i).Cells(29).Value.ToString

                .Label47.DataBindings.Clear()
                .Label47.DataBindings.Add(New Binding("Text", BS, "ZasiceneMasti_g"))
                .Label47.Text = .DataGridView8.Rows(i).Cells(30).Value.ToString

                .Label48.DataBindings.Clear()
                .Label48.DataBindings.Add(New Binding("Text", BS, "JednostrukoNezasiceneMasti_g"))
                .Label48.Text = .DataGridView8.Rows(i).Cells(31).Value.ToString

                .Label49.DataBindings.Clear()
                .Label49.DataBindings.Add(New Binding("Text", BS, "VisestrukoNezasiceneMasti_g"))
                .Label49.Text = .DataGridView8.Rows(i).Cells(32).Value.ToString

                .Label50.DataBindings.Clear()
                .Label50.DataBindings.Add(New Binding("Text", BS, "TransMasneKiseline_g"))
                .Label50.Text = .DataGridView8.Rows(i).Cells(33).Value.ToString

                .Label51.DataBindings.Clear()
                .Label51.DataBindings.Add(New Binding("Text", BS, "Kolesterol_mg"))
                .Label51.Text = .DataGridView8.Rows(i).Cells(34).Value.ToString

                .Label52.DataBindings.Clear()
                .Label52.DataBindings.Add(New Binding("Text", BS, "Natrij_mg"))
                .Label52.Text = .DataGridView8.Rows(i).Cells(35).Value.ToString

                .Label53.DataBindings.Clear()
                .Label53.DataBindings.Add(New Binding("Text", BS, "Kalij_mg"))
                .Label53.Text = .DataGridView8.Rows(i).Cells(36).Value.ToString

                .Label54.DataBindings.Clear()
                .Label54.DataBindings.Add(New Binding("Text", BS, "Kalcij_mg"))
                .Label54.Text = .DataGridView8.Rows(i).Cells(37).Value.ToString

                .Label55.DataBindings.Clear()
                .Label55.DataBindings.Add(New Binding("Text", BS, "Magnezij_mg"))
                .Label55.Text = .DataGridView8.Rows(i).Cells(38).Value.ToString

                .Label56.DataBindings.Clear()
                .Label56.DataBindings.Add(New Binding("Text", BS, "Fosfor_mg"))
                .Label56.Text = .DataGridView8.Rows(i).Cells(39).Value.ToString

                .Label57.DataBindings.Clear()
                .Label57.DataBindings.Add(New Binding("Text", BS, "Zeljezo_mg"))
                .Label57.Text = .DataGridView8.Rows(i).Cells(40).Value.ToString

                .Label58.DataBindings.Clear()
                .Label58.DataBindings.Add(New Binding("Text", BS, "Bakar_mg"))
                .Label58.Text = .DataGridView8.Rows(i).Cells(41).Value.ToString

                .Label59.DataBindings.Clear()
                .Label59.DataBindings.Add(New Binding("Text", BS, "Cink_mg"))
                .Label59.Text = .DataGridView8.Rows(i).Cells(42).Value.ToString

                .Label60.DataBindings.Clear()
                .Label60.DataBindings.Add(New Binding("Text", BS, "Klor_mg"))
                .Label60.Text = .DataGridView8.Rows(i).Cells(43).Value.ToString

                .Label61.DataBindings.Clear()
                .Label61.DataBindings.Add(New Binding("Text", BS, "Mangan_mg"))
                .Label61.Text = .DataGridView8.Rows(i).Cells(44).Value.ToString

                .Label62.DataBindings.Clear()
                .Label62.DataBindings.Add(New Binding("Text", BS, "Selen_mikro_g"))
                .Label62.Text = .DataGridView8.Rows(i).Cells(45).Value.ToString

                .Label63.DataBindings.Clear()
                .Label63.DataBindings.Add(New Binding("Text", BS, "Jod_mikro_g"))
                .Label63.Text = .DataGridView8.Rows(i).Cells(46).Value.ToString

                .Label64.DataBindings.Clear()
                .Label64.DataBindings.Add(New Binding("Text", BS, "Retinol_mikro_g"))
                .Label64.Text = .DataGridView8.Rows(i).Cells(47).Value.ToString

                .Label65.DataBindings.Clear()
                .Label65.DataBindings.Add(New Binding("Text", BS, "Karoten_mikro_g"))
                .Label65.Text = .DataGridView8.Rows(i).Cells(48).Value.ToString

                .Label66.DataBindings.Clear()
                .Label66.DataBindings.Add(New Binding("Text", BS, "VitaminD_mikro_g"))
                .Label66.Text = .DataGridView8.Rows(i).Cells(49).Value.ToString

                .Label67.DataBindings.Clear()
                .Label67.DataBindings.Add(New Binding("Text", BS, "VitaminE_mg"))
                .Label67.Text = .DataGridView8.Rows(i).Cells(50).Value.ToString

                .Label68.DataBindings.Clear()
                .Label68.DataBindings.Add(New Binding("Text", BS, "VitaminB1_mg"))
                .Label68.Text = .DataGridView8.Rows(i).Cells(51).Value.ToString

                .Label69.DataBindings.Clear()
                .Label69.DataBindings.Add(New Binding("Text", BS, "VitaminB2_mg"))
                .Label69.Text = .DataGridView8.Rows(i).Cells(52).Value.ToString

                .Label70.DataBindings.Clear()
                .Label70.DataBindings.Add(New Binding("Text", BS, "VitaminB3_mg"))
                .Label70.Text = .DataGridView8.Rows(i).Cells(53).Value.ToString

                .Label71.DataBindings.Clear()
                .Label71.DataBindings.Add(New Binding("Text", BS, "VitaminB6_mg"))
                .Label71.Text = .DataGridView8.Rows(i).Cells(54).Value.ToString

                .Label72.DataBindings.Clear()
                .Label72.DataBindings.Add(New Binding("Text", BS, "VitaminB12_mikro_g"))
                .Label72.Text = .DataGridView8.Rows(i).Cells(55).Value.ToString

                .Label73.DataBindings.Clear()
                .Label73.DataBindings.Add(New Binding("Text", BS, "Folat_mikro_g"))
                .Label73.Text = .DataGridView8.Rows(i).Cells(56).Value.ToString

                .Label74.DataBindings.Clear()
                .Label74.DataBindings.Add(New Binding("Text", BS, "PantotenskaKiselina_mg"))
                .Label74.Text = .DataGridView8.Rows(i).Cells(57).Value.ToString

                .Label75.DataBindings.Clear()
                .Label75.DataBindings.Add(New Binding("Text", BS, "Biotin_mikro_g"))
                .Label75.Text = .DataGridView8.Rows(i).Cells(58).Value.ToString

                .Label76.DataBindings.Clear()
                .Label76.DataBindings.Add(New Binding("Text", BS, "VitaminC_mg"))
                .Label76.Text = .DataGridView8.Rows(i).Cells(59).Value.ToString

                .Label77.DataBindings.Clear()
                .Label77.DataBindings.Add(New Binding("Text", BS, "VitaminK_mikro_g"))
                .Label77.Text = .DataGridView8.Rows(i).Cells(60).Value.ToString

                .Label78.DataBindings.Clear()
                .Label78.DataBindings.Add(New Binding("Text", BS, "Zitarice"))
                .Label78.Text = .DataGridView8.Rows(i).Cells(61).Value.ToString

                .Label79.DataBindings.Clear()
                .Label79.DataBindings.Add(New Binding("Text", BS, "Povrce"))
                .Label79.Text = .DataGridView8.Rows(i).Cells(62).Value.ToString

                .Label80.DataBindings.Clear()
                .Label80.DataBindings.Add(New Binding("Text", BS, "Voce"))
                .Label80.Text = .DataGridView8.Rows(i).Cells(63).Value.ToString

                .Label81.DataBindings.Clear()
                .Label81.DataBindings.Add(New Binding("Text", BS, "Meso"))
                .Label81.Text = .DataGridView8.Rows(i).Cells(64).Value.ToString

                .Label82.DataBindings.Clear()
                .Label82.DataBindings.Add(New Binding("Text", BS, "Mlijeko"))
                .Label82.Text = .DataGridView8.Rows(i).Cells(65).Value.ToString

                .Label104.DataBindings.Clear()
                .Label104.DataBindings.Add(New Binding("Text", BS, "Masti"))
                .Label104.Text = .DataGridView8.Rows(i).Cells(66).Value.ToString

                .Label105.DataBindings.Clear()
                .Label105.DataBindings.Add(New Binding("Text", BS, "OstaleNamirnice"))
                .Label105.Text = .DataGridView8.Rows(i).Cells(67).Value.ToString

                .Label360.DataBindings.Clear()
                .Label360.DataBindings.Add(New Binding("Text", BS, "Cijena"))
                .Label360.Text = .DataGridView8.Rows(i).Cells(68).Value.ToString


                If .Label27.Text <> "" Then
                    BS.AddNew()
                End If

            Next i

            .Label27.Visible = False
            .Label31.Visible = False
            .Label32.Visible = False
            .Label33.Visible = False
            .Label34.Visible = False
            .Label35.Visible = False
            .Label36.Visible = False
            .Label37.Visible = False
            .Label38.Visible = False
            .Label39.Visible = False
            .Label40.Visible = False
            .Label41.Visible = False
            .Label42.Visible = False
            .Label43.Visible = False
            .Label44.Visible = False
            .Label45.Visible = False
            .Label46.Visible = False
            .Label47.Visible = False
            .Label48.Visible = False
            .Label49.Visible = False
            .Label50.Visible = False
            .Label51.Visible = False
            .Label52.Visible = False
            .Label53.Visible = False
            .Label54.Visible = False
            .Label55.Visible = False
            .Label56.Visible = False
            .Label57.Visible = False
            .Label58.Visible = False
            .Label59.Visible = False
            .Label60.Visible = False
            .Label61.Visible = False
            .Label62.Visible = False
            .Label63.Visible = False
            .Label64.Visible = False
            .Label65.Visible = False
            .Label66.Visible = False
            .Label67.Visible = False
            .Label68.Visible = False
            .Label69.Visible = False
            .Label70.Visible = False
            .Label71.Visible = False
            .Label72.Visible = False
            .Label73.Visible = False
            .Label74.Visible = False
            .Label75.Visible = False
            .Label76.Visible = False
            .Label77.Visible = False
            .Label78.Visible = False
            .Label79.Visible = False
            .Label80.Visible = False
            .Label81.Visible = False
            .Label82.Visible = False
            .Label104.Visible = False
            .Label105.Visible = False
            .Label106.Visible = False
            .Label228.Visible = False
            .Label229.Visible = False
            .Label230.Visible = False

            .Label358.Visible = False
            .Label359.Visible = False
            .Label360.Visible = False

            If Val(.TextBox3.Text) > 500 And Val(.TextBox3.Text) < 20000 Then
                UkupneVrijednosti()
                ObrociNutrijentiUkupno()
            Else
                MsgBox("Prije izrade jelovnika izračunajte preporučeni energetski unos")
                .TabControl1.SelectedIndex = 0
                .TabControl3.SelectedIndex = 0
            End If

            'graf - pita
            If .ListBox10.Items(0) = "0 kcal" Then
                .Chart2.Series(0).Points.Clear()
                .Chart2.Series(0).Points.AddY(20)
                .Chart2.Series(0).Points.AddY(30)
                .Chart2.Series(0).Points.AddY(50)
            End If

        End With

    End Sub
End Module
