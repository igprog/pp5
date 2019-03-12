Module FavoritiSpremiModul
    Sub FavoritiSpremi()
        On Error Resume Next
        If Form1.CheckBox1.Checked = True Then

            With Form1

                Dim Namirnica As TextBox = .TextBox5
                If Namirnica.Text = "" Then
                    MsgBox("Odaberite namirnicu.")

                    Exit Sub
                End If
                'Provjera dali vec postoji namirnica u favoritima
                Dim DGV As DataGridView = .DataGridView15
                Dim j As Integer
                For j = 0 To DGV.RowCount - 1
                    If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                        If DGV.Rows(j).Cells(1).Value = Namirnica.Text Then
                            '   MsgBox("Namirnica " & Namirnica.Text & " već postoji u Favoritima.")
                            ' .TabControl1.SelectedIndex = 4   'vrati u izradu jelovnika
                            Exit Sub
                        End If
                    End If
                Next j
            End With


            With Form1
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
                .Label364.Visible = True
                .Label365.Visible = True

                'cijene-favoriti-izrada jelovnika form
                .Label329.Visible = True
                .Label319.Visible = True
                .Label331.Visible = True

                'cijene
                .Label327.Visible = True
                .Label325.Visible = True
                .Label326.Visible = True


                .Label27.DataBindings.Clear()
                .Label31.DataBindings.Clear()
                .Label32.DataBindings.Clear()
                .Label33.DataBindings.Clear()
                .Label34.DataBindings.Clear()
                .Label35.DataBindings.Clear()
                .Label36.DataBindings.Clear()
                .Label37.DataBindings.Clear()
                .Label38.DataBindings.Clear()
                .Label39.DataBindings.Clear()
                .Label40.DataBindings.Clear()
                .Label41.DataBindings.Clear()
                .Label42.DataBindings.Clear()
                .Label43.DataBindings.Clear()
                .Label44.DataBindings.Clear()
                .Label45.DataBindings.Clear()
                .Label46.DataBindings.Clear()
                .Label47.DataBindings.Clear()
                .Label48.DataBindings.Clear()
                .Label49.DataBindings.Clear()
                .Label50.DataBindings.Clear()
                .Label51.DataBindings.Clear()
                .Label52.DataBindings.Clear()
                .Label53.DataBindings.Clear()
                .Label54.DataBindings.Clear()
                .Label55.DataBindings.Clear()
                .Label56.DataBindings.Clear()
                .Label57.DataBindings.Clear()
                .Label58.DataBindings.Clear()
                .Label59.DataBindings.Clear()
                .Label60.DataBindings.Clear()
                .Label61.DataBindings.Clear()
                .Label62.DataBindings.Clear()
                .Label63.DataBindings.Clear()
                .Label64.DataBindings.Clear()
                .Label65.DataBindings.Clear()
                .Label66.DataBindings.Clear()
                .Label67.DataBindings.Clear()
                .Label68.DataBindings.Clear()
                .Label69.DataBindings.Clear()
                .Label70.DataBindings.Clear()
                .Label71.DataBindings.Clear()
                .Label72.DataBindings.Clear()
                .Label73.DataBindings.Clear()
                .Label74.DataBindings.Clear()
                .Label75.DataBindings.Clear()
                .Label76.DataBindings.Clear()
                .Label77.DataBindings.Clear()
                .Label78.DataBindings.Clear()
                .Label79.DataBindings.Clear()
                .Label80.DataBindings.Clear()
                .Label81.DataBindings.Clear()
                .Label82.DataBindings.Clear()
                .Label104.DataBindings.Clear()
                .Label105.DataBindings.Clear()
                .Label106.DataBindings.Clear()
                .Label364.DataBindings.Clear()
                .Label365.DataBindings.Clear()


                ' .Label324.DataBindings.Clear()
                ' .Label325.DataBindings.Clear()   'jedinicna cijena


                Dim i As Integer = .DataGridView2.CurrentRow.Index
                Dim BS As BindingSource = .FavoritiBindingSource
                Dim BS1 As BindingSource = .CijeneBindingSource

                If .DataGridView15.RowCount <= 1 Then
                    BS.AddNew()   'favoriti
                    BS1.AddNew()   'cijene
                End If

                BS.MoveLast()
                BS1.MoveLast()

                .TabControl1.SelectedIndex = 5    'izrada jelovnika Tab

                ' Dim Jed As Integer = 1
                '  Dim Jed As Double = Format((.TextBox85.Text / .Label363.Text), "0.00")
                'Dim Jed As Double = .TextBox85.Text   'staro, ne valja
                Dim Jed As Double = Format((.TextBox85.Text / .Label363.Text), "0.00")


                .Label27.DataBindings.Add(New Binding("Text", BS, "NazivNamirnice"))
                .Label27.Text = .TextBox5.Text   'favoriti  DBV

                .Label106.DataBindings.Add(New Binding("Text", BS, "Serviranja"))
                .Label106.Text = Jed

                .Label31.DataBindings.Add(New Binding("Text", BS, "Kolicina"))
                If .DataGridView2.Rows(i).Cells(2).Value.ToString = "N" Then
                    .Label31.Text = "N"
                Else
                    .Label31.Text = .DataGridView2.Rows(i).Cells(2).Value
                End If

                .Label32.DataBindings.Add(New Binding("Text", BS, "Mjera"))
                If .DataGridView2.Rows(i).Cells(3).Value.ToString = "N" Then
                    .Label32.Text = "N"
                Else
                    .Label32.Text = .DataGridView2.Rows(i).Cells(3).Value
                End If

                .Label33.DataBindings.Add(New Binding("Text", BS, "Masa_g"))
                If .DataGridView2.Rows(i).Cells(4).Value.ToString = "N" Then
                    .Label33.Text = "N"
                Else
                    .Label33.Text = .DataGridView2.Rows(i).Cells(4).Value
                End If

                .Label34.DataBindings.Add(New Binding("Text", BS, "Energija_kcal"))
                If .DataGridView2.Rows(i).Cells(5).Value.ToString = "N" Then
                    .Label34.Text = "N"
                Else
                    .Label34.Text = .DataGridView2.Rows(i).Cells(5).Value * Jed
                End If

                .Label35.DataBindings.Add(New Binding("Text", BS, "Energija_kJ"))
                If .DataGridView2.Rows(i).Cells(6).Value.ToString = "N" Then
                    .Label35.Text = "N"
                Else
                    .Label35.Text = .DataGridView2.Rows(i).Cells(6).Value * Jed
                End If

                .Label36.DataBindings.Add(New Binding("Text", BS, "Ugljikohidrati_g"))
                If .DataGridView2.Rows(i).Cells(7).Value.ToString = "N" Then
                    .Label36.Text = "N"
                Else
                    .Label36.Text = .DataGridView2.Rows(i).Cells(7).Value * Jed
                End If

                .Label37.DataBindings.Add(New Binding("Text", BS, "Bjelancevine_g"))
                If .DataGridView2.Rows(i).Cells(8).Value.ToString = "N" Then
                    .Label37.Text = "N"
                Else
                    .Label37.Text = .DataGridView2.Rows(i).Cells(8).Value * Jed
                End If

                .Label38.DataBindings.Add(New Binding("Text", BS, "Masti_g"))
                If .DataGridView2.Rows(i).Cells(9).Value.ToString = "N" Then
                    .Label38.Text = "N"
                Else
                    .Label38.Text = .DataGridView2.Rows(i).Cells(9).Value * Jed
                End If

                .Label39.DataBindings.Add(New Binding("Text", BS, "Skrob_g"))
                If .DataGridView2.Rows(i).Cells(10).Value.ToString = "N" Then
                    .Label39.Text = "N"
                Else
                    .Label39.Text = .DataGridView2.Rows(i).Cells(10).Value * Jed
                End If

                .Label40.DataBindings.Add(New Binding("Text", BS, "UkupniSeceri_g"))
                If .DataGridView2.Rows(i).Cells(11).Value.ToString = "N" Then
                    .Label40.Text = "N"
                Else
                    .Label40.Text = .DataGridView2.Rows(i).Cells(11).Value * Jed
                End If

                .Label41.DataBindings.Add(New Binding("Text", BS, "Glukoza_g"))
                If .DataGridView2.Rows(i).Cells(12).Value.ToString = "N" Then
                    .Label41.Text = "N"
                Else
                    .Label41.Text = .DataGridView2.Rows(i).Cells(12).Value * Jed
                End If

                .Label42.DataBindings.Add(New Binding("Text", BS, "Fruktoza_g"))
                If .DataGridView2.Rows(i).Cells(13).Value.ToString = "N" Then
                    .Label42.Text = "N"
                Else
                    .Label42.Text = .DataGridView2.Rows(i).Cells(13).Value * Jed
                End If

                .Label43.DataBindings.Add(New Binding("Text", BS, "Saharoza_g"))
                If .DataGridView2.Rows(i).Cells(14).Value.ToString = "N" Then
                    .Label43.Text = "N"
                Else
                    .Label43.Text = .DataGridView2.Rows(i).Cells(14).Value * Jed
                End If

                .Label44.DataBindings.Add(New Binding("Text", BS, "Maltoza_g"))
                If .DataGridView2.Rows(i).Cells(15).Value.ToString = "N" Then
                    .Label44.Text = "N"
                Else
                    .Label44.Text = .DataGridView2.Rows(i).Cells(15).Value * Jed
                End If

                .Label45.DataBindings.Add(New Binding("Text", BS, "Laktoza_g"))
                If .DataGridView2.Rows(i).Cells(16).Value.ToString = "N" Then
                    .Label45.Text = "N"
                Else
                    .Label45.Text = .DataGridView2.Rows(i).Cells(16).Value * Jed
                End If

                .Label46.DataBindings.Add(New Binding("Text", BS, "Vlakna_g"))
                If .DataGridView2.Rows(i).Cells(17).Value.ToString = "N" Then
                    .Label46.Text = "N"
                Else
                    .Label46.Text = .DataGridView2.Rows(i).Cells(17).Value * Jed
                End If

                .Label47.DataBindings.Add(New Binding("Text", BS, "ZasiceneMasti_g"))
                If .DataGridView2.Rows(i).Cells(18).Value.ToString = "N" Then
                    .Label47.Text = "N"
                Else
                    .Label47.Text = .DataGridView2.Rows(i).Cells(18).Value * Jed
                End If

                .Label48.DataBindings.Add(New Binding("Text", BS, "JednostrukoNezasiceneMasti_g"))
                If .DataGridView2.Rows(i).Cells(19).Value.ToString = "N" Then
                    .Label48.Text = "N"
                Else
                    .Label48.Text = .DataGridView2.Rows(i).Cells(19).Value * Jed
                End If

                .Label49.DataBindings.Add(New Binding("Text", BS, "VisestrukoNezasiceneMasti_g"))
                If .DataGridView2.Rows(i).Cells(20).Value.ToString = "N" Then
                    .Label49.Text = "N"
                Else
                    .Label49.Text = .DataGridView2.Rows(i).Cells(20).Value * Jed
                End If

                .Label50.DataBindings.Add(New Binding("Text", BS, "TransMasneKiseline_g"))
                If .DataGridView2.Rows(i).Cells(21).Value.ToString = "N" Then
                    .Label50.Text = "N"
                Else
                    .Label50.Text = .DataGridView2.Rows(i).Cells(21).Value * Jed
                End If

                .Label51.DataBindings.Add(New Binding("Text", BS, "Kolesterol_mg"))
                If .DataGridView2.Rows(i).Cells(22).Value.ToString = "N" Then
                    .Label51.Text = "N"
                Else
                    .Label51.Text = .DataGridView2.Rows(i).Cells(22).Value * Jed
                End If

                .Label52.DataBindings.Add(New Binding("Text", BS, "Natrij_mg"))
                If .DataGridView2.Rows(i).Cells(23).Value.ToString = "N" Then
                    .Label52.Text = "N"
                Else
                    .Label52.Text = .DataGridView2.Rows(i).Cells(23).Value * Jed
                End If

                .Label53.DataBindings.Add(New Binding("Text", BS, "Kalij_mg"))
                If .DataGridView2.Rows(i).Cells(24).Value.ToString = "N" Then
                    .Label53.Text = "N"
                Else
                    .Label53.Text = .DataGridView2.Rows(i).Cells(24).Value * Jed
                End If

                .Label54.DataBindings.Add(New Binding("Text", BS, "Kalcij_mg"))
                If .DataGridView2.Rows(i).Cells(25).Value.ToString = "N" Then
                    .Label54.Text = "N"
                Else
                    .Label54.Text = .DataGridView2.Rows(i).Cells(25).Value * Jed
                End If

                .Label55.DataBindings.Add(New Binding("Text", BS, "Magnezij_mg"))
                If .DataGridView2.Rows(i).Cells(26).Value.ToString = "N" Then
                    .Label55.Text = "N"
                Else
                    .Label55.Text = .DataGridView2.Rows(i).Cells(26).Value * Jed
                End If

                .Label56.DataBindings.Add(New Binding("Text", BS, "Fosfor_mg"))
                If .DataGridView2.Rows(i).Cells(27).Value.ToString = "N" Then
                    .Label56.Text = "N"
                Else
                    .Label56.Text = .DataGridView2.Rows(i).Cells(27).Value * Jed
                End If

                .Label57.DataBindings.Add(New Binding("Text", BS, "Zeljezo_mg"))
                If .DataGridView2.Rows(i).Cells(28).Value.ToString = "N" Then
                    .Label57.Text = "N"
                Else
                    .Label57.Text = .DataGridView2.Rows(i).Cells(28).Value * Jed
                End If

                .Label58.DataBindings.Add(New Binding("Text", BS, "Bakar_mg"))
                If .DataGridView2.Rows(i).Cells(29).Value.ToString = "N" Then
                    .Label58.Text = "N"
                Else
                    .Label58.Text = .DataGridView2.Rows(i).Cells(29).Value * Jed
                End If

                .Label59.DataBindings.Add(New Binding("Text", BS, "Cink_mg"))
                If .DataGridView2.Rows(i).Cells(30).Value.ToString = "N" Then
                    .Label59.Text = "N"
                Else
                    .Label59.Text = .DataGridView2.Rows(i).Cells(30).Value * Jed
                End If

                .Label60.DataBindings.Add(New Binding("Text", BS, "Klor_mg"))
                If .DataGridView2.Rows(i).Cells(31).Value.ToString = "N" Then
                    .Label60.Text = "N"
                Else
                    .Label60.Text = .DataGridView2.Rows(i).Cells(31).Value * Jed
                End If

                .Label61.DataBindings.Add(New Binding("Text", BS, "Mangan_mg"))
                If .DataGridView2.Rows(i).Cells(32).Value.ToString = "N" Then
                    .Label61.Text = "N"
                Else
                    .Label61.Text = .DataGridView2.Rows(i).Cells(32).Value * Jed
                End If

                .Label62.DataBindings.Add(New Binding("Text", BS, "Selen_mikro_g"))
                If .DataGridView2.Rows(i).Cells(33).Value.ToString = "N" Then
                    .Label62.Text = "N"
                Else
                    .Label62.Text = .DataGridView2.Rows(i).Cells(33).Value * Jed
                End If

                .Label63.DataBindings.Add(New Binding("Text", BS, "Jod_mikro_g"))
                If .DataGridView2.Rows(i).Cells(34).Value.ToString = "N" Then
                    .Label63.Text = "N"
                Else
                    .Label63.Text = .DataGridView2.Rows(i).Cells(34).Value * Jed
                End If

                .Label64.DataBindings.Add(New Binding("Text", BS, "Retinol_mikro_g"))
                If .DataGridView2.Rows(i).Cells(35).Value.ToString = "N" Then
                    .Label64.Text = "N"
                Else
                    .Label64.Text = .DataGridView2.Rows(i).Cells(35).Value * Jed
                End If

                .Label65.DataBindings.Add(New Binding("Text", BS, "Karoten_mikro_g"))
                If .DataGridView2.Rows(i).Cells(36).Value.ToString = "N" Then
                    .Label65.Text = "N"
                Else
                    .Label65.Text = .DataGridView2.Rows(i).Cells(36).Value * Jed
                End If

                .Label66.DataBindings.Add(New Binding("Text", BS, "VitaminD_mikro_g"))
                If .DataGridView2.Rows(i).Cells(37).Value.ToString = "N" Then
                    .Label66.Text = "N"
                Else
                    .Label66.Text = .DataGridView2.Rows(i).Cells(37).Value * Jed
                End If

                .Label67.DataBindings.Add(New Binding("Text", BS, "VitaminE_mg"))
                If .DataGridView2.Rows(i).Cells(38).Value.ToString = "N" Then
                    .Label67.Text = "N"
                Else
                    .Label67.Text = .DataGridView2.Rows(i).Cells(38).Value * Jed
                End If

                .Label68.DataBindings.Add(New Binding("Text", BS, "VitaminB1_mg"))
                If .DataGridView2.Rows(i).Cells(39).Value.ToString = "N" Then
                    .Label68.Text = "N"
                Else
                    .Label68.Text = .DataGridView2.Rows(i).Cells(39).Value * Jed
                End If

                .Label69.DataBindings.Add(New Binding("Text", BS, "VitaminB2_mg"))
                If .DataGridView2.Rows(i).Cells(40).Value.ToString = "N" Then
                    .Label69.Text = "N"
                Else
                    .Label69.Text = .DataGridView2.Rows(i).Cells(40).Value * Jed
                End If

                .Label70.DataBindings.Add(New Binding("Text", BS, "VitaminB3_mg"))
                If .DataGridView2.Rows(i).Cells(41).Value.ToString = "N" Then
                    .Label70.Text = "N"
                Else
                    .Label70.Text = .DataGridView2.Rows(i).Cells(41).Value * Jed
                End If

                .Label71.DataBindings.Add(New Binding("Text", BS, "VitaminB6_mg"))
                If .DataGridView2.Rows(i).Cells(42).Value.ToString = "N" Then
                    .Label71.Text = "N"
                Else
                    .Label71.Text = .DataGridView2.Rows(i).Cells(42).Value * Jed
                End If

                .Label72.DataBindings.Add(New Binding("Text", BS, "VitaminB12_mikro_g"))
                If .DataGridView2.Rows(i).Cells(43).Value.ToString = "N" Then
                    .Label72.Text = "N"
                Else
                    .Label72.Text = .DataGridView2.Rows(i).Cells(43).Value * Jed
                End If

                .Label73.DataBindings.Add(New Binding("Text", BS, "Folat_mikro_g"))
                If .DataGridView2.Rows(i).Cells(44).Value.ToString = "N" Then
                    .Label73.Text = "N"
                Else
                    .Label73.Text = .DataGridView2.Rows(i).Cells(44).Value * Jed
                End If

                .Label74.DataBindings.Add(New Binding("Text", BS, "PantotenskaKiselina_mg"))
                If .DataGridView2.Rows(i).Cells(45).Value.ToString = "N" Then
                    .Label74.Text = "N"
                Else
                    .Label74.Text = .DataGridView2.Rows(i).Cells(45).Value * Jed
                End If

                .Label75.DataBindings.Add(New Binding("Text", BS, "Biotin_mikro_g"))
                If .DataGridView2.Rows(i).Cells(46).Value.ToString = "N" Then
                    .Label75.Text = "N"
                Else
                    .Label75.Text = .DataGridView2.Rows(i).Cells(46).Value * Jed
                End If

                .Label76.DataBindings.Add(New Binding("Text", BS, "VitaminC_mg"))
                If .DataGridView2.Rows(i).Cells(47).Value.ToString = "N" Then
                    .Label76.Text = "N"
                Else
                    .Label76.Text = .DataGridView2.Rows(i).Cells(47).Value * Jed
                End If

                .Label77.DataBindings.Add(New Binding("Text", BS, "VitaminK_mikro_g"))
                If .DataGridView2.Rows(i).Cells(48).Value.ToString = "N" Then
                    .Label77.Text = "N"
                Else
                    .Label77.Text = .DataGridView2.Rows(i).Cells(48).Value * Jed
                End If

                .Label78.DataBindings.Add(New Binding("Text", BS, "Zitarice"))
                .Label78.Text = .DataGridView2.Rows(i).Cells(49).Value * Jed
                .Label79.DataBindings.Add(New Binding("Text", BS, "Povrce"))
                .Label79.Text = .DataGridView2.Rows(i).Cells(50).Value * Jed
                .Label80.DataBindings.Add(New Binding("Text", BS, "Voce"))
                .Label80.Text = .DataGridView2.Rows(i).Cells(51).Value * Jed
                .Label81.DataBindings.Add(New Binding("Text", BS, "Meso"))
                .Label81.Text = .DataGridView2.Rows(i).Cells(52).Value * Jed
                .Label82.DataBindings.Add(New Binding("Text", BS, "Mlijeko"))
                .Label82.Text = .DataGridView2.Rows(i).Cells(53).Value * Jed
                .Label104.DataBindings.Add(New Binding("Text", BS, "Masti"))
                .Label104.Text = .DataGridView2.Rows(i).Cells(54).Value * Jed
                .Label105.DataBindings.Add(New Binding("Text", BS, "OstaleNamirnice"))
                .Label105.Text = .DataGridView2.Rows(i).Cells(55).Value * Jed
                .Label364.DataBindings.Add(New Binding("Text", BS, "SkupinaNamirnicaGubiciVitamina"))
                .Label364.Text = .DataGridView2.Rows(i).Cells(56).Value
                .Label365.DataBindings.Add(New Binding("Text", BS, "SkupinaNamirnica"))
                .Label365.Text = .DataGridView2.Rows(i).Cells(57).Value


                '.Label105.Text = .TextBox4.Text
                BS.AddNew()


                .TabControl1.SelectedIndex = 10  'cijene Tab
                .Label327.DataBindings.Clear()
                .Label327.DataBindings.Add(New Binding("Text", BS1, "NazivNamirnice"))
                .Label327.Text = .TextBox5.Text    'cijene DBV

                .Label326.DataBindings.Clear()
                .Label326.DataBindings.Add(New Binding("Text", BS1, "JedinicnaCijena"))    'cijene DBV
                .Label326.Text = .Label331.Text


                BS1.AddNew()


                .DataGridView5.CurrentRow.Selected = False
                .DataGridView9.CurrentRow.Selected = False
                .DataGridView11.CurrentRow.Selected = False
                .DataGridView12.CurrentRow.Selected = False
                .DataGridView13.CurrentRow.Selected = False
                .DataGridView14.CurrentRow.Selected = False

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
                .Label364.Visible = False
                .Label365.Visible = False

                'cijene-favoriti-izrada jelovnika form
                .Label329.Visible = False
                .Label319.Visible = False
                .Label331.Visible = False

                'cijene
                .Label327.Visible = False
                .Label325.Visible = False
                .Label326.Visible = False

            End With
        End If
    End Sub
End Module
