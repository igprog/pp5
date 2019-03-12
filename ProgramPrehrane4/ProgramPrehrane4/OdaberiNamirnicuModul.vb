Module OdaberiNamirnicuModul
    Sub OdaberiNamirnicu()
        On Error Resume Next
        With Form1
            .ComboBox11.Text = "Termička obrada"   'Termicka obrada
            .TextBox77.Text = ""   'Cijena
            Dim Serviranje As TextBox = .TextBox85      'Serviranja
            '.ComboBox4.Text = 1
            .TextBox81.Text = 1000   'Kolicina g
            Dim i As Integer = .DataGridView2.CurrentRow.Index
            Dim a As Integer = 0     'DGV2, skupina namirnica
            Dim Serv As Label = .Label337     'Serviranje, skupina namirnica

            .TextBox5.Text = .DataGridView2.Rows(i).Cells(1).Value   'Naziv Namirnice
            .TextBox12.Text = .DataGridView2.Rows(i).Cells(2).Value   'Kolicina
            .Label229.Text = .DataGridView2.Rows(i).Cells(2).Value   'Kolicina
            .Label103.Text = .DataGridView2.Rows(i).Cells(3).Value   'Mjera
            .Label230.Text = .DataGridView2.Rows(i).Cells(3).Value   'Mjera
            .TextBox6.Text = .DataGridView2.Rows(i).Cells(4).Value   'Masa_g
            .Label228.Text = .DataGridView2.Rows(i).Cells(4).Value   'Masa_g

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Zitarice" Then
                a = 49
                Serv.Text = "serv. žitarica i proizvoda od žita"
                '      Serviranje.Text = .DataGridView2.Rows(i).Cells(49).Value
                '      .Label363.Text = .DataGridView2.Rows(i).Cells(49).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Povrce" Then
                a = 50
                Serv.Text = "serv. povrća"
                '       Serviranje.Text = .DataGridView2.Rows(i).Cells(50).Value
                '     .Label363.Text = .DataGridView2.Rows(i).Cells(50).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Voce" Then
                a = 51
                Serv.Text = "serv. voća"
                '      Serviranje.Text = .DataGridView2.Rows(i).Cells(51).Value
                '      .Label363.Text = .DataGridView2.Rows(i).Cells(51).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "IzrazitoNemasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "NemasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "SrednjeMasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "MasnoMeso" Then
                a = 52
                Serv.Text = "serv. mesa i zamjena"
                '   Serviranje.Text = .DataGridView2.Rows(i).Cells(52).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(52).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "ObranoMlijeko" _
               Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "DjelomicnoObranoMlijeko" _
               Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "PunomasnoMlijeko" Then
                a = 53
                Serv.Text = "serv. mlijeka i mliječnih proizvoda"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(53).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(53).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "ZasiceneMasti" _
              Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "VisestrukoNezasiceneMasti" _
              Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "JednostrukoNezasiceneMasti" Then
                a = 54
                Serv.Text = "serv. masti"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(54).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(54).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "OstaleNamirnice" Then
                a = 55
                Serv.Text = "serv. ostalih namirnica"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(55).Value
                ' .Label363.Text = .DataGridView2.Rows(i).Cells(55).Value
            End If
            Serviranje.Text = .DataGridView2.Rows(i).Cells(a).Value
            .Label363.Text = .DataGridView2.Rows(i).Cells(a).Value


            If Serviranje.Text = 0.5 Or Serviranje.Text = 1.5 Or Serviranje.Text = 2.5 Or Serviranje.Text = 3.5 Then
                Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
                .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            Else
                If Serviranje.Text = 0.006 Then
                    Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
                    .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
                Else
                    If Serviranje.Text = 1 Or Serviranje.Text = 2 Or Serviranje.Text = 3 Or Serviranje.Text = 4 Or Serviranje.Text = 5 Then
                        Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
                        .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
                    Else
                        Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
                        .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
                    End If
                End If
            End If
            '        Dim Serv As Double = .DataGridView2.Rows(i).Cells(a).Value
            '       Select Case Serv
            '  Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            '  Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
            '  .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
            '    Case 0.006
            '      Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
            '     .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
            '       Case 0.5, 1.5, 2.5, 3.5, 4.5
            '  Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            ' .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            '        Case 0 To 1
            '   Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '  .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '     Case Else
            '    Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '   .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            'End Select

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "MjesoviteNamirnice" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = "serv. mješovitih namirnica"
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Jela" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = "serv."
                .TextBox12.Text = 1
                .RadioButton19.Checked = True
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString <> "Jela" And My.Settings.PP5PremiumAktivacija = "Da" Then
                .TextBox77.Enabled = True   'omoguci unos cijene
                .CheckBox1.Enabled = True   'omoguci unos cijena namirnica
            Else
                ' Serviranje.Text = 1
                ' .Label363.Text = 1
                '  Serv.Text = "serv."
                '.RadioButton19.Checked = True
                '.TextBox12.Text = 1
                .TextBox77.Enabled = False   'onemoguci unos cijene
                .CheckBox1.Enabled = False   'onemoguci unos cijene pripremljenog jela
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "MojeNamirnice" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = ""
            End If

            '    Serviranje.Text = .DataGridView2.Rows(i).Cells(a).Value
            '   .Label363.Text = .DataGridView2.Rows(i).Cells(a).Value



            '           Dim number As Integer = 8
            '          Select Case number
            '             Case 1 To 5
            '        Debug.WriteLine("Between 1 and 5, inclusive")
            '       ' The following is the only Case clause that evaluates to True. 
            '          Case 6, 7, 8
            '     Debug.WriteLine("Between 6 and 8, inclusive")
            '        Case 9 To 10
            '   Debug.WriteLine("Equal to 9 or 10")
            '      Case Else
            ' Debug.WriteLine("Not between 1 and 10, inclusive")
            'End Select

            ' .TabControl1.SelectedIndex = 10    'Cijene
            Dim Namirnica As TextBox = .TextBox5
            Dim DGV As DataGridView = .DataGridView15
            Dim j As Integer
            For j = 0 To DGV.RowCount - 1
                If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                    If DGV.Rows(j).Cells(1).Value = Namirnica.Text Then
                        .TextBox77.Text = DGV.Rows(j).Cells(2).Value   'Cijena
                        '  MsgBox("ok")
                        '   MsgBox("Namirnica " & Namirnica.Text & " već postoji u Favoritima.")
                        '      .TabControl1.SelectedIndex = 4   'vrati u izradu jelovnika
                        '   Exit Sub
                        ' Else
                        '   .TextBox77.Text = ""
                    End If
                End If
            Next j
            '   .TabControl1.SelectedIndex = 5   'Izrada jelovnika
            .TextBox7.Text = "Pretraži"
        End With
        Mjera()
        TermickeObrade()   'ComboBox11


        'jos jednom zbog greske kod promjene nacina prikazivanja kolicine
        With Form1
            .ComboBox11.Text = "Termička obrada"   'Termicka obrada
            ' .ComboBox4.Text = 1
            Dim Serviranje As TextBox = .TextBox85      'Serviranja
            .TextBox81.Text = 1000   'Kolicina g
            Dim i As Integer = .DataGridView2.CurrentRow.Index
            Dim a As Integer = 0     'DGV2, skupina namirnica
            Dim Serv As Label = .Label337     'Serviranje, skupina namirnica

            .TextBox5.Text = .DataGridView2.Rows(i).Cells(1).Value   'Naziv Namirnice
            .TextBox12.Text = .DataGridView2.Rows(i).Cells(2).Value   'Kolicina
            .Label229.Text = .DataGridView2.Rows(i).Cells(2).Value   'Kolicina
            .Label103.Text = .DataGridView2.Rows(i).Cells(3).Value   'Mjera
            .Label230.Text = .DataGridView2.Rows(i).Cells(3).Value   'Mjera
            .TextBox6.Text = .DataGridView2.Rows(i).Cells(4).Value   'Masa_g
            .Label228.Text = .DataGridView2.Rows(i).Cells(4).Value   'Masa_g

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Zitarice" Then
                a = 49
                Serv.Text = "serv. žitarica i proizvoda od žita"
                '      Serviranje.Text = .DataGridView2.Rows(i).Cells(49).Value
                '      .Label363.Text = .DataGridView2.Rows(i).Cells(49).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Povrce" Then
                a = 50
                Serv.Text = "serv. povrća"
                '       Serviranje.Text = .DataGridView2.Rows(i).Cells(50).Value
                '     .Label363.Text = .DataGridView2.Rows(i).Cells(50).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Voce" Then
                a = 51
                Serv.Text = "serv. voća"
                '      Serviranje.Text = .DataGridView2.Rows(i).Cells(51).Value
                '      .Label363.Text = .DataGridView2.Rows(i).Cells(51).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "IzrazitoNemasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "NemasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "SrednjeMasnoMeso" _
                Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "MasnoMeso" Then
                a = 52
                Serv.Text = "serv. mesa i zamjena"
                '   Serviranje.Text = .DataGridView2.Rows(i).Cells(52).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(52).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "ObranoMlijeko" _
               Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "DjelomicnoObranoMlijeko" _
               Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "PunomasnoMlijeko" Then
                a = 53
                Serv.Text = "serv. mlijeka i mliječnih proizvoda"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(53).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(53).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "ZasiceneMasti" _
              Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "VisestrukoNezasiceneMasti" _
              Or .DataGridView2.Rows(i).Cells(57).Value.ToString = "JednostrukoNezasiceneMasti" Then
                a = 54
                Serv.Text = "serv. masti"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(54).Value
                '  .Label363.Text = .DataGridView2.Rows(i).Cells(54).Value
            End If
            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "OstaleNamirnice" Then
                a = 55
                Serv.Text = "serv. ostalih namirnica"
                '  Serviranje.Text = .DataGridView2.Rows(i).Cells(55).Value
                ' .Label363.Text = .DataGridView2.Rows(i).Cells(55).Value
            End If
            Serviranje.Text = .DataGridView2.Rows(i).Cells(a).Value
            .Label363.Text = .DataGridView2.Rows(i).Cells(a).Value


            If Serviranje.Text = 0.5 Or Serviranje.Text = 1.5 Or Serviranje.Text = 2.5 Or Serviranje.Text = 3.5 Then
                Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
                .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            Else
                If Serviranje.Text = 0.006 Then
                    Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
                    .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
                Else
                    If Serviranje.Text = 1 Or Serviranje.Text = 2 Or Serviranje.Text = 3 Or Serviranje.Text = 4 Or Serviranje.Text = 5 Then
                        Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
                        .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
                    Else
                        Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
                        .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
                    End If
                End If
            End If
            '        Dim Serv As Double = .DataGridView2.Rows(i).Cells(a).Value
            '       Select Case Serv
            '  Case 1, 2, 3, 4, 5, 6, 7, 8, 9
            '  Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
            '  .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0")
            '    Case 0.006
            '      Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
            '     .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.000")
            '       Case 0.5, 1.5, 2.5, 3.5, 4.5
            '  Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            ' .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.0")
            '        Case 0 To 1
            '   Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '  .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '     Case Else
            '    Serviranje.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            '   .Label363.Text = Format(.DataGridView2.Rows(i).Cells(a).Value, "0.00")
            'End Select

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "MjesoviteNamirnice" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = "serv. mješovitih namirnica"
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "Jela" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = "serv."
                .TextBox12.Text = 1
                .RadioButton19.Checked = True
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString <> "Jela" And My.Settings.PP5PremiumAktivacija = "Da" Then
                .TextBox77.Enabled = True   'omoguci unos cijene
                .CheckBox1.Enabled = True   'omoguci unos cijena namirnica
            Else
                ' Serviranje.Text = 1
                ' .Label363.Text = 1
                '  Serv.Text = "serv."
                '.RadioButton19.Checked = True
                '.TextBox12.Text = 1
                .TextBox77.Enabled = False   'onemoguci unos cijene
                .CheckBox1.Enabled = False   'onemoguci unos cijene pripremljenog jela
            End If

            If .DataGridView2.Rows(i).Cells(57).Value.ToString = "MojeNamirnice" Then
                Serviranje.Text = 1
                .Label363.Text = 1
                Serv.Text = ""
            End If




            '   .TabControl1.SelectedIndex = 10   'Cijena
            Dim Namirnica As TextBox = .TextBox5
            Dim DGV As DataGridView = .DataGridView15
            Dim j As Integer
            For j = 0 To DGV.RowCount - 1
                If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                    If DGV.Rows(j).Cells(1).Value = Namirnica.Text Then
                        .TextBox77.Text = DGV.Rows(j).Cells(2).Value   'Cijena
                        '   MsgBox("ok")
                        '   MsgBox("Namirnica " & Namirnica.Text & " već postoji u Favoritima.")
                        '      .TabControl1.SelectedIndex = 4   'vrati u izradu jelovnika
                        '   Exit Sub
                        '    Else
                        '  .TextBox77.Text = ""
                    End If
                End If
            Next j
            '  .TabControl1.SelectedIndex = 5   'Izrada jelovnika

            .TextBox7.Text = "Pretraži"
        End With
        Mjera()
        TermickeObrade()   'ComboBox11

    End Sub
End Module
