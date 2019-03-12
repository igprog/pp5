Module ParametriModul
    Sub Parametri()
        On Error Resume Next
        With Form1
            .ListBox15.Items.Clear()   'Energija i nutrijenti
            .ListBox82.Items.Clear()   'Nutrijenti % 
            .ListBox21.Items.Clear()   'Karoten
            .ListBox22.Items.Clear()   'Zasicene masti, Trans, Kolesterol
            .ListBox23.Items.Clear()   'Nartij, Kalij, Klor
            .ListBox24.Items.Clear()   'Ostali mikronutrijenti
            .ListBox46.Items.Clear()   'Skrob - Laktoza
            .ListBox35.Items.Clear()   'Ostale namirnice
         
            Dim DGV As DataGridView
            Dim Ko As Integer = 1   'broj korisnika jelovnika

            Dim Energija_kcal As Double = 0
            Dim Ugljikohidrati As Double = 0
            Dim Bjelancevine As Double = 0
            Dim Masti As Double = 0

            Dim Energija_kJ As Double = 0

            Dim Skrob As Double = 0
            Dim UkupniSeceri As Double = 0
            Dim Glukoza As Double = 0
            Dim Fruktoza As Double = 0
            Dim Saharoza As Double = 0
            Dim Maltoza As Double = 0
            Dim Laktoza As Double = 0


            Dim Karoten As Double = 0
            Dim ZasiceneMasti As Double = 0
            Dim Trans As Double = 0
            Dim Kolesterol As Double = 0
            Dim Natrij As Double = 0
            Dim Kalij As Double = 0
            Dim Klor As Double = 0
            Dim Vlakna As Double = 0
            Dim JednostrukoNezasicene As Double = 0
            Dim VisestrukoNezasicene As Double = 0
            Dim Kalcij As Double = 0
            Dim Magnezij As Double = 0
            Dim Fosfor As Double = 0
            Dim Zeljezo As Double = 0
            Dim Bakar As Double = 0
            Dim Cink As Double = 0
            Dim Mangan As Double = 0
            Dim Selen As Double = 0
            Dim Jod As Double = 0
            Dim Retinol As Double = 0
            Dim D As Double = 0
            Dim E As Double = 0
            Dim B1 As Double = 0
            Dim B2 As Double = 0
            Dim B3 As Double = 0
            Dim B6 As Double = 0
            Dim B12 As Double = 0
            Dim Folat As Double = 0
            Dim PantotenskaKiselina As Double = 0
            Dim Biotin As Double = 0
            Dim C As Double = 0
            Dim K As Double = 0

            Dim i As Integer


            'DORUČAK
            '  DGV = .DataGridView5
            Dim b As Integer
            For b = 1 To 6
                If b = 1 Then DGV = .DataGridView5 'dorucak
                If b = 2 Then DGV = .DataGridView9 'jutarnja uzina
                If b = 3 Then DGV = .DataGridView11 'rucak
                If b = 4 Then DGV = .DataGridView12 'popodnevna uzina
                If b = 5 Then DGV = .DataGridView13 'vecera
                If b = 6 Then DGV = .DataGridView14 'obrok pred spavanje

                For i = 0 To DGV.RowCount - 1
                    Energija_kcal = Energija_kcal + DGV.Rows(i).Cells(11).Value / Ko
                    Ugljikohidrati = Ugljikohidrati + DGV.Rows(i).Cells(13).Value / Ko
                    Bjelancevine = Bjelancevine + DGV.Rows(i).Cells(14).Value / Ko
                    Masti = Masti + DGV.Rows(i).Cells(15).Value / Ko

                    '  Energija_kJ = Energija_kJ + DGV.Rows(i).Cells(12).Value / Ko
                    Energija_kJ = Energija_kcal * 4.186

                    Skrob = Skrob + DGV.Rows(i).Cells(16).Value / Ko
                    UkupniSeceri = UkupniSeceri + DGV.Rows(i).Cells(17).Value / Ko
                    Glukoza = Glukoza + DGV.Rows(i).Cells(18).Value / Ko
                    Fruktoza = Fruktoza + DGV.Rows(i).Cells(19).Value / Ko
                    Saharoza = Saharoza + DGV.Rows(i).Cells(20).Value / Ko
                    Maltoza = Maltoza + DGV.Rows(i).Cells(21).Value / Ko
                    Laktoza = Laktoza + DGV.Rows(i).Cells(22).Value / Ko

                    Karoten = Karoten + DGV.Rows(i).Cells(42).Value / Ko
                    ZasiceneMasti = ZasiceneMasti + DGV.Rows(i).Cells(24).Value / Ko
                    Trans = Trans + DGV.Rows(i).Cells(27).Value / Ko
                    Kolesterol = Kolesterol + DGV.Rows(i).Cells(28).Value / Ko
                    Natrij = Natrij + DGV.Rows(i).Cells(29).Value / Ko
                    Kalij = Kalij + DGV.Rows(i).Cells(30).Value / Ko
                    Klor = Klor + DGV.Rows(i).Cells(37).Value / Ko
                    Vlakna = Vlakna + DGV.Rows(i).Cells(23).Value / Ko
                    JednostrukoNezasicene = JednostrukoNezasicene + DGV.Rows(i).Cells(25).Value / Ko
                    VisestrukoNezasicene = VisestrukoNezasicene + DGV.Rows(i).Cells(26).Value / Ko
                    Kalcij = Kalcij + DGV.Rows(i).Cells(31).Value / Ko
                    Magnezij = Magnezij + DGV.Rows(i).Cells(32).Value / Ko
                    Fosfor = Fosfor + DGV.Rows(i).Cells(33).Value / Ko
                    Zeljezo = Zeljezo + DGV.Rows(i).Cells(34).Value / Ko
                    Bakar = Bakar + DGV.Rows(i).Cells(35).Value / Ko
                    Cink = Cink + DGV.Rows(i).Cells(36).Value / Ko
                    Mangan = Mangan + DGV.Rows(i).Cells(38).Value / Ko
                    Selen = Selen + DGV.Rows(i).Cells(39).Value / Ko
                    Jod = Jod + DGV.Rows(i).Cells(40).Value / Ko
                    Retinol = Retinol + DGV.Rows(i).Cells(41).Value / Ko
                    D = D + DGV.Rows(i).Cells(43).Value / Ko
                    E = E + DGV.Rows(i).Cells(44).Value / Ko
                    B1 = B1 + DGV.Rows(i).Cells(45).Value / Ko
                    B2 = B2 + DGV.Rows(i).Cells(46).Value / Ko
                    B3 = B3 + DGV.Rows(i).Cells(47).Value / Ko
                    B6 = B6 + DGV.Rows(i).Cells(48).Value / Ko
                    B12 = B12 + DGV.Rows(i).Cells(49).Value / Ko
                    Folat = Folat + DGV.Rows(i).Cells(50).Value / Ko
                    PantotenskaKiselina = PantotenskaKiselina + DGV.Rows(i).Cells(51).Value / Ko
                    Biotin = Biotin + DGV.Rows(i).Cells(52).Value / Ko
                    C = C + DGV.Rows(i).Cells(53).Value / Ko
                    K = K + DGV.Rows(i).Cells(54).Value / Ko
                Next i
            Next b

            ' .Label287.Text = "(" & Format(Energija_kJ, "0") & " kJ)"   'Energija_kJ

            .ListBox21.Items.Insert(0, Format(Karoten, "0.0"))

            .ListBox22.Items.Insert(0, Format(ZasiceneMasti, "0.0"))
            .ListBox22.Items.Insert(1, Format(Trans, "0.0"))
            .ListBox22.Items.Insert(2, Format(Kolesterol, "0"))

            .ListBox23.Items.Insert(0, Format(Natrij, "0"))
            .ListBox23.Items.Insert(1, Format(Kalij, "0"))
            .ListBox23.Items.Insert(2, Format(Klor, "0"))

            .ListBox24.Items.Insert(0, Format(Vlakna, "0.0"))
            .ListBox24.Items.Insert(1, Format(JednostrukoNezasicene, "0.0"))
            .ListBox24.Items.Insert(2, Format(VisestrukoNezasicene, "0.0"))
            .ListBox24.Items.Insert(3, Format(Kalcij, "0"))
            .ListBox24.Items.Insert(4, Format(Magnezij, "0"))
            .ListBox24.Items.Insert(5, Format(Fosfor, "0"))
            .ListBox24.Items.Insert(6, Format(Zeljezo, "0.0"))
            .ListBox24.Items.Insert(7, Format(Bakar, "0.0"))
            .ListBox24.Items.Insert(8, Format(Cink, "0.0"))
            .ListBox24.Items.Insert(9, Format(Mangan, "0.0"))
            .ListBox24.Items.Insert(10, Format(Selen, "0.0"))
            .ListBox24.Items.Insert(11, Format(Jod, "0"))
            .ListBox24.Items.Insert(12, Format(Retinol, "0"))
            .ListBox24.Items.Insert(13, Format(D, "0.0"))
            .ListBox24.Items.Insert(14, Format(E, "0.0"))
            .ListBox24.Items.Insert(15, Format(B1, "0.0"))
            .ListBox24.Items.Insert(16, Format(B2, "0.0"))
            .ListBox24.Items.Insert(17, Format(B3, "0.0"))
            .ListBox24.Items.Insert(18, Format(B6, "0.0"))
            .ListBox24.Items.Insert(19, Format(B12, "0.0"))
            .ListBox24.Items.Insert(20, Format(Folat, "0.0"))
            .ListBox24.Items.Insert(21, Format(PantotenskaKiselina, "0.0"))
            .ListBox24.Items.Insert(22, Format(Biotin, "0.0"))
            .ListBox24.Items.Insert(23, Format(C, "0.0"))
            .ListBox24.Items.Insert(24, Format(K, "0.0"))

            .ListBox46.Items.Insert(0, Format(Skrob, "0.0"))
            .ListBox46.Items.Insert(1, Format(UkupniSeceri, "0.0"))
            .ListBox46.Items.Insert(2, Format(Glukoza, "0.0"))
            .ListBox46.Items.Insert(3, Format(Fruktoza, "0.0"))
            .ListBox46.Items.Insert(4, Format(Saharoza, "0.0"))
            .ListBox46.Items.Insert(5, Format(Maltoza, "0.0"))
            .ListBox46.Items.Insert(6, Format(Laktoza, "0.0"))

            'Ostale namirnice
            .ListBox35.Items.Insert(0, Format(Val(.Label285.Text), "0"))  'Suma

            'Energija i nutrijenti
            '        .ListBox15.Items.Insert(0, .ListBox10.Items(0) & "  (" & Format(Energija_kJ, "0") & " kJ) ")   'Energija kcal  (Energija kJ)  %
            '       .ListBox15.Items.Insert(1, .ListBox10.Items(1))   'prazno
            '      .ListBox15.Items.Insert(2, .ListBox10.Items(2))   'Ugljikohidrati g
            '     .ListBox15.Items.Insert(3, .ListBox10.Items(3))   'Bjelančevine g
            '    .ListBox15.Items.Insert(4, .ListBox10.Items(4))   'Masti g

            .ListBox15.Items.Insert(0, Format(Energija_kcal, "0") & " kcal" & "  (" & Format(Energija_kJ, "0") & " kJ) ")   'Energija kcal  (Energija kJ)  %
            .ListBox15.Items.Insert(1, "")   'prazno
            .ListBox15.Items.Insert(2, Format(Ugljikohidrati, "0") & " g")   'Ugljikohidrati g
            .ListBox15.Items.Insert(3, Format(Bjelancevine, "0") & " g")   'Bjelančevine g
            .ListBox15.Items.Insert(4, Format(Masti, "0") & " g")   'Masti g

            .Label287.Text = .ListBox11.Items(0)  'Energija %
            .ListBox82.Items.Insert(0, .ListBox11.Items(2))   'Ugljikohidrati (%)
            .ListBox82.Items.Insert(1, .ListBox11.Items(3))   'Bjelancevine (%)
            .ListBox82.Items.Insert(2, .ListBox11.Items(4))   'Masti (%)

        End With
    End Sub
End Module
