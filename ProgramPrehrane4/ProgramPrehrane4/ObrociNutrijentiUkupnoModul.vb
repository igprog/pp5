Module ObrociNutrijentiUkupnoModul
    Sub ObrociNutrijentiUkupno()
        On Error Resume Next
        With Form1
            Dim DGV As DataGridView
            Dim K As Integer = 1

            Dim energ As Double = 0
            Dim uglj As Double = 0
            Dim bjel As Double = 0
            Dim mast As Double = 0

            Dim i As Integer

            Dim EnergMin As Integer = 0   'Minimalna energetska vrijednost obroka
            Dim EnergMax As Integer = 0  ' Maksimalna energetska vrijednost obroka

            Dim EnergJelovnika As Double = .Label14.Text    'Energetska vrijednost odabranih namirnica

            .ListBox8.Items.Clear()   'dorucak - ukupno (g)
            .ListBox12.Items.Clear()    'dorucak - ukupno (%)

            'DORUČAK
            DGV = .DataGridView5
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox8.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox8.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox8.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox8.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox12.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox12.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox12.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox12.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox39.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            If .CheckBox3.Checked = False And .CheckBox7.Checked = False Then    'jutarnja uzina i popodnevna uzina
                EnergMin = 25
                EnergMax = 35
                .ListBox16.Items.Clear()
                .ListBox16.Items.Add(EnergMin & "-" & EnergMax & "%")
            Else
                EnergMin = 20        'ostali obroci
                EnergMax = 25
                .ListBox16.Items.Clear()
                .ListBox16.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If

            If Format(energ / (EnergJelovnika)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox39.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox39.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox39.Items.Insert(0, ">")
            End If

            'JUTARNJA UŽINA
            DGV = .DataGridView9
            energ = 0
            uglj = 0
            bjel = 0
            mast = 0
            .ListBox58.Items.Clear()   'ukupno (g)
            .ListBox59.Items.Clear()   'ukupno (%)
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox58.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox58.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox58.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox58.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox59.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox59.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox59.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox59.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox60.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            If .CheckBox5.Checked = False Then    'popodnevna uzina
                EnergMin = 10
                EnergMax = 15
                .ListBox57.Items.Clear()
                .ListBox57.Items.Add(EnergMin & "-" & EnergMax & "%")
            Else
                EnergMin = 5        'ostali obroci
                EnergMax = 10
                .ListBox57.Items.Clear()
                .ListBox57.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If

            If Format(energ / (EnergJelovnika * K)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox60.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox60.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox60.Items.Insert(0, ">")
            End If

            'RUČAK
            DGV = .DataGridView11
            energ = 0
            uglj = 0
            bjel = 0
            mast = 0
            .ListBox63.Items.Clear()   'ukupno (g)
            .ListBox64.Items.Clear()   'ukupno (%)
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox63.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox63.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox63.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox63.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox64.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox64.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox64.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox64.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox65.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            If .CheckBox3.Checked = False And .CheckBox5.Checked = False And .CheckBox7.Checked = False Then    'jut. uzina, popodnevna uzina i obrok pred spavanje
                EnergMin = 35
                EnergMax = 45
                .ListBox62.Items.Clear()
                .ListBox62.Items.Add(EnergMin & "-" & EnergMax & "%")
            Else
                EnergMin = 30        'ostali obroci
                EnergMax = 40
                .ListBox62.Items.Clear()
                .ListBox62.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If

            If Format(energ / (EnergJelovnika * K)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox65.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox65.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox65.Items.Insert(0, ">")
            End If

            'POPODNEVNA UŽINA
            DGV = .DataGridView12
            energ = 0
            uglj = 0
            bjel = 0
            mast = 0
            .ListBox68.Items.Clear()   'ukupno (g)
            .ListBox69.Items.Clear()   'ukupno (%)
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox68.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox68.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox68.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox68.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox69.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox69.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox69.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox69.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox70.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            EnergMin = 5
            EnergMax = 10
            .ListBox67.Items.Clear()
            .ListBox67.Items.Add(EnergMin & "-" & EnergMax & "%")

            If Format(energ / (EnergJelovnika * K)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox70.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox70.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox70.Items.Insert(0, ">")
            End If

            'VEČERA
            DGV = .DataGridView13
            energ = 0
            uglj = 0
            bjel = 0
            mast = 0
            .ListBox73.Items.Clear()   'ukupno (g)
            .ListBox74.Items.Clear()   'ukupno (%)
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox73.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox73.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox73.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox73.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox74.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox74.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox74.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox74.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox75.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            If .CheckBox2.Checked = True And .CheckBox3.Checked = True And .CheckBox4.Checked = True And .CheckBox5.Checked = True And .CheckBox6.Checked = True And .CheckBox7.Checked = True Then    'svi obroci
                EnergMin = 20
                EnergMax = 23
                .ListBox72.Items.Clear()
                .ListBox72.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If
            If .CheckBox2.Checked = True And .CheckBox3.Checked = True And .CheckBox4.Checked = True And .CheckBox5.Checked = True And .CheckBox6.Checked = True And .CheckBox7.Checked = False Then
                EnergMin = 20
                EnergMax = 25
                .ListBox72.Items.Clear()
                .ListBox72.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If
            If .CheckBox2.Checked = True And .CheckBox3.Checked = True And .CheckBox4.Checked = True And .CheckBox5.Checked = False And .CheckBox6.Checked = True And .CheckBox7.Checked = False Then
                EnergMin = 25
                EnergMax = 30
                .ListBox72.Items.Clear()
                .ListBox72.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If
            If .CheckBox2.Checked = True And .CheckBox3.Checked = False And .CheckBox4.Checked = True And .CheckBox5.Checked = True And .CheckBox6.Checked = True And .CheckBox7.Checked = False Then
                EnergMin = 20
                EnergMax = 25
                .ListBox72.Items.Clear()
                .ListBox72.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If
            If .CheckBox2.Checked = True And .CheckBox3.Checked = False And .CheckBox4.Checked = True And .CheckBox5.Checked = False And .CheckBox6.Checked = True And .CheckBox7.Checked = False Then
                EnergMin = 25
                EnergMax = 30
                .ListBox72.Items.Clear()
                .ListBox72.Items.Add(EnergMin & "-" & EnergMax & "%")
            End If
            '         Else
            '            EnergMin = 30        'ostali obroci
            '           EnergMax = 40
            '          .ListBox62.Items.Clear()
            '         .ListBox62.Items.Add(EnergMin & "-" & EnergMax & "%")


            If Format(energ / (EnergJelovnika * K)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox75.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox75.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox75.Items.Insert(0, ">")
            End If

            'OBROK PRED SPAVANJE
            DGV = .DataGridView14
            energ = 0
            uglj = 0
            bjel = 0
            mast = 0
            .ListBox78.Items.Clear()   'ukupno (g)
            .ListBox79.Items.Clear()   'ukupno (%)
            For i = 0 To DGV.RowCount - 1
                'nutrijenti
                energ = energ + DGV.Rows(i).Cells(11).Value
                uglj = uglj + DGV.Rows(i).Cells(13).Value
                bjel = bjel + DGV.Rows(i).Cells(14).Value
                mast = mast + DGV.Rows(i).Cells(15).Value
            Next i
            'nutrijenti - ukupno
            .ListBox78.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox78.Items.Insert(1, Format(uglj, "0") & " g")
            .ListBox78.Items.Insert(2, Format(bjel, "0") & " g")
            .ListBox78.Items.Insert(3, Format(mast, "0") & " g")
            ' .ListBox8.Items.Insert(3, "")   'prazno polje
            'nutrijenti - postoci
            .ListBox79.Items.Insert(0, Format((energ / (EnergJelovnika * K)) * 100, "0") & "%")
            .ListBox79.Items.Insert(1, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox79.Items.Insert(2, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox79.Items.Insert(3, Format((mast * 9 / energ) * 100, "0") & "%")
            '.ListBox8.Items.Insert(3, "")   'prazno polje

            'OK, > , < - ENERGIJA
            .ListBox80.Items.Clear()
            'Energija

            'Preporucena energetska vrijednost obroka
            EnergMin = 2
            EnergMax = 5
            .ListBox77.Items.Clear()
            .ListBox77.Items.Add(EnergMin & "-" & EnergMax & "%")

            If Format(energ / (EnergJelovnika * K)) * 100 >= EnergMin And Format(energ / (EnergJelovnika * K)) * 100 <= EnergMax Then
                .ListBox80.Items.Insert(0, "OK")
            Else
                If Format(energ / (EnergJelovnika * K)) * 100 < EnergMin Then .ListBox80.Items.Insert(0, "<")
                If Format(energ / (EnergJelovnika * K)) * 100 > EnergMax Then .ListBox80.Items.Insert(0, ">")
            End If

        End With
    End Sub
End Module
