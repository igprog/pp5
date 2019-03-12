Module ParametriPreporuceneVrijednostiModul
    Sub ParametriPreporuceneVrijednosti()
        On Error Resume Next
        With Form1
            .ListBox26.Items.Clear()    'UI
            .ListBox29.Items.Clear()    'RDA
            .ListBox30.Items.Clear()    'UI
            .ListBox36.Items.Clear()  'Ostale namirnice
            .ListBox27.Items.Clear()   'MDA (natrij, kalij, klor)
            .ListBox28.Items.Clear()   'UI (natrij)

            Dim Dob As Double = Val(.ComboBox1.Text)   'Dob

            Dim Energija10Posto As Integer = (.TextBox3.Text * 0.1) / 9
            Dim Energija2Posto As Integer = (.TextBox3.Text * 0.02) / 9
            Dim Energija7Posto As Integer = (.TextBox3.Text * 0.7) / 9
            Dim Energija05Posto As Integer = (.TextBox3.Text * 0.005) / 9
            Dim Energija5Posto As Integer = (.TextBox3.Text * 0.05) / 9
            Dim Energija1Posto As Integer = (.TextBox3.Text * 0.01) / 9
            Dim Energija15Posto As Integer = (.TextBox3.Text * 0.15) / 9
            Dim Energija20Posto As Integer = (.TextBox3.Text * 0.2) / 9
            Dim Energija6Posto As Integer = (.TextBox3.Text * 0.06) / 9
            Dim Energija11Posto As Integer = (.TextBox3.Text * 0.11) / 9
            Dim Energija8Posto As Integer = (.TextBox3.Text * 0.08) / 9


            'PUNOLJETNE OSOBE
            If Dob >= 18 Then
                'Loši parametri - UI najveći prihvatljiv dnevni unos
                If .Label286.Text <> "" Then
                    .ListBox26.Items.Insert(0, Energija7Posto) 'zasićene masti - hipolipemička dijeta (7%)
                Else
                    .ListBox26.Items.Insert(0, Energija10Posto)   'zasićene masti
                End If

                .ListBox26.Items.Insert(1, Energija2Posto)   'trans masne kiseline
                .ListBox26.Items.Insert(2, 300)   'kolesterol

                .ListBox27.Items.Insert(0, 500)  ' Natrij - MDA
                .ListBox27.Items.Insert(1, 2000)  ' Kalij - MDA
                .ListBox27.Items.Insert(2, 800)  ' Klor - MDA

                .ListBox28.Items.Insert(0, 2400)  'Natrij - UI

                .ListBox36.Items.Insert(0, Format(Val(.Label284.Text), "0"))  'Ostale namirnice - UI

                'RDA- -ostali mikronutrijenti
                .ListBox29.Items.Insert(0, 25)  'vlakna
                .ListBox29.Items.Insert(1, Energija15Posto)   'jednostruko nezasićene masti
                .ListBox29.Items.Insert(2, Energija8Posto)   'visestruko nezasicene masti
                .ListBox29.Items.Insert(3, 800)  'kalcij
                .ListBox29.Items.Insert(4, 375)  'magnezij
                .ListBox29.Items.Insert(5, 700)  'fosfor
                .ListBox29.Items.Insert(6, 14)  'zeljezo
                .ListBox29.Items.Insert(7, 1)  'bakar
                .ListBox29.Items.Insert(8, 10)  'cink
                .ListBox29.Items.Insert(9, 2)  'mangan
                .ListBox29.Items.Insert(10, 55)  'selen
                .ListBox29.Items.Insert(11, 150)  'jod
                .ListBox29.Items.Insert(12, 800)  'retinol
                .ListBox29.Items.Insert(13, 5)  'vitamin D
                .ListBox29.Items.Insert(14, 12)  'vitamin E
                .ListBox29.Items.Insert(15, 1.1)  'vitamin B1
                .ListBox29.Items.Insert(16, 1.4)  'vitamin B2
                .ListBox29.Items.Insert(17, 16)  'vitamin B3
                .ListBox29.Items.Insert(18, 1.4)  'vitamin B6
                .ListBox29.Items.Insert(19, 2.5)  'vitamin B12
                .ListBox29.Items.Insert(20, 200)  'folat
                .ListBox29.Items.Insert(21, 6)  'pantetonska kiselina
                .ListBox29.Items.Insert(22, 50)  'biotin
                .ListBox29.Items.Insert(23, 80)  'vitamin C
                .ListBox29.Items.Insert(24, 75)  'vitamin K

                'UI - Najveci dopusteni dnevni unos - Ostali mikronutrijenti
                .ListBox30.Items.Insert(0, "")  'vlakna
                .ListBox30.Items.Insert(1, Energija20Posto)   'jednostruko nezasićene masti
                .ListBox30.Items.Insert(2, Energija11Posto)   'visestruko nezasicene masti
                .ListBox30.Items.Insert(3, 1500)  'kalcij
                .ListBox30.Items.Insert(4, 700)  'magnezij
                .ListBox30.Items.Insert(5, 1400)  'fosfor
                .ListBox30.Items.Insert(6, 30)  'zeljezo
                .ListBox30.Items.Insert(7, 3)  'bakar
                .ListBox30.Items.Insert(8, 15)  'cink
                .ListBox30.Items.Insert(9, 4)  'mangan
                .ListBox30.Items.Insert(10, 100)  'selen
                .ListBox30.Items.Insert(11, 225)  'jod
                .ListBox30.Items.Insert(12, 1500)  'retinol
                .ListBox30.Items.Insert(13, 10)  'vitamin D
                .ListBox30.Items.Insert(14, 100)  'vitamin E
                .ListBox30.Items.Insert(15, 4)  'vitamin B1
                .ListBox30.Items.Insert(16, 4)  'vitamin B2
                .ListBox30.Items.Insert(17, 35)  'vitamin B3
                .ListBox30.Items.Insert(18, 6)  'vitamin B6
                .ListBox30.Items.Insert(19, 9)  'vitamin B12
                .ListBox30.Items.Insert(20, 600)  'folat
                .ListBox30.Items.Insert(21, 15)  'pantetonska kiselina
                .ListBox30.Items.Insert(22, 100)  'biotin
                .ListBox30.Items.Insert(23, 500)  'vitamin C
                .ListBox30.Items.Insert(24, 100)  'vitamin K
            End If

            'DJECA - 9-10 GOD
            If Dob >= 9 And Dob < 10 Then
                'Loši parametri - UI najveći prihvatljiv dnevni unos
                If .Label286.Text <> "" Then
                    .ListBox26.Items.Insert(0, Energija7Posto) 'zasićene masti - hipolipemička dijeta (7%)
                Else
                    .ListBox26.Items.Insert(0, Energija10Posto)   'zasićene masti
                End If

                .ListBox26.Items.Insert(1, Energija1Posto)   'trans masne kiseline
                .ListBox26.Items.Insert(2, 300)   'kolesterol

                .ListBox27.Items.Insert(0, 1580)  ' Natrij - MDA
                .ListBox27.Items.Insert(1, 3800)  ' Kalij - MDA
                .ListBox27.Items.Insert(2, 690)  ' Klor - MDA

                ' .ListBox28.Items.Insert(0, 24)  'Natrij - UI

                .ListBox36.Items.Insert(0, Format(Val(.Label284.Text), "0"))  'Ostale namirnice - UI()

                'RDA- -ostali mikronutrijenti
                .ListBox29.Items.Insert(0, 20)  'vlakna
                .ListBox29.Items.Insert(1, Energija15Posto)   'jednostruko nezasićene masti
                .ListBox29.Items.Insert(2, Energija8Posto)   'visestruko nezasicene masti
                .ListBox29.Items.Insert(3, 900)  'kalcij
                .ListBox29.Items.Insert(4, 170)  'magnezij
                .ListBox29.Items.Insert(5, 800)  'fosfor
                .ListBox29.Items.Insert(6, 10)  'zeljezo
                .ListBox29.Items.Insert(7, 1.5)  'bakar
                .ListBox29.Items.Insert(8, 7)  'cink
                .ListBox29.Items.Insert(9, 3)  'mangan
                .ListBox29.Items.Insert(10, 50)  'selen
                .ListBox29.Items.Insert(11, 130)  'jod
                .ListBox29.Items.Insert(12, 800)  'retinol
                .ListBox29.Items.Insert(13, 5)  'vitamin D
                .ListBox29.Items.Insert(14, 9.5)  'vitamin E
                .ListBox29.Items.Insert(15, 1)  'vitamin B1
                .ListBox29.Items.Insert(16, 1.1)  'vitamin B2
                .ListBox29.Items.Insert(17, 12)  'vitamin B3
                .ListBox29.Items.Insert(18, 0.7)  'vitamin B6
                .ListBox29.Items.Insert(19, 1.8)  'vitamin B12
                .ListBox29.Items.Insert(20, 300)  'folat
                .ListBox29.Items.Insert(21, 5)  'pantetonska kiselina
                .ListBox29.Items.Insert(22, 20)  'biotin
                .ListBox29.Items.Insert(23, 80)  'vitamin C
                .ListBox29.Items.Insert(24, 30)  'vitamin K

                'UI - Najveci dopusteni dnevni unos - Ostali mikronutrijenti
                .ListBox30.Items.Insert(0, "")  'vlakna
                .ListBox30.Items.Insert(1, Energija20Posto)   'jednostruko nezasićene masti
                .ListBox30.Items.Insert(2, Energija11Posto)   'visestruko nezasicene masti
            End If

                'DJECA - 10-14 GOD
                If Dob >= 10 And Dob < 14 Then
                    'Loši parametri - UI najveći prihvatljiv dnevni unos
                If .Label286.Text <> "" Then
                    .ListBox26.Items.Insert(0, Energija7Posto) 'zasićene masti - hipolipemička dijeta (7%)
                Else
                    .ListBox26.Items.Insert(0, Energija10Posto)   'zasićene masti
                End If

                .ListBox26.Items.Insert(1, Energija1Posto)   'trans masne kiseline
                .ListBox26.Items.Insert(2, 300)   'kolesterol

                .ListBox27.Items.Insert(0, 1680)  ' Natrij - MDA
                .ListBox27.Items.Insert(1, 4500)  ' Kalij - MDA
                .ListBox27.Items.Insert(2, 770)  ' Klor - MDA

                .ListBox36.Items.Insert(0, Format(Val(.Label284.Text), "0"))  'Ostale namirnice - UI

                'RDA- -ostali mikronutrijenti
                .ListBox29.Items.Insert(0, 22)  'vlakna
                .ListBox29.Items.Insert(1, Energija15Posto)   'jednostruko nezasićene masti
                .ListBox29.Items.Insert(2, Energija8Posto)   'visestruko nezasicene masti
                .ListBox29.Items.Insert(3, 1100)  'kalcij
                .ListBox29.Items.Insert(4, 240)  'magnezij
                .ListBox29.Items.Insert(5, 1250)  'fosfor
                .ListBox29.Items.Insert(6, 13.5)  'zeljezo
                .ListBox29.Items.Insert(7, 1.5)  'bakar
                .ListBox29.Items.Insert(8, 8)  'cink
                .ListBox29.Items.Insert(9, 5)  'mangan
                .ListBox29.Items.Insert(10, 60)  'selen
                .ListBox29.Items.Insert(11, 150)  'jod
                .ListBox29.Items.Insert(12, 900)  'retinol
                .ListBox29.Items.Insert(13, 5)  'vitamin D
                .ListBox29.Items.Insert(14, 12)  'vitamin E
                .ListBox29.Items.Insert(15, 1.1)  'vitamin B1
                .ListBox29.Items.Insert(16, 1.3)  'vitamin B2
                .ListBox29.Items.Insert(17, 14)  'vitamin B3
                .ListBox29.Items.Insert(18, 1)  'vitamin B6
                .ListBox29.Items.Insert(19, 2)  'vitamin B12
                .ListBox29.Items.Insert(20, 400)  'folat
                .ListBox29.Items.Insert(21, 5)  'pantetonska kiselina
                .ListBox29.Items.Insert(22, 30)  'biotin
                .ListBox29.Items.Insert(23, 90)  'vitamin C
                .ListBox29.Items.Insert(24, 40)  'vitamin K

                'UI - Najveci dopusteni dnevni unos - Ostali mikronutrijenti
                .ListBox30.Items.Insert(0, "")  'vlakna
                .ListBox30.Items.Insert(1, Energija20Posto)   'jednostruko nezasićene masti
                .ListBox30.Items.Insert(2, Energija11Posto)   'visestruko nezasicene masti
            End If

                'DJECA - 14-18 GOD
                If Dob >= 14 And Dob < 18 Then
                    'Loši parametri - UI najveći prihvatljiv dnevni unos
                If .Label286.Text <> "" Then
                    .ListBox26.Items.Insert(0, Energija7Posto) 'zasićene masti - hipolipemička dijeta (7%)
                Else
                    .ListBox26.Items.Insert(0, Energija10Posto)   'zasićene masti
                End If

                .ListBox26.Items.Insert(1, Energija1Posto)   'trans masne kiseline
                .ListBox26.Items.Insert(2, 300)   'kolesterol

                .ListBox27.Items.Insert(0, 2000)  ' Natrij - MDA
                .ListBox27.Items.Insert(1, 4700)  ' Kalij - MDA
                .ListBox27.Items.Insert(2, 830)  ' Klor - MDA

                .ListBox36.Items.Insert(0, Format(Val(.Label284.Text), "0"))  'Ostale namirnice - UI

                'RDA- -ostali mikronutrijenti
                .ListBox29.Items.Insert(0, 25)  'vlakna
                .ListBox29.Items.Insert(1, Energija15Posto)   'jednostruko nezasićene masti
                .ListBox29.Items.Insert(2, Energija8Posto)   'visestruko nezasicene masti
                .ListBox29.Items.Insert(3, 1200)  'kalcij
                .ListBox29.Items.Insert(4, 342.5)  'magnezij
                .ListBox29.Items.Insert(5, 1250)  'fosfor
                .ListBox29.Items.Insert(6, 13.5)  'zeljezo
                .ListBox29.Items.Insert(7, 1.5)  'bakar
                .ListBox29.Items.Insert(8, 8.38)  'cink
                .ListBox29.Items.Insert(9, 5)  'mangan
                .ListBox29.Items.Insert(10, 65)  'selen
                .ListBox29.Items.Insert(11, 175)  'jod
                .ListBox29.Items.Insert(12, 1030)  'retinol
                .ListBox29.Items.Insert(13, 5)  'vitamin D
                .ListBox29.Items.Insert(14, 13.25)  'vitamin E
                .ListBox29.Items.Insert(15, 1.2)  'vitamin B1
                .ListBox29.Items.Insert(16, 1.4)  'vitamin B2
                .ListBox29.Items.Insert(17, 15.75)  'vitamin B3
                .ListBox29.Items.Insert(18, 1.4)  'vitamin B6
                .ListBox29.Items.Insert(19, 3)  'vitamin B12
                .ListBox29.Items.Insert(20, 400)  'folat
                .ListBox29.Items.Insert(21, 6)  'pantetonska kiselina
                .ListBox29.Items.Insert(22, 47.5)  'biotin
                .ListBox29.Items.Insert(23, 100)  'vitamin C
                .ListBox29.Items.Insert(24, 57.5)  'vitamin K

                'UI - Najveci dopusteni dnevni unos - Ostali mikronutrijenti
                .ListBox30.Items.Insert(0, "")  'vlakna
                .ListBox30.Items.Insert(1, Energija20Posto)   'jednostruko nezasićene masti
                .ListBox30.Items.Insert(2, Energija11Posto)   'visestruko nezasicene masti
            End If

        End With
    End Sub
End Module
