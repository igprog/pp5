Module BazaJelovnikaSpremiModul
    Sub BazaJelovnikaSpremi()
        On Error Resume Next
        With Form1
            Dim BS As BindingSource = .BazaJelovnikaBindingSource
            BS.MoveLast()
            If .DataGridView8.RowCount <= 1 Then BS.AddNew()
            Dim DGV As DataGridView
            Dim Obrok As String
            Dim ObrokBroj As Integer
            Dim PripremaJela As String

            Dim b As Integer
            For b = 1 To 6
                If b = 1 Then
                    DGV = .DataGridView5
                    '   Obrok = "DORUČAK"
                    Obrok = .TabPage7.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox11.Text
                End If
                If b = 2 Then
                    DGV = .DataGridView9
                    ' Obrok = "JUTARNJA UŽINA"
                    Obrok = .TabPage8.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox64.Text
                End If
                If b = 3 Then
                    DGV = .DataGridView11
                    ' Obrok = "RUČAK"
                    Obrok = .TabPage9.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox65.Text
                End If
                If b = 4 Then
                    DGV = .DataGridView12
                    ' Obrok = "POPODNEVNA UŽINA"
                    Obrok = .TabPage10.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox66.Text
                End If
                If b = 5 Then
                    DGV = .DataGridView13
                    ' Obrok = "VEČERA"
                    Obrok = .TabPage11.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox67.Text
                End If
                If b = 6 Then
                    DGV = .DataGridView14
                    'Obrok = "OBROK PRED SPAVANJE"
                    Obrok = .TabPage12.Text
                    ObrokBroj = b  'BrojObroka
                    PripremaJela = .TextBox68.Text
                End If

                For i = 0 To DGV.RowCount - 1
                    .Label108.Text = .TextBox1.Text & " " & .TextBox2.Text  'Korisnik
                    .Label109.Text = Val(.Label179.Text)  'BrojDijete
                    .Label110.Text = .Label21.Text 'NazivDijete
                    .Label111.Text = Date.Today             'DatumIzradeJelovnika
                    .Label112.Text = .DateTimePicker1.Value.Date 'DatumJelovnika
                    .Label113.Text = .TextBox13.Text  'NazivJelovnika
                    .Label114.Text = Val(.Label175.Text)  'Energetska vrijednost jelovnika
                    .Label115.Text = Obrok  'Obrok
                    .Label358.Text = ObrokBroj  'ObrokBroj
                    .Label116.Text = PripremaJela  'NazivPripremaJela
                    .Label119.Text = DGV.Rows(i).Cells(5).Value  'NazivNamirnice
                    .Label359.Text = DGV.Rows(i).Cells(6).Value  'TermickaObrada
                    .Label120.Text = DGV.Rows(i).Cells(7).Value  'Seriranja
                    .Label121.Text = DGV.Rows(i).Cells(8).Value  'Kolicina
                    .Label122.Text = DGV.Rows(i).Cells(9).Value  'Mjera
                    .Label123.Text = DGV.Rows(i).Cells(10).Value  'Masa_g
                    .Label124.Text = DGV.Rows(i).Cells(11).Value  'Energija_kcal
                    .Label125.Text = DGV.Rows(i).Cells(12).Value  'Energija_kJ
                    .Label126.Text = DGV.Rows(i).Cells(13).Value  'Ugljikohidrati_g
                    .Label127.Text = DGV.Rows(i).Cells(14).Value  'Bjelancevine_g
                    .Label128.Text = DGV.Rows(i).Cells(15).Value  'Masti_g
                    .Label129.Text = DGV.Rows(i).Cells(16).Value  'Skrob
                    .Label130.Text = DGV.Rows(i).Cells(17).Value  'UkupniSeceri
                    .Label131.Text = DGV.Rows(i).Cells(18).Value  'Glukoza
                    .Label132.Text = DGV.Rows(i).Cells(19).Value  'Fruktoza
                    .Label133.Text = DGV.Rows(i).Cells(20).Value  'Saharoza
                    .Label134.Text = DGV.Rows(i).Cells(21).Value  'Maltoza
                    .Label135.Text = DGV.Rows(i).Cells(22).Value  'Laktoza
                    .Label136.Text = DGV.Rows(i).Cells(23).Value  'Vlakna
                    .Label137.Text = DGV.Rows(i).Cells(24).Value  'ZasiceneMasti
                    .Label138.Text = DGV.Rows(i).Cells(25).Value  'JednostrukoNezasiceneMasti
                    .Label139.Text = DGV.Rows(i).Cells(26).Value  'VisestrukoNezasiceneMasti
                    .Label140.Text = DGV.Rows(i).Cells(27).Value  'TransMasneKiseline
                    .Label141.Text = DGV.Rows(i).Cells(28).Value  'Kolesterol
                    .Label142.Text = DGV.Rows(i).Cells(29).Value  'Natrij
                    .Label143.Text = DGV.Rows(i).Cells(30).Value  'Kalij
                    .Label144.Text = DGV.Rows(i).Cells(31).Value  'Kalcij
                    .Label145.Text = DGV.Rows(i).Cells(32).Value  'Magnezij
                    .Label146.Text = DGV.Rows(i).Cells(33).Value  'Fosfor
                    .Label147.Text = DGV.Rows(i).Cells(34).Value  'Zeljezo
                    .Label148.Text = DGV.Rows(i).Cells(35).Value  'Bakar
                    .Label149.Text = DGV.Rows(i).Cells(36).Value  'Cink
                    .Label150.Text = DGV.Rows(i).Cells(37).Value  'Klor
                    .Label151.Text = DGV.Rows(i).Cells(38).Value  'Mangan
                    .Label152.Text = DGV.Rows(i).Cells(39).Value  'Selen
                    .Label153.Text = DGV.Rows(i).Cells(40).Value  'Jod
                    .Label154.Text = DGV.Rows(i).Cells(41).Value  'Retinol
                    .Label155.Text = DGV.Rows(i).Cells(42).Value  'Karoten
                    .Label156.Text = DGV.Rows(i).Cells(43).Value  'VitaminD
                    .Label157.Text = DGV.Rows(i).Cells(44).Value  'VitaminE
                    .Label158.Text = DGV.Rows(i).Cells(45).Value  'VitaminB1
                    .Label159.Text = DGV.Rows(i).Cells(46).Value  'VitaminB2
                    .Label160.Text = DGV.Rows(i).Cells(47).Value  'VitaminB3
                    .Label161.Text = DGV.Rows(i).Cells(48).Value  'VitaminB6
                    .Label162.Text = DGV.Rows(i).Cells(49).Value  'VitaminB12
                    .Label163.Text = DGV.Rows(i).Cells(50).Value  'Folat
                    .Label164.Text = DGV.Rows(i).Cells(51).Value  'PantotenskaKiselina
                    .Label165.Text = DGV.Rows(i).Cells(52).Value  'Biotin
                    .Label166.Text = DGV.Rows(i).Cells(53).Value  'VitaminC
                    .Label167.Text = DGV.Rows(i).Cells(54).Value  'VitaminK
                    .Label168.Text = DGV.Rows(i).Cells(55).Value  'Zitarice
                    .Label169.Text = DGV.Rows(i).Cells(56).Value  'Povrce
                    .Label170.Text = DGV.Rows(i).Cells(57).Value  'Voce
                    .Label171.Text = DGV.Rows(i).Cells(58).Value  'Meso
                    .Label172.Text = DGV.Rows(i).Cells(59).Value  'Mlijeko
                    .Label173.Text = DGV.Rows(i).Cells(60).Value  'Masti
                    .Label174.Text = DGV.Rows(i).Cells(61).Value  'OstaleNamirnice
                    .Label360.Text = DGV.Rows(i).Cells(62).Value  'Cijena

                    If DGV.Rows(i).Cells(5).Value.ToString <> "" Then
                        BS.AddNew()
                    End If

                Next i
            Next b

        End With
    End Sub
End Module
