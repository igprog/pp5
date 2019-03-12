Module IspisModul
    Sub Ispis()
        With Form1
            On Error Resume Next
            .TextBox70.Text = ""
            Dim DGV As DataGridView = .DataGridView5
            Dim Obrok As String = ""
            Dim PripremaJela As String = ""
            Dim UkupnObrok As ListBox
            Dim UkupnObrokPost As ListBox
            Dim Namirnice As String = "  Namirnice"
            Dim K As Integer = .ComboBox16.Text     'Broj korisnika jelovnika
            Dim BrojKorisnika As String = ""
            Dim CijenaUkupno As String = ""
            Dim Crta As String = "-------------------------------------------------------------------------------------------------------------------------------------------------------------------"
            Dim Prazno As String = "~"

            .TextBox69.Text = "Jelovnik je izrađen pomoću računalnog programa PROGRAM PREHRANE 5.0" _
                    & vbCrLf & "www.programprehrane.com" _
                    & vbCrLf & Crta _
                    & vbCrLf & vbCrLf & "Klijent: " & .TextBox1.Text & " " & .TextBox2.Text & BrojKorisnika _
                    & vbCrLf & .Label21.Text & vbCrLf & .TextBox13.Text & vbCrLf & Crta
            ' & vbCrLf & "_______________________________________________________________________"
            ' & vbCrLf & "-----------------------------------------------------------------------"

            Dim b As Integer
            For b = 1 To 6
                If b = 1 Then
                    DGV = .DataGridView5
                    If DGV.RowCount > 2 Then
                        'Obrok = "DORUČAK"
                        Obrok = .TabPage7.Text
                        PripremaJela = .TextBox11.Text
                        UkupnObrok = .ListBox8
                        UkupnObrokPost = .ListBox12
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox8.Items.Clear()
                        .ListBox12.Items.Clear()
                        UkupnObrok = .ListBox8
                        UkupnObrokPost = .ListBox12
                        Namirnice = ""
                    End If
                End If
                If b = 2 Then
                    DGV = .DataGridView9
                    If DGV.RowCount >= 2 Then
                        ' Obrok = "JUTARNJA UŽINA" 
                        Obrok = .TabPage8.Text
                        PripremaJela = .TextBox64.Text
                        UkupnObrok = .ListBox58
                        UkupnObrokPost = .ListBox59
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox58.Items.Clear()
                        .ListBox59.Items.Clear()
                        UkupnObrok = .ListBox58
                        UkupnObrokPost = .ListBox59
                        Namirnice = ""
                    End If
                End If
                If b = 3 Then
                    DGV = .DataGridView11
                    If DGV.RowCount > 2 Then
                        ' Obrok = "RUČAK"
                        Obrok = .TabPage9.Text
                        PripremaJela = .TextBox65.Text
                        UkupnObrok = .ListBox63
                        UkupnObrokPost = .ListBox64
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox63.Items.Clear()
                        .ListBox64.Items.Clear()
                        UkupnObrok = .ListBox63
                        UkupnObrokPost = .ListBox64
                        Namirnice = ""
                    End If
                End If
                If b = 4 Then
                    DGV = .DataGridView12
                    If DGV.RowCount > 2 Then
                        ' Obrok = "POPODNEVNA UŽINA"
                        Obrok = .TabPage10.Text
                        PripremaJela = .TextBox66.Text
                        UkupnObrok = .ListBox68
                        UkupnObrokPost = .ListBox69
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox68.Items.Clear()
                        .ListBox69.Items.Clear()
                        UkupnObrok = .ListBox68
                        UkupnObrokPost = .ListBox69
                        Namirnice = ""
                    End If
                End If
                If b = 5 Then
                    DGV = .DataGridView13
                    If DGV.RowCount > 2 Then
                        'Obrok = "VEČERA"
                        Obrok = .TabPage11.Text
                        PripremaJela = .TextBox67.Text
                        UkupnObrok = .ListBox73
                        UkupnObrokPost = .ListBox74
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox73.Items.Clear()
                        .ListBox74.Items.Clear()
                        UkupnObrok = .ListBox73
                        UkupnObrokPost = .ListBox74
                        Namirnice = ""
                    End If
                End If
                If b = 6 Then
                    DGV = .DataGridView14
                    If DGV.RowCount > 2 Then
                        'Obrok = "OBROK PRED SPAVANJE"
                        Obrok = .TabPage12.Text
                        PripremaJela = .TextBox68.Text
                        UkupnObrok = .ListBox78
                        UkupnObrokPost = .ListBox79
                        Namirnice = "  Namirnice:"
                    Else
                        Obrok = ""
                        PripremaJela = ""
                        .ListBox78.Items.Clear()
                        .ListBox79.Items.Clear()
                        UkupnObrok = .ListBox78
                        UkupnObrokPost = .ListBox79
                        Namirnice = ""
                    End If
                End If

                .ListBox48.Items.Clear()     'briši listbox
                .ListBox49.Items.Clear()
                .ListBox50.Items.Clear()
                .ListBox51.Items.Clear()
                .ListBox81.Items.Clear()

                For i = 0 To DGV.RowCount - 1
                    .ListBox48.Items.Add(DGV.Rows(i).Cells(5).Value)  'namirnica
                    If .CheckBox8.Checked = True Then
                        If DGV.Rows(i).Cells(8).Value * K >= 10 Or DGV.Rows(i).Cells(8).Value * K = 1 Or DGV.Rows(i).Cells(8).Value * K = 2 Or DGV.Rows(i).Cells(8).Value * K = 3 Or DGV.Rows(i).Cells(8).Value * K = 4 Or DGV.Rows(i).Cells(8).Value * K = 5 Or DGV.Rows(i).Cells(8).Value * K = 6 Or DGV.Rows(i).Cells(8).Value * K = 7 Or DGV.Rows(i).Cells(8).Value * K = 8 Or DGV.Rows(i).Cells(8).Value * K = 9 Then
                            .ListBox49.Items.Add(", " & Format(DGV.Rows(i).Cells(8).Value * K, "0") & " " & (DGV.Rows(i).Cells(9).Value))  'kolicina i mjera
                        Else
                            .ListBox49.Items.Add(", " & Format(DGV.Rows(i).Cells(8).Value * K, "0.0") & " " & (DGV.Rows(i).Cells(9).Value))  'kolicina i mjera
                        End If
                    Else
                        .ListBox49.Items.Add("")
                    End If
                    '        .ListBox49.Items.Add(DGV.Rows(i).Cells(8).Value)  'kolicina
                    ' .ListBox50.Items.Add(DGV.Rows(i).Cells(9).Value)  'mjera
                    If .CheckBox9.Checked = True Then
                        If DGV.Rows(i).Cells(10).Value * K >= 10 Or DGV.Rows(i).Cells(10).Value * K = 1 Or DGV.Rows(i).Cells(10).Value * K = 2 Or DGV.Rows(i).Cells(10).Value * K = 3 Or DGV.Rows(i).Cells(10).Value * K = 4 Or DGV.Rows(i).Cells(10).Value * K = 5 Or DGV.Rows(i).Cells(10).Value * K = 6 Or DGV.Rows(i).Cells(10).Value * K = 7 Or DGV.Rows(i).Cells(10).Value * K = 8 Or DGV.Rows(i).Cells(10).Value * K = 9 Then
                            .ListBox81.Items.Add(", " & Format(DGV.Rows(i).Cells(10).Value * K, "0") & " g") 'masa (g)
                        Else
                            .ListBox81.Items.Add(", " & Format(DGV.Rows(i).Cells(10).Value * K, "0.0") & " g") 'masa (g)
                        End If

                    Else
                        .ListBox81.Items.Add("")
                    End If
                        ' .ListBox81.Items.Add(DGV.Rows(i).Cells(10).Value)  'masa (g)

                        Dim ZitariceServ As String = ""
                        Dim PovrceServ As String = ""
                        Dim VoceServ As String = ""
                        Dim MesoServ As String = ""
                        Dim MlijekoServ As String = ""
                        Dim MastiServ As String = ""
                        Dim OstaleNamirniceServ As String = ""

                    If DGV.Rows(i).Cells(55).Value <> 0 Then
                        If DGV.Rows(i).Cells(55).Value * K >= 10 Or DGV.Rows(i).Cells(55).Value * K = 1 Or DGV.Rows(i).Cells(55).Value * K = 2 Or DGV.Rows(i).Cells(55).Value * K = 3 Or DGV.Rows(i).Cells(55).Value * K = 4 Or DGV.Rows(i).Cells(55).Value * K = 5 Or DGV.Rows(i).Cells(55).Value * K = 6 Or DGV.Rows(i).Cells(55).Value * K = 7 Or DGV.Rows(i).Cells(55).Value * K = 8 Or DGV.Rows(i).Cells(55).Value * K = 9 Then
                            ZitariceServ = ", " & Format(DGV.Rows(i).Cells(55).Value * K, "0") & " serv. žitarica i proizvoda od žita"
                        Else
                            ZitariceServ = ", " & Format(DGV.Rows(i).Cells(55).Value * K, "0.0") & " serv. žitarica i proizvoda od žita"
                        End If
                    End If

                    If DGV.Rows(i).Cells(56).Value <> 0 Then
                        If DGV.Rows(i).Cells(56).Value * K >= 10 Or DGV.Rows(i).Cells(56).Value * K = 1 Or DGV.Rows(i).Cells(56).Value * K = 2 Or DGV.Rows(i).Cells(56).Value * K = 3 Or DGV.Rows(i).Cells(56).Value * K = 4 Or DGV.Rows(i).Cells(56).Value * K = 5 Or DGV.Rows(i).Cells(56).Value * K = 6 Or DGV.Rows(i).Cells(56).Value * K = 7 Or DGV.Rows(i).Cells(56).Value * K = 8 Or DGV.Rows(i).Cells(56).Value * K = 9 Then
                            PovrceServ = ", " & Format(DGV.Rows(i).Cells(56).Value * K, "0") & " serv. povrća"
                        Else
                            PovrceServ = ", " & Format(DGV.Rows(i).Cells(56).Value * K, "0.0") & " serv. povrća"
                        End If
                    End If

                    If DGV.Rows(i).Cells(57).Value <> 0 Then
                        If DGV.Rows(i).Cells(57).Value * K >= 10 Or DGV.Rows(i).Cells(57).Value * K = 1 Or DGV.Rows(i).Cells(57).Value * K = 2 Or DGV.Rows(i).Cells(57).Value * K = 3 Or DGV.Rows(i).Cells(57).Value * K = 4 Or DGV.Rows(i).Cells(57).Value * K = 5 Or DGV.Rows(i).Cells(57).Value * K = 6 Or DGV.Rows(i).Cells(57).Value * K = 7 Or DGV.Rows(i).Cells(57).Value * K = 8 Or DGV.Rows(i).Cells(57).Value * K = 9 Then
                            VoceServ = ", " & Format(DGV.Rows(i).Cells(57).Value * K, "0") & " serv. voća"
                        Else
                            VoceServ = ", " & Format(DGV.Rows(i).Cells(57).Value * K, "0.0") & " serv. voća"
                        End If
                        If DGV.Rows(i).Cells(57).Value * K < 0.1 Then
                            VoceServ = ", " & Format(DGV.Rows(i).Cells(57).Value * K, "0.000") & " serv. voća"   'limunov sok 0.006
                        End If

                    End If
                    If DGV.Rows(i).Cells(58).Value <> 0 Then
                        If DGV.Rows(i).Cells(58).Value * K >= 10 Or DGV.Rows(i).Cells(58).Value * K = 1 Or DGV.Rows(i).Cells(58).Value * K = 2 Or DGV.Rows(i).Cells(58).Value * K = 3 Or DGV.Rows(i).Cells(58).Value * K = 4 Or DGV.Rows(i).Cells(58).Value * K = 5 Or DGV.Rows(i).Cells(58).Value * K = 6 Or DGV.Rows(i).Cells(58).Value * K = 7 Or DGV.Rows(i).Cells(58).Value * K = 8 Or DGV.Rows(i).Cells(58).Value * K = 9 Then
                            MesoServ = ", " & Format(DGV.Rows(i).Cells(58).Value * K, "0") & " serv. mesa i zamjena"
                        Else
                            MesoServ = ", " & Format(DGV.Rows(i).Cells(58).Value * K, "0.0") & " serv. mesa i zamjena"
                        End If
                    End If

                    If DGV.Rows(i).Cells(59).Value <> 0 Then
                        If DGV.Rows(i).Cells(59).Value * K >= 10 Or DGV.Rows(i).Cells(59).Value * K = 1 Or DGV.Rows(i).Cells(59).Value * K = 2 Or DGV.Rows(i).Cells(59).Value * K = 3 Or DGV.Rows(i).Cells(59).Value * K = 4 Or DGV.Rows(i).Cells(59).Value * K = 5 Or DGV.Rows(i).Cells(59).Value * K = 6 Or DGV.Rows(i).Cells(59).Value * K = 7 Or DGV.Rows(i).Cells(59).Value * K = 8 Or DGV.Rows(i).Cells(59).Value * K = 9 Then
                            MlijekoServ = ", " & Format(DGV.Rows(i).Cells(59).Value * K, "0") & " serv. mlijeka i mliječnih proizvoda"
                        Else
                            MlijekoServ = ", " & Format(DGV.Rows(i).Cells(59).Value * K, "0.0") & " serv. mlijeka i mliječnih proizvoda"
                        End If
                    End If

                    If DGV.Rows(i).Cells(60).Value <> 0 Then
                        If DGV.Rows(i).Cells(60).Value * K >= 10 Or DGV.Rows(i).Cells(60).Value * K = 1 Or DGV.Rows(i).Cells(60).Value * K = 2 Or DGV.Rows(i).Cells(60).Value * K = 3 Or DGV.Rows(i).Cells(60).Value * K = 4 Or DGV.Rows(i).Cells(60).Value * K = 5 Or DGV.Rows(i).Cells(60).Value * K = 6 Or DGV.Rows(i).Cells(60).Value * K = 7 Or DGV.Rows(i).Cells(60).Value * K = 8 Or DGV.Rows(i).Cells(60).Value * K = 9 Then
                            MastiServ = ", " & Format(DGV.Rows(i).Cells(60).Value * K, "0") & " serv. masti"
                        Else
                            MastiServ = ", " & Format(DGV.Rows(i).Cells(60).Value * K, "0.0") & " serv. masti"
                        End If
                    End If

                    If DGV.Rows(i).Cells(61).Value <> 0 Then
                        If DGV.Rows(i).Cells(61).Value * K >= 10 Or DGV.Rows(i).Cells(61).Value * K = 1 Or DGV.Rows(i).Cells(61).Value * K = 2 Or DGV.Rows(i).Cells(61).Value * K = 3 Or DGV.Rows(i).Cells(61).Value * K = 4 Or DGV.Rows(i).Cells(61).Value * K = 5 Or DGV.Rows(i).Cells(61).Value * K = 6 Or DGV.Rows(i).Cells(61).Value * K = 7 Or DGV.Rows(i).Cells(61).Value * K = 8 Or DGV.Rows(i).Cells(61).Value * K = 9 Then
                            OstaleNamirniceServ = ", " & Format(DGV.Rows(i).Cells(61).Value * K, "0") & " serv. ostalih namirnica"
                        Else
                            OstaleNamirniceServ = ", " & Format(DGV.Rows(i).Cells(61).Value * K, "0.0") & " serv. ostalih namirnica"
                        End If
                    End If

                    If .CheckBox10.Checked = True Then
                        .ListBox51.Items.Add(ZitariceServ & PovrceServ & VoceServ & MesoServ & MlijekoServ & MastiServ & OstaleNamirniceServ)  'serviranja
                    Else
                        .ListBox51.Items.Add("")
                    End If
                    '  .ListBox51.Items.Add(ZitariceServ & PovrceServ & VoceServ & MesoServ & MlijekoServ & MastiServ & OstaleNamirniceServ)  'serviranja

                Next i

                If PripremaJela <> "" Then
                    .TextBox70.Text = .TextBox70.Text & vbCrLf & Prazno & vbCrLf & Obrok & vbCrLf & PripremaJela & _
   vbCrLf & Prazno & vbCrLf & Namirnice  'obrok (izvaden gornji red)
                    'obrok
                    ' vbCrLf & vbCrLf & "Namirnice:"  'obrok (izvaden gornji red)
                Else
                    .TextBox70.Text = .TextBox70.Text & vbCrLf & Prazno & vbCrLf & Obrok & _
                        vbCrLf & Prazno & vbCrLf & Namirnice  'obrok
                End If
                'naziv namirnice, količina, mjera, serviranja
                Dim a As Integer
                For a = 0 To .ListBox48.Items.Count - 1
                    If .ListBox48.Items(a).ToString <> "" Then
                        '        .TextBox70.Text = .TextBox70.Text & vbCrLf & "  - " & .ListBox48.Items(a) & " - (" & _
                        '             .ListBox49.Items(a) & " " & .ListBox50.Items(a) & ", " & .ListBox81.Items(a) & " g" & .ListBox51.Items(a) & ")"
                        .TextBox70.Text = .TextBox70.Text & vbCrLf & "  - " & .ListBox48.Items(a) & _
                             .ListBox49.Items(a) & .ListBox81.Items(a) & .ListBox51.Items(a)

                    End If
                Next

                'ukupne vrijednosti obroka
                If .CheckBox11.Checked = True Then
                    .TextBox70.Text = .TextBox70.Text & vbCrLf _
                   & Prazno & vbCrLf & "Energetska vrijednost obroka: " & UkupnObrok.Items(0) & " (" & UkupnObrokPost.Items(0) _
                   & "), ugljikohidrati: " & UkupnObrok.Items(1) & " (" & UkupnObrokPost.Items(1) _
                              & "), bjelančevine: " & UkupnObrok.Items(2) & " (" & UkupnObrokPost.Items(2) _
                   & "), masti: " & UkupnObrok.Items(3) & " (" & UkupnObrokPost.Items(3) & ")"
                End If

            Next


            'ukupne vrijednosti
            If .CheckBox12.Checked = True Then
                .TextBox70.Text = .TextBox70.Text & vbCrLf _
                    & vbCrLf & Crta _
               & vbCrLf & "UKUPNE VRIJEDNOSTI JELOVNIKA:" _
               & vbCrLf & "Energetska vrijednost jelovnika: " & .ListBox10.Items(0) & " (" & Format(.Label175.Text * 4.186, "0") & " kJ), ugljikohidrati: " & .ListBox10.Items(2) & " (" & .ListBox11.Items(2) _
               & "), bjelančevine: " & .ListBox10.Items(3) & " (" & .ListBox11.Items(3) & "), masti: " & .ListBox10.Items(4) & " (" & .ListBox11.Items(4) & ")"
            End If

            'Dodatna tjelesna aktivnost
            If .CheckBox13.Checked = True Then
                .ListBox52.Items.Clear()
                .ListBox53.Items.Clear()
                .ListBox54.Items.Clear()
                .TextBox73.Text = ""

                DGV = .DataGridView4
                For i = 0 To DGV.RowCount - 1
                    .ListBox52.Items.Add(DGV.Rows(i).Cells(1).Value)  'aktivnost
                    .ListBox53.Items.Add(DGV.Rows(i).Cells(6).Value)  'minuta
                    .ListBox54.Items.Add(DGV.Rows(i).Cells(7).Value)  'kcal
                Next i

                For a = 0 To .ListBox52.Items.Count - 1
                    If .ListBox52.Items(a).ToString <> "" Then
                        .TextBox73.Text = .TextBox73.Text & vbCrLf & "  " & .ListBox52.Items(a) & " - " & _
                            .ListBox53.Items(a) & " min (" & .ListBox54.Items(a) & " kcal)"
                    End If
                Next
            End If

            CijenaMinimalno()    'Modul (Cijena > )

            If K > 1 Then
                BrojKorisnika = ", Broj korisnika jelovnika: " & K
                CijenaUkupno = "  (Ukupna cijena za " & K & " korisnika: " & .Label338.Text & Format(.Label361.Text, "0.00") * K & " " & .ComboBox26.Text & ")"
            End If


            If .CheckBox13.Checked = True Then
                ' .TextBox83.Text = "Cijena jelovnika: " & .Label318.Text
                .TextBox70.Text = .TextBox70.Text & vbCrLf _
                  & Prazno & vbCrLf & "Cijena jelovnika: " & .Label338.Text & .Label318.Text & CijenaUkupno
            End If

            'ispis - sve (jelovnik i dodatna tjelesna aktivnost)
            If .CheckBox14.Checked = True Then
                If Val(.TextBox10.Text) > 0 Then
                    '             .TextBox69.Text = .TextBox69.Text & vbCrLf & .TextBox70.Text _
                    '               & vbCrLf & vbCrLf & vbCrLf & "DODATNA TJELESNA AKTIVNOST:" & .TextBox73.Text
                    .TextBox70.Text = .TextBox70.Text & vbCrLf _
                                   & Prazno & vbCrLf & "DODATNA TJELESNA AKTIVNOST:" & .TextBox73.Text & vbCrLf & Crta
                Else
                    ' .TextBox69.Text = .TextBox69.Text & vbCrLf & .TextBox70.Text
                End If
            End If
            .TextBox69.Text = .TextBox69.Text & vbCrLf & .TextBox70.Text   'ispis


            '.RichTextBox1.Clear()
            .RichTextBoxPrintCtrl1.Clear()
            '  .RichTextBox1.Text = .TextBox69.Text    'rich text box
            .RichTextBoxPrintCtrl1.Text = .TextBox69.Text    'rich text box

            'RICH TEXT BOX 1
            .RichTextBoxPrintCtrl1.Lines = .RichTextBoxPrintCtrl1.Text.Split(New Char() {ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)

            ' Dim Rtb As RichTextBox = .RichTextBox1
            Dim Rtb As RichTextBox = .RichTextBoxPrintCtrl1

          

            Dim Dorucak As String = .TabPage7.Text
            '  Rtb.Text = Replace(Rtb.Text, .TabPage7.Text, vbCrLf & .TabPage7.Text)
            Rtb.Select(Rtb.Text.IndexOf(Dorucak), Dorucak.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim JutarnjaUzina As String = .TabPage8.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage8.Text, vbCrLf & .TabPage8.Text)
            Rtb.Select(Rtb.Text.IndexOf(JutarnjaUzina), JutarnjaUzina.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim Rucak As String = .TabPage9.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage9.Text, vbCrLf & .TabPage9.Text)
            Rtb.Select(Rtb.Text.IndexOf(Rucak), Rucak.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim PopodnevnaUzina As String = .TabPage10.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage10.Text, vbCrLf & .TabPage10.Text)
            Rtb.Select(Rtb.Text.IndexOf(PopodnevnaUzina), PopodnevnaUzina.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim Vecera As String = .TabPage11.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage11.Text, vbCrLf & .TabPage11.Text)
            Rtb.Select(Rtb.Text.IndexOf(Vecera), Vecera.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim ObrokPredSpavanje As String = .TabPage12.Text
            '    Rtb.Text = Replace(Rtb.Text, .TabPage12.Text, vbCrLf & .TabPage12.Text)
            Rtb.Select(Rtb.Text.IndexOf(ObrokPredSpavanje), ObrokPredSpavanje.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim UkupneVrijednosti As String = "UKUPNE VRIJEDNOSTI JELOVNIKA:"
            '   Rtb.Text = Replace(Rtb.Text, "UKUPNE VRIJEDNOSTI JELOVNIKA:", vbCrLf & "UKUPNE VRIJEDNOSTI JELOVNIKA:")
            Rtb.Select(Rtb.Text.IndexOf(UkupneVrijednosti), UkupneVrijednosti.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim CijenaJelovnika As String = "Cijena jelovnika:"
            '   Rtb.Text = Replace(Rtb.Text, "Cijena jelovnika:", vbCrLf & "Cijena jelovnika:")
            Rtb.Select(Rtb.Text.IndexOf(CijenaJelovnika), CijenaJelovnika.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim DodatnaTjelesnaAktivnost As String = "DODATNA TJELESNA AKTIVNOST:"
            '  Rtb.Text = Replace(Rtb.Text, "DODATNA TJELESNA AKTIVNOST:", vbCrLf & "DODATNA TJELESNA AKTIVNOST:")
            Rtb.Select(Rtb.Text.IndexOf(DodatnaTjelesnaAktivnost), DodatnaTjelesnaAktivnost.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            'Namirnice italic
            '  Dim txtSelection As RichTextBox = Me.RichTextBox1
            Dim txtSearch As String = "Namirnice:"
            Dim textEnd As Integer = Rtb.TextLength
            Dim index As Integer = 0
            Dim lastIndex As Integer = Rtb.Text.LastIndexOf(txtSearch)
            Dim myStyle As FontStyle
            myStyle = FontStyle.Bold + FontStyle.Italic
            '  Me.RichTextBox1.SelectionFont = New Font("Arial", 8, myStyle)
            While index < lastIndex
                Rtb.Find(txtSearch, index, textEnd, RichTextBoxFinds.None)
                'txtSelection.SelectionBackColor = Color.Yellow
                Rtb.SelectionFont = New Font("Arial", 8, myStyle)
                index = Rtb.Text.IndexOf(txtSearch, index) + 1
            End While

            'Prazno polje
            Dim txtSearch1 As String = "~"
            Dim textEnd1 As Integer = Rtb.TextLength
            Dim index1 As Integer = 0
            Dim lastIndex1 As Integer = Rtb.Text.LastIndexOf(txtSearch1)
            While index1 < lastIndex1
                Rtb.Find(txtSearch1, index1, textEnd1, RichTextBoxFinds.None)
                Rtb.SelectionColor = Color.White
                index1 = Rtb.Text.IndexOf(txtSearch1, index1) + 1
            End While

        End With
    End Sub
End Module
