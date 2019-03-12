Module BmiModul
    Sub BMI()
        On Error Resume Next
        With Form1
            Dim Visina As Integer = .ComboBox2.Text
            Dim Masa As Integer = .ComboBox3.Text
            Dim Bmi As Double

              Bmi = Format(Masa / ((Visina / 100) * (Visina / 100)), "0.0")
            .Label11.Text = Bmi

            'Cilj
            .RadioButton15.Enabled = True   'redukcija tjelesne mase
            .RadioButton16.Enabled = True    'povecanje tjelesne mase
            .RadioButton17.Enabled = True    'zadrzavanje postojece tjelesne mase
            .RadioButton18.Enabled = True   'povecanje misicne mase

            'Progress Bar
            .PictureBox5.Location = New Point(6 + (Bmi - 12) * 10, 26)
            If Bmi > 41 Then .PictureBox5.Location = New Point(300, 26)
            If Bmi < 12 Then .PictureBox5.Location = New Point(6, 26)

            'Primjerena masa
            .Label12.Text = Format((18.5 * (Visina / 100) * (Visina / 100)), "0") & " - " & Format((25 * (Visina / 100) * (Visina / 100)), "0")

            'Djeca 9-18 god
            '      If Val(.ComboBox1.Text) >= 9 And Val(.ComboBox1.Text) < 18 Then
            'EnergetskaPotrosnjaDjeca()
            '   Exit Sub
            '  End If

            'BMR
            Dim BMR As Double
            Dim Dob As Integer = Val(.ComboBox1.Text)
            '   Dob = .ComboBox1.Text


            'BMR - NOVO - Harris-Benedict 1918-1919
            'muškarci
            '       If .RadioButton1.Checked = True Then
            'BMR = 66 + (13.7 * Masa) + (5 * (Visina / 100)) - (6.8 * Dob)
            '    End If
            'žene
            '   If .RadioButton2.Checked = True Then
            'BMR = 655 + (9.6 * Masa) + (1.8 * (Visina / 100)) - (4.7 * Dob)
            '   End If

            'BMR - NOVO - Roza & Shizgal revizija iz 1984 g
            'muškarci
            '      If .RadioButton1.Checked = True Then
            'BMR = 88.362 + (13.397 * Masa) + (4.799 * Visina) - (5.677 * Dob)
            '    End If
            'žene
            '   If .RadioButton2.Checked = True Then
            'BMR = 447.593 + (9.347 * Masa) + (3.098 * Visina) - (4.33 * Dob)
            'End If


            'BMR - St Jeor equation
            'muškarci
            If .RadioButton1.Checked = True Then
                BMR = 10 * Masa + 6.25 * Visina - 5 * Dob + 5
            End If
            'žene
            If .RadioButton2.Checked = True Then
                BMR = 10 * Masa + 6.25 * Visina - 5 * Dob - 161
            End If

            'BMR - STARO
            '18-30, muškarci
            '     If Dob >= 18 And Dob < 30 And .RadioButton1.Checked = True Then
            'BMR = 15.06 * Masa + 10.04 * (Visina / 100) + 705
            '     End If
            '18-30, žene
            '    If Dob >= 18 And Dob < 30 And .RadioButton2.Checked = True Then
            'BMR = 13.62 * Masa + 283 * (Visina / 100) + 98
            ' End If
            '30-60, muškarci
            'If Dob >= 30 And Dob < 60 And .RadioButton1.Checked = True Then
            'BMR = 11.47 * Masa + 2.629 * (Visina / 100) + 877
            '  End If
            '30-60, žene
            ' If Dob >= 30 And Dob < 60 And .RadioButton2.Checked = True Then
            'BMR = 8.126 * Masa + 4.434 * (Visina / 100) + 843
            '     End If
            '>60, muškarci
            ' If Dob >= 60 And .RadioButton1.Checked = True Then
            'BMR = 13.5 * Masa + 487
            '  End If
            '>60, žene
            ' If Dob >= 60 And .RadioButton2.Checked = True Then
            'BMR = 10.5 * Masa + 596
            '  End If

            'PAL
            Dim PAL As Double

            'muškarci - na poslu-izrazito slab, izvan posla-izrazito slab
            If .RadioButton1.Checked = True And .RadioButton3.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.2
            End If

            'muškarci - na poslu-izrazito slab, izvan posla-slab
            If .RadioButton1.Checked = True And .RadioButton3.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.3
            End If

            'muškarci - na poslu-izrazito slab, izvan umjeren
            If .RadioButton1.Checked = True And .RadioButton3.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.4
            End If

            'muškarci - na poslu-izrazito slab, izvan izražen
            If .RadioButton1.Checked = True And .RadioButton3.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.5
            End If

            'muškarci - na poslu-slab, izvan posla-izrazito slab
            If .RadioButton1.Checked = True And .RadioButton4.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.3
            End If

            'muškarci - na poslu-slab, izvan posla-slab
            If .RadioButton1.Checked = True And .RadioButton4.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.375
            End If

            'muškarci - na poslu-slab, izvan posla-umjeren
            If .RadioButton1.Checked = True And .RadioButton4.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.475
            End If

            'muškarci - na poslu-slab, izvan posla-izražen
            If .RadioButton1.Checked = True And .RadioButton4.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.6
            End If

            'muškarci - na poslu-umjeren, izvan posla-izrazito slab
            If .RadioButton1.Checked = True And .RadioButton5.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.5
            End If

            'muškarci - na poslu-umjeren, izvan posla-slab
            If .RadioButton1.Checked = True And .RadioButton5.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.6
            End If

            'muškarci - na poslu-umjeren, izvan posla-umjeren
            If .RadioButton1.Checked = True And .RadioButton5.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.7
            End If

            'muškarci - na poslu-umjeren, izvan posla-izražen
            If .RadioButton1.Checked = True And .RadioButton5.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.8
            End If

            'muškarci - na poslu-izražen, izvan posla-izrazito slab
            If .RadioButton1.Checked = True And .RadioButton6.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.6
            End If

            'muškarci - na poslu-izražen, izvan posla-slab
            If .RadioButton1.Checked = True And .RadioButton6.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.7
            End If

            'muškarci - na poslu-izražen, izvan posla-umjeren
            If .RadioButton1.Checked = True And .RadioButton6.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.8
            End If

            'muškarci - na poslu-izražen, izvan posla-izražen
            If .RadioButton1.Checked = True And .RadioButton6.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.9
            End If

            'žene - na poslu-izrazito slab, izvan posla-izrazito slab
            If .RadioButton2.Checked = True And .RadioButton3.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.2
            End If

            'žene - na poslu-izrazito slab, izvan posla slab
            If .RadioButton2.Checked = True And .RadioButton3.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.3
            End If

            'žene - na poslu-izrazito slab, izvan posla-umjeren
            If .RadioButton2.Checked = True And .RadioButton3.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.4
            End If

            'žene - na poslu-izrazito slab, izvan posla-izražen
            If .RadioButton2.Checked = True And .RadioButton3.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.5
            End If


            'žene - na poslu-slab, izvan posla-izrazito slab
            If .RadioButton2.Checked = True And .RadioButton4.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.3
            End If

            'žene - na poslu-slab, izvan posla-slab
            If .RadioButton2.Checked = True And .RadioButton4.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.375
            End If

            'žene - na poslu-slab, izvan posla-umjeren
            If .RadioButton2.Checked = True And .RadioButton4.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.475
            End If

            'žene - na poslu-slab, izvan posla-izražen
            If .RadioButton2.Checked = True And .RadioButton4.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.6
            End If

            'žene - na poslu-umjeren, izvan posla-izrazito slab
            If .RadioButton2.Checked = True And .RadioButton5.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.4
            End If

            'žene - na poslu-umjeren, izvan posla-slab
            If .RadioButton2.Checked = True And .RadioButton5.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.5
            End If

            'žene - na poslu-umjeren, izvan posla-umjeren
            If .RadioButton2.Checked = True And .RadioButton5.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.6
            End If

            'žene - na poslu-umjeren, izvan posla-izražen
            If .RadioButton2.Checked = True And .RadioButton5.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.7
            End If

            'žene - na poslu-izražen, izvan posla-izrazito slab
            If .RadioButton2.Checked = True And .RadioButton6.Checked = True And .RadioButton7.Checked = True Then
                PAL = 1.5
            End If

            'žene - na poslu-izražen, izvan posla-slab
            If .RadioButton2.Checked = True And .RadioButton6.Checked = True And .RadioButton8.Checked = True Then
                PAL = 1.6
            End If

            'žene - na poslu-izražen, izvan posla-umjeren
            If .RadioButton2.Checked = True And .RadioButton6.Checked = True And .RadioButton9.Checked = True Then
                PAL = 1.7
            End If

            'žene - na poslu-izražen, izvan posla-izražen
            If .RadioButton2.Checked = True And .RadioButton6.Checked = True And .RadioButton10.Checked = True Then
                PAL = 1.8
            End If

            Dim DIT As Double
            DIT = 0.1 * (PAL * BMR)

            Dim TEE As Double
            TEE = Format((PAL * BMR + DIT), "0")

            .Label13.Text = TEE

            'Detalja energetska potrosnja
            If .Label315.Text > 0 Then
                .Label13.Text = .Label315.Text
                TEE = .Label315.Text
            End If

            '   .TextBox3.Text = TEE - 300

            ' DodatnaEnergetskaPotrosnjaPreporuka()    'Preporucena dodatna energetska potrosnja
            Dim DodatniEnergetskiUnos As Integer
            Dim DodatnaEnergetskaPotrosnja As Integer

            If .RadioButton3.Checked = True And .RadioButton7.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton4.Checked = True And .RadioButton7.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton5.Checked = True And .RadioButton7.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton6.Checked = True And .RadioButton7.Checked = True Then
                DodatnaEnergetskaPotrosnja = 50
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton3.Checked = True And .RadioButton8.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton4.Checked = True And .RadioButton8.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton5.Checked = True And .RadioButton8.Checked = True Then
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton6.Checked = True And .RadioButton8.Checked = True Then
                DodatnaEnergetskaPotrosnja = 100
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton3.Checked = True And .RadioButton9.Checked = True Then
                DodatnaEnergetskaPotrosnja = 100
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton4.Checked = True And .RadioButton9.Checked = True Then
                DodatnaEnergetskaPotrosnja = 100
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton5.Checked = True And .RadioButton9.Checked = True Then
                DodatnaEnergetskaPotrosnja = 100
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton6.Checked = True And .RadioButton9.Checked = True Then
                DodatnaEnergetskaPotrosnja = 0
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton3.Checked = True And .RadioButton10.Checked = True Then
                DodatnaEnergetskaPotrosnja = 0
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton4.Checked = True And .RadioButton10.Checked = True Then
                DodatnaEnergetskaPotrosnja = 0
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton5.Checked = True And .RadioButton10.Checked = True Then
                DodatnaEnergetskaPotrosnja = 0
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If
            If .RadioButton6.Checked = True And .RadioButton10.Checked = True Then
                DodatnaEnergetskaPotrosnja = 0
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If

            'Detalja energetska potrosnja (izracun dodatne energetske potrosnje)
            If .Label315.Text > 0 Then
                TEE = .Label315.Text
                DodatnaEnergetskaPotrosnja = 200
                DodatniEnergetskiUnos = 500 - DodatnaEnergetskaPotrosnja
            End If

            'CILJ
            .Label196.Text = DodatnaEnergetskaPotrosnja
            .Label197.Text = DodatniEnergetskiUnos

            'djeca
            ' If Dob >= 9 And Dob < 18 Then

            '  End If
            '          Dim DodatniEnergetskiUnos As Integer
            '         Dim DodatnaEnergetskaPotrosnja As Integer
            '        If PAL >= 1.4 And PAL < 1.5 Then
            '          'DodatnaEnergetskaPotrosnja = 200
            '         DodatniEnergetskiUnos = 300
            '        End If
            '       If PAL >= 1.5 And PAL < 1.6 Then
            '           'DodatnaEnergetskaPotrosnja = 150
            '          DodatniEnergetskiUnos = 350
            '         End If
            '        If PAL >= 1.6 And PAL < 1.7 Then
            'DodatnaEnergetskaPotrosnja = 100
            '           DodatniEnergetskiUnos = 400
            '          End If
            '         If PAL >= 1.7 And PAL < 1.8 Then
            'DodatnaEnergetskaPotrosnja = 50
            '        DodatniEnergetskiUnos = 450
            '       End If
            '      If PAL >= 1.8 And PAL <= 1.9 Then
            'DodatnaEnergetskaPotrosnja = 0
            '         DodatniEnergetskiUnos = 500
            '        End If

            'CILJ
            '          .Label196.Text = DodatnaEnergetskaPotrosnja
            '         .Label197.Text = DodatniEnergetskiUnos

            'BMI
            If Bmi = 0 Then
                .Label25.Text = ""
            End If

            If Bmi < 18.5 And Bmi > 0 Then
                .Label25.Text = "SNIŽENA TJELESNA MASA"
                .Label25.BackColor = Color.CornflowerBlue
                .TextBox4.Text = DodatnaEnergetskaPotrosnja
                .TextBox3.Text = TEE + DodatniEnergetskiUnos + DodatnaEnergetskaPotrosnja
                .RadioButton16.Checked = True   'smanjenje tjelesne mase
            End If

            If Bmi >= 18.5 And Bmi < 25 Then
                .Label25.Text = "NORMALNA TJELESNA MASA"
                .Label25.BackColor = Color.GreenYellow
                .TextBox4.Text = DodatnaEnergetskaPotrosnja
                .TextBox3.Text = TEE + DodatnaEnergetskaPotrosnja
                .RadioButton17.Checked = True    'zadrzavanj postojece tjelesne mase
            End If

            If Bmi >= 25 And Bmi < 30 Then
                .Label25.Text = "POVIŠENA TJELESNA MASA"
                .Label25.BackColor = Color.Yellow
                .TextBox4.Text = DodatnaEnergetskaPotrosnja
                .TextBox3.Text = TEE - DodatniEnergetskiUnos
                .RadioButton15.Checked = True    'smanjenje tjelesne mase
            End If

            If Bmi >= 30 Then
                .Label25.Text = "GOJAZNOST"
                .Label25.BackColor = Color.Red
                .TextBox4.Text = DodatnaEnergetskaPotrosnja
                .TextBox3.Text = TEE - DodatniEnergetskiUnos
                .RadioButton15.Checked = True    'smanjenje tjelesne mase
            End If

            If Bmi = 0 Then
                .TextBox3.Text = ""
            End If

            '           If Bmi < 18.5 And Bmi > 0 Then
            '.TextBox3.Text = TEE + 300
            '           End If

            '            If Bmi >= 18.5 And Bmi < 25 Then
            '.TextBox3.Text = TEE
            '          End If

            '           If Bmi >= 25 And Bmi < 30 Then
            '.TextBox3.Text = TEE - 300
            '           End If

            '            If Bmi >= 30 Then
            '.TextBox3.Text = TEE - 300
            '           End If



            ' OMJER OPSEGA STRUKA I BOKOVA
            Dim OpsegStruka As Double = Val(.ComboBox6.Text)
            Dim OpsegBokova As Double = Val(.ComboBox7.Text)
            Dim WHR As Double = Format(OpsegStruka / OpsegBokova, "0.00")
            .Label195.Visible = False
            .Label193.Text = WHR
            .Label208.Text = .ComboBox6.Text   'Omjer struka
            'muskarci
            If .RadioButton1.Checked = True And WHR < 1 Then
                .Label194.Text = "GINOIDNI TIP (u slučaju nakupljanja masnog tkiva, ono se nakuplja u području bokova)"
            End If
            If .RadioButton1.Checked = True And WHR >= 1 Then
                .Label194.Text = "ANDROIDNI TIP (u slučaju nakupljanja masnog tkiva, ono se nakuplja u području struka)"
            End If
            'opseg struka - muskarci
            If .RadioButton1.Checked = True And OpsegStruka < 94 Then
                .Label195.Text = ""
            End If
            If .RadioButton1.Checked = True And OpsegStruka >= 94 And OpsegStruka < 102 Then
                .Label195.Visible = True
                .Label195.Text = "OPSEG STRUKA IZMEĐU 94 I 102 CM PREDSTAVLJA POVEĆAN RIZIK OD POJAVE RAZLIČITIH BOLESTI (npr. šećerne bolesti i bolesti srca)!"
            End If
            If .RadioButton1.Checked = True And OpsegStruka >= 102 Then
                .Label195.Visible = True
                .Label195.Text = "OPSEG STRUKA IZNAD 102 CM PREDSTAVLJA VRLO VISOK RIZIK OD POJAVE RAZLIČITIH BOLESTI (npr. šećerne bolesti i bolesti srca)!"
            End If
            'zene
            If .RadioButton2.Checked = True And WHR < 0.8 Then
                .Label194.Text = "GINOIDNI TIP (u slučaju nakupljanja masnog tkiva, ono se nakuplja u području bokova)"
            End If
            If .RadioButton2.Checked = True And WHR >= 0.8 Then
                .Label194.Text = "ANDROIDNI TIP (u slučaju nakupljanja masnog tkiva, ono se nakuplja u području struka)"
            End If
            'opseg struka - zene
            If .RadioButton2.Checked = True And OpsegStruka < 80 Then
                .Label195.Text = ""
            End If
            If .RadioButton2.Checked = True And OpsegStruka >= 80 And OpsegStruka < 88 Then
                .Label195.Visible = True
                .Label195.Text = "OPSEG STRUKA IZMEĐU 80 I 88 CM PREDSTAVLJA POVEĆAN RIZIK OD POJAVE RAZLIČITIH BOLESTI (npr. šećerne bolesti i bolesti srca)!"
            End If
            If .RadioButton2.Checked = True And OpsegStruka >= 88 Then
                .Label195.Visible = True
                .Label195.Text = "OPSEG STRUKA IZNAD 88 CM PREDSTAVLJA VRLO VISOK RIZIK OD POJAVE RAZLIČITIH BOLESTI (npr. šećerne bolesti i bolesti srca)!"
            End If






            '       If .TextBox3.Text = -300 Or .TextBox3.Text = 0 Then .TextBox3.Text = ""
            '       If .Label18.Text = 0 Then .Label18.Text = ""
            '     If .Label17.Text = 0 Then .Label17.Text = ""



            'Progress Bar

            'Me.thePictureBox.Location = new Point(x, y)
            '          .PictureBox5.Location = New Point(8 + (Bmi - 12) * 10, 30)
            '        If Bmi > 41 Then .PictureBox5.Location = New Point(300, 30)
            '     If Bmi < 12 Then .PictureBox5.Location = New Point(8, 30)

            'korisnik - slika - muškarac
            '      If Bmi > 5 And Bmi < 60 And .RadioButton1.Checked = True Then
            '.PictureBox6.Visible = True
            '           Else
            '          .PictureBox6.Visible = False
            '         End If
            '        .PictureBox6.Size = New Point((Bmi * 4.8), 100)
            '       .PictureBox6.Location = New Point(300 - (Bmi * 2.4), 81)


            'korisnik - slika - žena
            '           If Bmi > 5 And Bmi < 60 And .RadioButton2.Checked = True Then
            '.PictureBox7.Visible = True
            '            Else
            '           .PictureBox7.Visible = False
            '          End If
            '         .PictureBox7.Size = New Point((Bmi * 3.6), 85)
            '        .PictureBox7.Location = New Point(305 - (Bmi * 1.8), 90)



            '          If .Label18.Text = "" Then .TextBox3.Text = ""

        End With
    End Sub
End Module
