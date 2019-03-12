Module EnergetskaPotrosnjaDjecaModul
    Sub EnergetskaPotrosnjaDjeca()
        With Form1
            Dim Dob As Integer = .ComboBox1.Text
            'Dim Spol As String
            Dim Visina As Double = .ComboBox2.Text / 100
            Dim Masa As Double = .ComboBox3.Text
            Dim PA As Double
            Dim EER As Integer   'ukupna energetska potrošnja

            '9-18
            If Dob >= 9 And Dob < 18 Then
                .RadioButton16.Enabled = False  'povecanje tjelesne mase
                .RadioButton18.Enabled = False   'povecanje misicne mase
            Else
                'Cilj
                .RadioButton15.Enabled = True   'redukcija tjelesne mase
                .RadioButton16.Enabled = True    'povecanje tjelesne mase
                .RadioButton17.Enabled = True    'zadrzavanje postojece tjelesne mase
                .RadioButton18.Enabled = True   'povecanje misicne mase
            End If

            'sedentary
            If .RadioButton3.Checked = True Then
                PA = 1
            End If
            'low active
            If .RadioButton4.Checked = True Then
                PA = 1.13
            End If
            'active
            If .RadioButton5.Checked = True Then
                PA = 1.26
            End If
            'very(active)
            If .RadioButton6.Checked = True Then
                PA = 1.42
            End If


            'Muški
            If .RadioButton1.Checked = True Then
                EER = Val(88.5 - (61.9 * Dob) + PA * (26.7 * Masa + 903 * Visina) + 25)
                .Label13.Text = EER
            End If

            'Ženski
            If .RadioButton2.Checked = True Then
                EER = Val(135.3 - (30.8 * Dob) + PA * (10.0 * Masa + 934 * Visina) + 25)
                .Label13.Text = EER
            End If

            'Detaljan izračun energetske potrošnje
            If .Label315.Text > 0 Then
                .Label13.Text = .Label315.Text
                EER = .Label315.Text
            End If

            .TextBox3.Text = EER    'Preporučena energetska vrijednost jelovnika

            If .Label11.Text > 25 Then
                .TextBox4.Text = 200 'Dodatna energetska potrosnja
                ' Form2.Label5.Text = "Za osobe mlađe od 18 godina, redukcijska dijeta smije se provoditi isključivo pod nadzorom stručne osobe uz liječničko dopuštenje i pračenje!"
            End If

            If .Label11.Text < 18.5 Then
                .TextBox3.Text = EER
            End If


            'BMI
            Dim BMI As Double = .Label11.Text

            If BMI = 0 Then
                .Label25.Text = ""
            End If

            If BMI < 18.5 And BMI > 0 Then
                .Label25.Text = "SNIŽENA TJELESNA MASA!"
                .Label25.BackColor = Color.CornflowerBlue
                ' .TextBox4.Text = DodatnaEnergetskaPotrosnja
                '.TextBox3.Text = TEE + DodatniEnergetskiUnos + DodatnaEnergetskaPotrosnja
                .RadioButton17.Checked = True   'zadrzavanje postojece tjelesne mase
                .RadioButton15.Enabled = False   'Enable - redukcija tjelesne mase
            End If

            If BMI >= 18.5 And BMI < 25 Then
                .Label25.Text = "NORMALNA TJELESNA MASA."
                .Label25.BackColor = Color.GreenYellow
                '   .TextBox4.Text = DodatnaEnergetskaPotrosnja
                '  .TextBox3.Text = TEE + DodatnaEnergetskaPotrosnja
                .RadioButton17.Checked = True    'zadrzavanj postojece tjelesne mase
                .RadioButton15.Enabled = False   'Enable - redukcija tjelesne mase
            End If

            If BMI >= 25 And BMI < 30 Then
                .Label25.Text = "POVIŠENA TJELESNA MASA!"
                .Label25.BackColor = Color.Yellow
                '    .TextBox4.Text = DodatnaEnergetskaPotrosnja
                '   .TextBox3.Text = TEE - DodatniEnergetskiUnos
                .RadioButton15.Checked = True    'smanjenje tjelesne mase
                .RadioButton15.Enabled = True  'Enable - redukcija tjelesne mase
            End If

            If BMI >= 30 Then
                .Label25.Text = "GOJAZNOST!"
                .Label25.BackColor = Color.Red
                '     .TextBox4.Text = DodatnaEnergetskaPotrosnja
                '    .TextBox3.Text = TEE - DodatniEnergetskiUnos
                .RadioButton15.Checked = True    'smanjenje tjelesne mase
                .RadioButton15.Enabled = True  'Enable - redukcija tjelesne mase
            End If

            If BMI = 0 Then
                .TextBox3.Text = ""
            End If



        End With
    End Sub
End Module
