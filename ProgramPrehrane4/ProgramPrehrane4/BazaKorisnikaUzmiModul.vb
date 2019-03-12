Module BazaKorisnikaUzmiModul
    Sub BazaKorisnikaUzmi()
        On Error Resume Next
        With Form1

            .TextBox1.Text = .Label211.Text   'Ime
            .TextBox2.Text = .Label212.Text    'Prezime
            .ComboBox1.Text = .Label213.Text     'Dob
            If .Label178.Text = "Muškarac" Then
                .RadioButton1.Checked = True 'Spol
            End If
            If .Label178.Text = "Žena" Then
                .RadioButton2.Checked = True 'Spol
            End If
            .ComboBox2.Text = .Label250.Text      'Visina
            .ComboBox3.Text = .Label251.Text     'Masa
            .ComboBox6.Text = .Label252.Text     'Opseg struka
            .ComboBox7.Text = .Label253.Text      'Opseg bokova
            .Label193.Text = .Label289.Text       'WHR

            '        If .ComboBox1.Text < 18 And .ComboBox1.Text >= 9 Then   'djeca
            'If .Label292.Text = "Izrazito slab" Then
            '.RadioButton3.Checked = True 'Intenzitet aktivnosti (djeca)
            'End If
            'If .Label292.Text = "Slab" Then
            '.RadioButton4.Checked = True 'Intenzitet aktivnosti (djeca)
            'End If
            'If .Label292.Text = "Umjeren" Then
            '.RadioButton5.Checked = True 'Intenzitet aktivnosti (djeca)
            'End If
            'If .Label292.Text = "Izražen" Then
            '.RadioButton6.Checked = True 'Intenzitet aktivnosti (djeca)
            'End If
            'Else    'odrasli
            If .Label290.Text = "Izrazito slab" Then
                .RadioButton3.Checked = True 'Intenzitet aktivnosti na poslu
            End If
            If .Label290.Text = "Slab" Then
                .RadioButton4.Checked = True 'Intenzitet aktivnosti na poslu
            End If
            If .Label290.Text = "Umjeren" Then
                .RadioButton5.Checked = True 'Intenzitet aktivnosti na poslu
            End If
            If .Label290.Text = "Izražen" Then
                .RadioButton6.Checked = True 'Intenzitet aktivnosti na poslu
            End If

            If .Label291.Text = "Izrazito slab" Then
                .RadioButton7.Checked = True   'Intenzitet aktivnosti izvan posla
            End If
            If .Label291.Text = "Slab" Then
                .RadioButton8.Checked = True  'Intenzitet aktivnosti izvan posla
            End If
            If .Label291.Text = "Umjeren" Then
                .RadioButton9.Checked = True 'Intenzitet aktivnosti izvan posla
            End If
            If .Label291.Text = "Izražen" Then
                .RadioButton10.Checked = True 'Intenzitet aktivnosti izvan posla
            End If

            'End If

            .TextBox75.Text = .Label296.Text     'Napomena
            .Label315.Text = .Label294.Text   'TEE


            BMI()    'Izračun BMI, TEE, WHR

            .TabControl3.SelectedIndex = 0

        End With
    End Sub
End Module
