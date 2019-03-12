Module ParametriProgressBarModul
    Sub ParametriProgressBar()
        On Error Resume Next
        With Form1
            Dim Dob As Double = Val(.ComboBox1.Text)    'Dob
            Dim Postotak As Integer

            .ProgressBar8.Value = .ProgressBar7.Value    'Energija

            Postotak = (.ListBox22.Items(0) / .ListBox26.Items(0)) * 100   'Zasicene masti
            .Label201.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label201.Size = New Size(100, 8)
            If Postotak > 100 Then
                .Label201.BackColor = Color.Red
            Else
                .Label201.BackColor = Color.LightCoral
            End If

            If Postotak > 200 Then .Label201.Size = New Size(100, 8)

            Postotak = (.ListBox22.Items(1) / .ListBox26.Items(1)) * 100   'Trans masne kiseline
            .Label243.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label243.Size = New Size(100, 8)
            If Postotak > 100 Then
                .Label243.BackColor = Color.Red
            Else
                .Label243.BackColor = Color.LightCoral
            End If

            Postotak = (.ListBox22.Items(2) / .ListBox26.Items(2)) * 100   'Kolesterol
            .Label244.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label244.Size = New Size(100, 8)
            If Postotak > 100 Then
                .Label244.BackColor = Color.Red
            Else
                .Label244.BackColor = Color.LightCoral
            End If

            Postotak = (.ListBox23.Items(0) / .ListBox28.Items(0)) * 100   'Natrij
            .Label246.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label246.Size = New Size(100, 8)
            If Postotak > 100 Then
                .Label246.BackColor = Color.Red
            Else
                .Label246.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox35.Items(0) / .ListBox36.Items(0)) * 100   'Ostale namirnice
            .Label249.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label249.Size = New Size(100, 8)
            If Postotak > 100 Then
                .Label249.BackColor = Color.Red
            Else
                .Label249.BackColor = Color.LightCoral
            End If


            Postotak = (.ListBox24.Items(0) / .ListBox29.Items(0)) * 100   'Vlakna
            .Label255.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label255.Size = New Size(100, 8)

            Postotak = (.ListBox24.Items(1) / .ListBox29.Items(1)) * 100   'Jednostruko nezasicene masti
            .Label256.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label256.Size = New Size(100, 8)
            If .ListBox24.Items(1) > .ListBox30.Items(1) Then
                .Label256.BackColor = Color.Red
            Else
                .Label256.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(2) / .ListBox29.Items(2)) * 100   'Visetruko nezasicene masti
            .Label257.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label257.Size = New Size(100, 8)
            If .ListBox24.Items(2) > .ListBox30.Items(2) Then
                .Label257.BackColor = Color.Red
            Else
                .Label257.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(3) / .ListBox29.Items(3)) * 100   'Kalcij
            .Label258.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label258.Size = New Size(100, 8)
            If .ListBox24.Items(3) > .ListBox30.Items(3) Then
                .Label258.BackColor = Color.Red
            Else
                .Label258.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(4) / .ListBox29.Items(4)) * 100   'Magnezij
            .Label259.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label259.Size = New Size(100, 8)
            If .ListBox24.Items(4) > .ListBox30.Items(4) Then
                .Label259.BackColor = Color.Red
            Else
                .Label259.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(5) / .ListBox29.Items(5)) * 100   'Fosfor
            .Label260.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label260.Size = New Size(100, 8)
            If .ListBox24.Items(5) > .ListBox30.Items(5) Then
                .Label260.BackColor = Color.Red
            Else
                .Label260.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(6) / .ListBox29.Items(6)) * 100   'Zeljezo
            .Label261.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label261.Size = New Size(100, 8)
            If .ListBox24.Items(6) > .ListBox30.Items(6) Then
                .Label261.BackColor = Color.Red
            Else
                .Label261.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(7) / .ListBox29.Items(7)) * 100   'Bakar
            .Label262.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label262.Size = New Size(100, 8)
            If .ListBox24.Items(7) > .ListBox30.Items(7) Then
                .Label262.BackColor = Color.Red
            Else
                .Label262.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(8) / .ListBox29.Items(8)) * 100   'Cink
            .Label263.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label263.Size = New Size(100, 8)
            If .ListBox24.Items(8) > .ListBox30.Items(8) Then
                .Label263.BackColor = Color.Red
            Else
                .Label263.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(9) / .ListBox29.Items(9)) * 100   'Mangan
            .Label264.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label264.Size = New Size(100, 8)
            If .ListBox24.Items(9) > .ListBox30.Items(9) Then
                .Label264.BackColor = Color.Red
            Else
                .Label264.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(10) / .ListBox29.Items(10)) * 100   'Selen
            .Label265.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label265.Size = New Size(100, 8)
            If .ListBox24.Items(10) > .ListBox30.Items(10) Then
                .Label265.BackColor = Color.Red
            Else
                .Label265.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(11) / .ListBox29.Items(11)) * 100   'Jod
            .Label266.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label266.Size = New Size(100, 8)
            If .ListBox24.Items(11) > .ListBox30.Items(11) Then
                .Label266.BackColor = Color.Red
            Else
                .Label266.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(12) / .ListBox29.Items(12)) * 100   'Retinol
            .Label267.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label267.Size = New Size(100, 8)
            If .ListBox24.Items(12) > .ListBox30.Items(12) Then
                .Label267.BackColor = Color.Red
            Else
                .Label267.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(13) / .ListBox29.Items(13)) * 100   'VitaminD
            .Label268.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label268.Size = New Size(100, 8)
            If .ListBox24.Items(13) > .ListBox30.Items(13) Then
                .Label268.BackColor = Color.Red
            Else
                .Label268.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(14) / .ListBox29.Items(14)) * 100   'VitaminE
            .Label269.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label269.Size = New Size(100, 8)
            If .ListBox24.Items(14) > .ListBox30.Items(14) Then
                .Label269.BackColor = Color.Red
            Else
                .Label269.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(15) / .ListBox29.Items(15)) * 100   'VitaminB1
            .Label270.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label270.Size = New Size(100, 8)
            If .ListBox24.Items(15) > .ListBox30.Items(15) Then
                .Label270.BackColor = Color.Red
            Else
                .Label270.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(16) / .ListBox29.Items(16)) * 100   'VitmainB2
            .Label271.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label271.Size = New Size(100, 8)
            If .ListBox24.Items(16) > .ListBox30.Items(16) Then
                .Label271.BackColor = Color.Red
            Else
                .Label271.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(17) / .ListBox29.Items(17)) * 100   'VitaminB3
            .Label272.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label272.Size = New Size(100, 8)
            If .ListBox24.Items(17) > .ListBox30.Items(17) Then
                .Label272.BackColor = Color.Red
            Else
                .Label272.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(18) / .ListBox29.Items(18)) * 100   'VitaminB6
            .Label273.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label273.Size = New Size(100, 8)
            If .ListBox24.Items(18) > .ListBox30.Items(18) Then
                .Label273.BackColor = Color.Red
            Else
                .Label273.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(19) / .ListBox29.Items(19)) * 100   'VitaminB12
            .Label274.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label274.Size = New Size(100, 8)
            If .ListBox24.Items(19) > .ListBox30.Items(19) Then
                .Label274.BackColor = Color.Red
            Else
                .Label274.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(20) / .ListBox29.Items(20)) * 100   'Folat
            .Label275.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label275.Size = New Size(100, 8)
            If .ListBox24.Items(20) > .ListBox30.Items(20) Then
                .Label275.BackColor = Color.Red
            Else
                .Label275.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(21) / .ListBox29.Items(21)) * 100   'Pontotenska kiselina
            .Label276.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label276.Size = New Size(100, 8)
            If .ListBox24.Items(21) > .ListBox30.Items(21) Then
                .Label276.BackColor = Color.Red
            Else
                .Label276.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(22) / .ListBox29.Items(22)) * 100   'Biotin
            .Label277.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label277.Size = New Size(100, 8)
            If .ListBox24.Items(22) > .ListBox30.Items(22) Then
                .Label277.BackColor = Color.Red
            Else
                .Label277.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(23) / .ListBox29.Items(23)) * 100   'VitaminC
            .Label278.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label278.Size = New Size(100, 8)
            If .ListBox24.Items(23) > .ListBox30.Items(23) Then
                .Label278.BackColor = Color.Red
            Else
                .Label278.BackColor = Color.SkyBlue
            End If

            Postotak = (.ListBox24.Items(24) / .ListBox29.Items(24)) * 100   'VitaminK
            .Label279.Size = New Size(Postotak / 2, 8)
            If Postotak > 200 Then .Label279.Size = New Size(100, 8)
            If .ListBox24.Items(24) > .ListBox30.Items(24) Then
                .Label279.BackColor = Color.Red
            Else
                .Label279.BackColor = Color.SkyBlue
            End If

            'djeca
            If Dob >= 9 And Dob < 18 Then
                .Label258.BackColor = Color.SkyBlue
                .Label259.BackColor = Color.SkyBlue
                .Label260.BackColor = Color.SkyBlue
                .Label261.BackColor = Color.SkyBlue
                .Label262.BackColor = Color.SkyBlue
                .Label263.BackColor = Color.SkyBlue
                .Label264.BackColor = Color.SkyBlue
                .Label265.BackColor = Color.SkyBlue
                .Label266.BackColor = Color.SkyBlue
                .Label267.BackColor = Color.SkyBlue
                .Label268.BackColor = Color.SkyBlue
                .Label269.BackColor = Color.SkyBlue
                .Label270.BackColor = Color.SkyBlue
                .Label271.BackColor = Color.SkyBlue
                .Label272.BackColor = Color.SkyBlue
                .Label273.BackColor = Color.SkyBlue
                .Label274.BackColor = Color.SkyBlue
                .Label275.BackColor = Color.SkyBlue
                .Label276.BackColor = Color.SkyBlue
                .Label277.BackColor = Color.SkyBlue
                .Label278.BackColor = Color.SkyBlue
                .Label279.BackColor = Color.SkyBlue
            End If
        End With
    End Sub
End Module
