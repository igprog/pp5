Module OstaleNamirnicePreporukaModul
    Sub OstaleNamirnicePreporuka()
        On Error Resume Next
        With Form1
            Dim K As Integer = 1
            Dim EnergVrJelovnika As Integer = .TextBox3.Text
            Dim Preporuka As Integer = 0
            If EnergVrJelovnika < 1900 And EnergVrJelovnika > 500 Then Preporuka = 195
            If EnergVrJelovnika >= 1900 And EnergVrJelovnika < 2100 Then Preporuka = 267
            If EnergVrJelovnika >= 2100 And EnergVrJelovnika < 2300 Then Preporuka = 290
            If EnergVrJelovnika >= 2300 And EnergVrJelovnika < 2500 Then Preporuka = 362
            If EnergVrJelovnika >= 2500 And EnergVrJelovnika < 2700 Then Preporuka = 410
            If EnergVrJelovnika >= 2700 And EnergVrJelovnika < 2900 Then Preporuka = 426
            If EnergVrJelovnika >= 2900 And EnergVrJelovnika < 3100 Then Preporuka = 512
            If EnergVrJelovnika >= 3100 And EnergVrJelovnika < 3300 Then Preporuka = 648
            If EnergVrJelovnika >= 3300 And EnergVrJelovnika < 3500 Then Preporuka = 700
            If EnergVrJelovnika >= 3500 And EnergVrJelovnika < 3700 Then Preporuka = 750
            If EnergVrJelovnika >= 3700 And EnergVrJelovnika < 3900 Then Preporuka = 780
            If EnergVrJelovnika >= 3900 Then Preporuka = 830
            .Label280.Text = Preporuka * K & " kcal"
            .Label284.Text = Preporuka * K

            'Dijabeticke dijete
            If .Label179.Text >= 16 And .Label179.Text <= 20 Then
                If EnergVrJelovnika < 1900 And EnergVrJelovnika > 500 Then Preporuka = 70
                If EnergVrJelovnika >= 1900 And EnergVrJelovnika < 2100 Then Preporuka = 80
                If EnergVrJelovnika >= 2100 And EnergVrJelovnika < 2300 Then Preporuka = 100
                If EnergVrJelovnika >= 2300 And EnergVrJelovnika < 2500 Then Preporuka = 110
                If EnergVrJelovnika >= 2500 And EnergVrJelovnika < 2700 Then Preporuka = 120
                If EnergVrJelovnika >= 2700 And EnergVrJelovnika < 2900 Then Preporuka = 140
                If EnergVrJelovnika >= 2900 And EnergVrJelovnika < 3100 Then Preporuka = 160
                If EnergVrJelovnika >= 3100 And EnergVrJelovnika < 3300 Then Preporuka = 180
                If EnergVrJelovnika >= 3300 And EnergVrJelovnika < 3500 Then Preporuka = 200
                If EnergVrJelovnika >= 3500 And EnergVrJelovnika < 3700 Then Preporuka = 220
                If EnergVrJelovnika >= 3700 And EnergVrJelovnika < 3900 Then Preporuka = 230
                If EnergVrJelovnika >= 3900 Then Preporuka = 240
                .Label280.Text = Preporuka * K & " kcal"
                .Label284.Text = Preporuka * K
            End If

            'Jelovnik u zavrsnoj fazi deponiranja glikogena u misice kod sportasa.
            If .Label179.Text = 6 Then
                If EnergVrJelovnika < 1900 And EnergVrJelovnika > 500 Then Preporuka = 200
                If EnergVrJelovnika >= 1900 And EnergVrJelovnika < 2100 Then Preporuka = 300
                If EnergVrJelovnika >= 2100 And EnergVrJelovnika < 2300 Then Preporuka = 400
                If EnergVrJelovnika >= 2300 And EnergVrJelovnika < 2500 Then Preporuka = 450
                If EnergVrJelovnika >= 2500 And EnergVrJelovnika < 2700 Then Preporuka = 530
                If EnergVrJelovnika >= 2700 And EnergVrJelovnika < 2900 Then Preporuka = 600
                If EnergVrJelovnika >= 2900 And EnergVrJelovnika < 3100 Then Preporuka = 750
                If EnergVrJelovnika >= 3100 And EnergVrJelovnika < 3300 Then Preporuka = 850
                If EnergVrJelovnika >= 3300 And EnergVrJelovnika < 3500 Then Preporuka = 950
                If EnergVrJelovnika >= 3500 And EnergVrJelovnika < 3700 Then Preporuka = 1200
                If EnergVrJelovnika >= 3700 And EnergVrJelovnika < 3900 Then Preporuka = 1400
                If EnergVrJelovnika >= 3900 Then Preporuka = 1500
                .Label280.Text = Preporuka * K & " kcal"
                .Label284.Text = Preporuka * K
            End If

            'Lakootovo - vegetarijanska dijeta
            If .Label179.Text = 21 Then   'provjetiri
                If EnergVrJelovnika < 1900 And EnergVrJelovnika > 500 Then Preporuka = 195
                If EnergVrJelovnika >= 1900 And EnergVrJelovnika < 2100 Then Preporuka = 267
                If EnergVrJelovnika >= 2100 And EnergVrJelovnika < 2300 Then Preporuka = 290
                If EnergVrJelovnika >= 2300 And EnergVrJelovnika < 2500 Then Preporuka = 362
                If EnergVrJelovnika >= 2500 And EnergVrJelovnika < 2700 Then Preporuka = 410
                If EnergVrJelovnika >= 2700 And EnergVrJelovnika < 2900 Then Preporuka = 426
                If EnergVrJelovnika >= 2900 And EnergVrJelovnika < 3100 Then Preporuka = 512
                If EnergVrJelovnika >= 3100 And EnergVrJelovnika < 3300 Then Preporuka = 642
                If EnergVrJelovnika >= 3300 And EnergVrJelovnika < 3500 Then Preporuka = 700
                If EnergVrJelovnika >= 3500 And EnergVrJelovnika < 3700 Then Preporuka = 750
                If EnergVrJelovnika >= 3700 And EnergVrJelovnika < 3900 Then Preporuka = 780
                If EnergVrJelovnika >= 3900 Then Preporuka = 830
                .Label280.Text = Preporuka * K & " kcal"
                .Label284.Text = Preporuka * K
            End If

        End With
    End Sub
End Module
