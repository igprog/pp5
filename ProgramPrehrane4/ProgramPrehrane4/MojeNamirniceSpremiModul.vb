Module MojeNamirniceSpremiModul
    Sub MojeNamirniceSpremi()
        On Error Resume Next
        With Form1


            If .TextBox19.Text = "" Then
                MsgBox("Unesite naziv namirnice.")
                Exit Sub
            End If
            ' If .ComboBox4.Text = "" Then
            'MsgBox("Odaberite skupinu namirnica.")
            'Exit Sub
            'End If
            If .TextBox21.Text = "" Then
                MsgBox("Unesite masu namirnice u gramima.")
                Exit Sub
            End If
            If .TextBox22.Text = "" Then
                MsgBox("Unesite energetsku vrijednost namirnice u kcal.")
                Exit Sub
            End If
            If .TextBox23.Text = "" Then
                MsgBox("Unesite masu ugljikohidrata u gramima.")
                Exit Sub
            End If
            If .TextBox24.Text = "" Then
                MsgBox("Unesite masu bjelančevina u gramima.")
                Exit Sub
            End If
            If .TextBox25.Text = "" Then
                MsgBox("Unesite masu masti u gramima.")
                Exit Sub
            End If

            Dim BS As BindingSource = .MojeNamirniceBindingSource

            .Label26.Text = Val(.TextBox22.Text) * 4.15   'Energija kJ

            If .TextBox19.Text = "" Then .TextBox19.Text = "N"
            If .ComboBox8.Text = "" Then .ComboBox8.Text = "N"
            If .TextBox20.Text = "" Then .TextBox20.Text = "N"
            If .TextBox21.Text = "" Then .TextBox21.Text = "N"
            If .Label26.Text = "" Then .Label26.Text = "N"
            If .TextBox22.Text = "" Then .TextBox22.Text = "N"
            If .TextBox23.Text = "" Then .TextBox23.Text = "N"
            If .TextBox24.Text = "" Then .TextBox24.Text = "N"
            If .TextBox25.Text = "" Then .TextBox25.Text = "N"
            If .TextBox26.Text = "" Then .TextBox26.Text = "N"
            If .TextBox27.Text = "" Then .TextBox27.Text = "N"
            If .TextBox28.Text = "" Then .TextBox28.Text = "N"
            If .TextBox29.Text = "" Then .TextBox29.Text = "N"
            If .TextBox30.Text = "" Then .TextBox30.Text = "N"
            If .TextBox31.Text = "" Then .TextBox31.Text = "N"
            If .TextBox32.Text = "" Then .TextBox32.Text = "N"
            If .TextBox33.Text = "" Then .TextBox33.Text = "N"
            If .TextBox34.Text = "" Then .TextBox34.Text = "N"
            If .TextBox35.Text = "" Then .TextBox35.Text = "N"
            If .TextBox36.Text = "" Then .TextBox36.Text = "N"
            If .TextBox37.Text = "" Then .TextBox37.Text = "N"
            If .TextBox38.Text = "" Then .TextBox38.Text = "N"
            If .TextBox39.Text = "" Then .TextBox39.Text = "N"
            If .TextBox40.Text = "" Then .TextBox40.Text = "N"
            If .TextBox41.Text = "" Then .TextBox41.Text = "N"
            If .TextBox42.Text = "" Then .TextBox42.Text = "N"
            If .TextBox43.Text = "" Then .TextBox43.Text = "N"
            If .TextBox44.Text = "" Then .TextBox44.Text = "N"
            If .TextBox45.Text = "" Then .TextBox45.Text = "N"
            If .TextBox46.Text = "" Then .TextBox46.Text = "N"
            If .TextBox47.Text = "" Then .TextBox47.Text = "N"
            If .TextBox48.Text = "" Then .TextBox48.Text = "N"
            If .TextBox49.Text = "" Then .TextBox49.Text = "N"
            If .TextBox50.Text = "" Then .TextBox50.Text = "N"
            If .TextBox51.Text = "" Then .TextBox51.Text = "N"
            If .TextBox52.Text = "" Then .TextBox52.Text = "N"
            If .TextBox53.Text = "" Then .TextBox53.Text = "N"
            If .TextBox54.Text = "" Then .TextBox54.Text = "N"
            If .TextBox55.Text = "" Then .TextBox55.Text = "N"
            If .TextBox56.Text = "" Then .TextBox56.Text = "N"
            If .TextBox57.Text = "" Then .TextBox57.Text = "N"
            If .TextBox58.Text = "" Then .TextBox58.Text = "N"
            If .TextBox59.Text = "" Then .TextBox59.Text = "N"
            If .TextBox60.Text = "" Then .TextBox60.Text = "N"
            If .TextBox61.Text = "" Then .TextBox61.Text = "N"
            If .TextBox62.Text = "" Then .TextBox62.Text = "N"
            If .TextBox71.Text = "" Then .TextBox71.Text = "N"
            If .TextBox72.Text = "" Then .TextBox72.Text = "N"

            .Label374.Text = "MojeNamirnice"    'skupina namirnica
            .Label375.Text = "N"

            .Label367.Text = 0                'zitarice
            .Label368.Text = 0     'povrce
            .Label369.Text = 0   'voce
            .Label370.Text = 0    'povrce
            .Label371.Text = 0     'meso
            .Label372.Text = 0     'mlijeko
            .Label373.Text = 0     'masti

            '      If .ComboBox4.Text = "Žitarice" Then
            '.Label367.Text = 1
            '.ComboBox4.Text = "Zitarice"
            'Else
            ' .Label367.Text = 0
            'End If
            'If .ComboBox4.Text = "Povrće" Then
            '.Label368.Text = 1
            '.ComboBox4.Text = "Povrce"
            'Else
            ' .Label368.Text = 0
            'End If
            'If .ComboBox4.Text = "Voće" Then
            '.Label369.Text = 1
            '.ComboBox4.Text = "Voce"
            'Else
            '.Label369.Text = 0
            'End If
            'If .ComboBox4.Text = "Izrazito nemasno meso i zamjene" Then
            '.Label370.Text = 1
            '.ComboBox4.Text = "IzrazitoNemasnoMeso"
            'Else
            '.Label370.Text = 0
            ''End If
            'If .ComboBox4.Text = "Nemasno meso i zamjene" Then
            '.Label370.Text = 1
            '.ComboBox4.Text = "NemasnoMeso"
            ' '  Else
            '  '  .Label370.Text = 0
            'End If
            'If .ComboBox4.Text = "Srednje masno meso i zamjene" Then
            '.Label370.Text = 1
            '.ComboBox4.Text = "SrednjeMasnoMeso"
            ' '  Else
            '  '  .Label370.Text = 0
            'End If
            'If .ComboBox4.Text = "Masno meso i zamjene" Then
            '.Label370.Text = 1
            '.ComboBox4.Text = "MasnoMeso"
            ''   Else
            ' '   .Label370.Text = 0
            'End If



            BS.MoveLast()
            BS.AddNew()
            MsgBox("Spremljeno u bazu. Namirnica se nalazi u skupini (Moje namirnice).")
            .DataGridView2.CurrentRow.Selected = False

        End With
    End Sub
End Module
