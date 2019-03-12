Module PremiumModul
    Sub Premium()
        With Form1
            .PriručnikToolStripMenuItem1.Enabled = True  'prirucnik
            .Button39.Enabled = True                   'prirucnik

            .TextBox77.Enabled = True           'cijene
            .Button43.Enabled = True        ' cijene
            .ComboBox26.Enabled = True    'cijena (postvke)
            .Button22.Enabled = True    'moje namirnice
            .GroupBox56.Enabled = True   'postavke ispisa
            .CheckBox11.Checked = True    'nutritivna i energetska vrijednost obroka - ispis
            .CheckBox13.Checked = True   'cijena jelovnika - ispis
            .ComboBox16.Enabled = True    'broj korisnika jelovnika
            .Button44.Enabled = True      'Energetska potrosnja
            .CheckBox1.Enabled = True     'Favoriti
            .Button20.Enabled = True      'spemi korisnika
            .Button17.Enabled = True       'pracenje antropometrijskih parametara

        End With
    End Sub
End Module
