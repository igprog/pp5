Module DemoModul
    Sub Demo()
        With Form1
            .PriručnikToolStripMenuItem1.Enabled = False
            .Button39.Enabled = False

            .TextBox77.Enabled = False           'cijene
            .Button43.Enabled = False         ' cijene
            .ComboBox26.Enabled = False    'cijena (postvke)
            .Button22.Enabled = False       'moje namirnice
            .GroupBox56.Enabled = False    'postavke ispisa
            .CheckBox11.Checked = False    'nutritivna i energetska vrijednost obroka - ispis
            .CheckBox13.Checked = False    'cijena jelovnika - ispis
            .ComboBox16.Enabled = False    'broj korisnika jelovnika
            .Button44.Enabled = False      'Energetska potrosnja
            .CheckBox1.Enabled = False     'Favoriti
            .Button20.Enabled = False      'spemi klijenta
            .Button17.Enabled = False       'pracenje antropometrijskih parametara

            'vrste prehrane
            Dim DGV As DataGridView = .DataGridView1
            Dim i As Integer = DGV.CurrentRow.Index
            If i > 6 Or i = 0 Or i = 1 Or i = 3 Or i = 4 Then
                .GroupBox10.Visible = False
                If MessageBox.Show("Preporuke za odabranu vrstu prehrane su dostupne samo u PREMIUM verziji." _
                                      & vbCrLf & "Želite li naručiti aktivacijski kod za pokretanje PREMIUM verzije programa.", "Program Prehrane 5.0", _
                                       MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                    = DialogResult.Yes Then
                    System.Diagnostics.Process.Start("http://www.programprehrane.com/pp5kod.htm")
                End If
                '  DGV.Rows(2).Selected = True   'normalna prehrana  
                .VrstaDijeteBindingSource.Position = 2
                .GroupBox10.Visible = True
            Else
                .GroupBox10.Visible = True
            End If

            'baza jelovnika
            If .RadioButton21.Checked = True Then
                Dim DGV1 As DataGridView = .DataGridView7
                Dim j As Integer = DGV1.CurrentRow.Index
                If j > 0 Then
                    '  .GroupBox10.Visible = False
                    If MessageBox.Show("Odabrani jelovnik je dostupan samo u PREMIUM verziji." _
                                          & vbCrLf & "Želite li naručiti aktivacijski kod za pokretanje PREMIUM verzije?", "Program Prehrane 5.0", _
                                           MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                        = DialogResult.Yes Then
                        System.Diagnostics.Process.Start("http://www.programprehrane.com/pp5kod.htm")
                    End If
                    ' DGV1.Rows(0).Selected = True   'vrati na prvi jelovnik 
                    '   .VrstaDijeteBindingSource.Position = 2
                    '   .GroupBox10.Visible = True
                    '  Else
                    '      .GroupBox10.Visible = True
                End If
            End If

        End With
    End Sub
End Module
