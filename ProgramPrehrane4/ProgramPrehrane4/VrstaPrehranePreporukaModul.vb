Module VrstaPrehranePreporukaModul
    Sub VrstaPrehranePreporuka()
        With Form1

            .DataGridView1.CurrentRow.Selected = False   'vrsta prehrane/dijete

            'DJECA OD 9 DO 14 GOD
            Dim Dob As Double = Val(.ComboBox1.Text)   'dob
            If Dob >= 9 And Dob < 14 Then   'Djeca - Preporuke nutrijenti
                If .RadioButton17.Checked = True Then
                    .DataGridView1.Rows(0).Selected = True   'normalna prehrana (9-18god)
                    .VrstaDijeteBindingSource.Position = 0
                End If
                If .RadioButton15.Checked = True Then
                    .DataGridView1.Rows(3).Selected = True   'redukcija tjelesne mase (redukcijska dijeta 9-18god)
                    .VrstaDijeteBindingSource.Position = 3
                End If
            End If
            'DJECA OD 14 DO 18 GOD
            If Dob >= 14 And Dob < 18 Then   'Djeca - Preporuke nutrijenti
                If .RadioButton17.Checked = True Then
                    .DataGridView1.Rows(1).Selected = True   'normalna prehrana (9-18god)
                    .VrstaDijeteBindingSource.Position = 1
                End If
                If .RadioButton15.Checked = True Then
                    .DataGridView1.Rows(4).Selected = True   'redukcija tjelesne mase (redukcijska dijeta 9-18god)
                    .VrstaDijeteBindingSource.Position = 4
                End If
            End If


            'PUNOLJETNE OSOBE
            If Dob >= 18 Then
                If .RadioButton17.Checked = True Then
                    .DataGridView1.Rows(2).Selected = True   'normalna prehrana  
                    .VrstaDijeteBindingSource.Position = 2
                End If
                If .RadioButton15.Checked = True Then
                    .DataGridView1.Rows(5).Selected = True   'redukcija tjelesne mase (redukcijska dijeta)
                    .VrstaDijeteBindingSource.Position = 5
                End If
                If .RadioButton16.Checked = True Or .RadioButton18.Checked = True Then
                    .DataGridView1.Rows(6).Selected = True   'prehrana za povecanje tjelesne mase
                    .VrstaDijeteBindingSource.Position = 6
                End If
            End If


        End With
    End Sub
End Module
