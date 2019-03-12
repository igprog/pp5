Module DodatnaEnergetskaPotrosnjaPreporukaModul
    Sub DodatnaEnergetskaPotrosnjaPreporuka()
        With Form1
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

            'CILJ
            .Label196.Text = DodatnaEnergetskaPotrosnja
            .Label197.Text = DodatniEnergetskiUnos

        End With
    End Sub
End Module
