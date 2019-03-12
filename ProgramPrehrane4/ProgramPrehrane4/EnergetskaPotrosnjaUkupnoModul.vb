Module EnergetskaPotrosnjaUkupnoModul
    Sub EnergetskaPotrosnjaUkupno()
        On Error Resume Next
        With Form1
            'Ukupne vrijednosti
            Dim i As Integer
            Dim DGV As DataGridView
            DGV = .DataGridView17
            Dim Min As Double = 0
            Dim Energ As Double = 0
            .Label315.Text = 0

            For i = 0 To DGV.RowCount - 1
                Min = Min + DGV.Rows(i).Cells(9).Value
                Energ = Energ + DGV.Rows(i).Cells(10).Value   'ukupna dodatna potrosnja
            Next i

            If Min > 60 * 24 Then
                MsgBox("Error.")
                '   .ComboBox14.SelectedIndex = .ComboBox14.SelectedIndex + 1
                Exit Sub
            End If

            .Label343.Text = "Ukupno: " & Format(Min / 60, "0.0") & " h,  " & Format(Energ, "0") & " kcal"   'ukupne vrijednosti
            '  .Label343.Text = "Ukupno: " & Min & " min,  " & Energ & " kcal"   'ukupne vrijednosti
            If Min = 60 * 24 Then
                .Label315.Text = Energ
            End If

        End With
    End Sub
End Module
