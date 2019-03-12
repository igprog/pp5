Module EnergetskaPotrosnjaModul
    Sub EnergetskaPotrosnja()
        With Form1
            On Error Resume Next
            .Label355.Visible = True
            .Label334.Visible = True
            .Label354.Visible = True
            .Label346.Visible = True
            .Label353.Visible = True

            Dim BS As BindingSource = .BazaEnergetskePotrosnjeBindingSource
            Dim DGV As DataGridView = Form1.DataGridView16
            Dim i As Integer = DGV.CurrentRow.Index
            Dim Faktor As Double = DGV.Rows(i).Cells(3).Value
            Dim Masa As Double = .ComboBox3.Text
            Dim Minuta As Double = .Label336.Text
            Dim EnergetskaPotrosnja As Integer = (Minuta * Faktor * Masa) / 60

            BS.MoveLast()
            If .DataGridView17.RowCount <= 1 Then BS.AddNew()

          
            .Label334.Text = .TextBox1.Text & " " & .TextBox2.Text   'klijent
            .Label355.Text = .ComboBox14.Text    'Dan
            .Label346.Text = DGV.Rows(i).Cells(1).Value    'Tjelesna aktivnost
            .Label354.Text = Minuta   'Vrijeme (min.)
            .Label353.Text = EnergetskaPotrosnja   '(Energetska potrosnja (kcal)


            BS.MoveLast()
            BS.AddNew()

            .TextBox83.Text = .ComboBox19.Text   'Do sat
            .TextBox84.Text = .ComboBox20.Text   'Od min


            'UKUPNO
            .Label343.Text = "Ukupno:"
            Dim j As Integer
            Dim DGV1 As DataGridView = .DataGridView17
            Dim Min As Double = 0
            Dim Energ As Double = 0

            For j = 0 To DGV1.RowCount - 1
                Min = Min + DGV1.Rows(j).Cells(9).Value   'ukupno minuta
                Energ = Energ + DGV1.Rows(j).Cells(10).Value   'ukupna energetska potrosnja
            Next j

            .Label343.Text = "Ukupno: " & Format(Min / 60, "0.0") & "h          " & Energ & " kcal"   'ukupna energetska potrošnja (24h)


            .DataGridView17.CurrentRow.Selected = False

            .Label355.Visible = False
            .Label334.Visible = False
            .Label354.Visible = False
            .Label346.Visible = False
            .Label353.Visible = False

        End With
    End Sub
End Module
