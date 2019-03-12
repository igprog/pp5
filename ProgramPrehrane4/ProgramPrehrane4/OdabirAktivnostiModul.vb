Module OdabirAktivnostiModul
    Sub OdabirAktivnosti()
        On Error Resume Next
        With Form1
            If .ComboBox5.Text = "" Then Exit Sub 'vrijeme trajanja aktivnosti
            If .TextBox9.Text = "" Or .TextBox10.Text = "" Then
                Exit Sub
            End If

            If Val(.Label176.Text) = Val(.TextBox10.Text) Then
                MsgBox("Energetska vrijednost odabranih aktivnosti je jednaka preporučenoj dodatnoj energetskoj potrošnji.")
                Exit Sub
            End If

            .OdabranaDodatnaTjelesnaAktivnostBindingSource.MoveLast()

            '           If ComboBox1.Text = "" Then
            'MsgBox("Unesite vrijeme trajanja aktivnosti.")
            '            Exit Sub
            '           End If

            '           Label1.Visible = True
            '          Label2.Visible = True
            '         Label3.Visible = True
            '        Label11.Visible = True
            '       Label12.Visible = True
            '      Label13.Visible = True



            If .DataGridView4.RowCount <= 1 Then .OdabranaDodatnaTjelesnaAktivnostBindingSource.AddNew()
            'Dim DodatnaEnergetskaPotrosnja As Integer
            Dim FaktorTjelesneAktivnosti As Double = .Label94.Text
            Dim Masa As Double = .TextBox9.Text
            Dim Vrijeme As Double = .ComboBox5.Text
            .Label98.Text = .Label20.Text    'Tjelesna aktivnost
            'DodatnaEnergetskaPotrosnja = TextBox1.Text
            '  FaktorTjelesneAktivnosti = Label3.Text
            '       Masa = TextBox2.Text
            '          Vrijeme = ComboBox1.Text
            'Vrijeme = (DodatnaEnergetskaPotrosnja / (FaktorTjelesneAktivnosti * Masa)) * 60
            .Label99.Text = Vrijeme
            .Label100.Text = Format(((Vrijeme * FaktorTjelesneAktivnosti * Masa) / 60), "0")  'Energija


            '       Label11.Text = Label1.Text

            .OdabranaDodatnaTjelesnaAktivnostBindingSource.MoveLast()
            .OdabranaDodatnaTjelesnaAktivnostBindingSource.AddNew()


            Dim i As Integer
            Dim DGV As DataGridView
            DGV = .DataGridView4
            Dim Energ As Double = 0

            For i = 0 To DGV.RowCount - 1
                Energ = Energ + DGV.Rows(i).Cells(7).Value   'ukupna dodatna potrosnja
            Next i

            .Label301.Text = "Ukupno: " & Energ & " kcal"   'ukupna dodatna energetska potrosnja

            '           Label1.Visible = False
            '          Label2.Visible = False
            '         Label3.Visible = False
            '        Label11.Visible = False
            '       Label12.Visible = False
            '      Label13.Visible = False

            .Label176.Text = Energ
            If Energ > Val(.TextBox10.Text) + 20 Then
                MsgBox("Energetska vrijednost odabranih aktivnosti veća od preporučene dodatne energetska potrošnje.")
            End If

            '          ComboBox1.Text = ""    'vrijeme
            '         DodatnaAktivnost()

            .DataGridView4.CurrentRow.Selected = False
            '      TextBox3.Text = ""
            '     TextBox3.Select()
            DodatnaTjelesnaAktivnost()

            '    .TjelesneAktivnostiBindingSource.RemoveFilter()
            .SportskeAktivnostiBindingSource.RemoveFilter()

        End With

    End Sub
End Module
