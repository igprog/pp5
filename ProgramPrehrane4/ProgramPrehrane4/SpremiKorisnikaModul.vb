Module SpremiKorisnikaModul
    Sub SpremiKorisnika()
        '      On Error Resume Next
        With Form1

            If .TextBox1.Text = "" And .TextBox2.Text = "" Then Exit Sub
            ' .TabControl3.SelectedIndex = 4   'Baza korisnka
            Dim DGV As DataGridView = .DataGridView6
            If DGV.RowCount < 1 Then .BazaKorisnikaBindingSource.AddNew()

            'Provjera dali vec postoji korisnik u bazi
            Dim i As Integer
            For i = 0 To DGV.RowCount - 1
                If DGV.Rows(i).Cells(2).Value IsNot DBNull.Value _
                  And DGV.Rows(i).Cells(3).Value IsNot DBNull.Value Then
                    If DGV.Rows(i).Cells(2).Value = .TextBox1.Text _
                               And DGV.Rows(i).Cells(3).Value = .TextBox2.Text Then
                        MsgBox("Korisnik " & .TextBox1.Text & " " & .TextBox2.Text & " već postoji u bazi.")
                        Exit Sub
                    End If
                End If
            Next i

            .BazaKorisnikaBindingSource.MoveLast()

            BMI()    'Izračun BMI, TEE, WHR

            .Label210.Text = .TextBox1.Text & " " & .TextBox2.Text    'Korisnik
            .Label211.Text = .TextBox1.Text   'Ime
            .Label212.Text = .TextBox2.Text     'Prezime
            .Label213.Text = Val(.ComboBox1.Text)       'Dob
            If .RadioButton1.Checked = True Then
                .Label178.Text = "Muškarac"       'Spol
            End If
            If .RadioButton2.Checked = True Then
                .Label178.Text = "Žena"       'Spol
            End If
            .Label250.Text = .ComboBox2.Text       'Visina
            .Label251.Text = .ComboBox3.Text       'Masa
            .Label252.Text = .ComboBox6.Text       'Opseg struka
            .Label253.Text = .ComboBox7.Text       'Opseg bokova
            .Label289.Text = .Label193.Text       'WHR

            '          If Val(.ComboBox1.Text) < 18 And Val(.ComboBox1.Text) >= 9 Then   'djeca
            'If .RadioButton3.Checked = True Then
            '.Label292.Text = "Izrazito slab"       'Intenzitet aktivnosti (djeca)
            'End If
            'If .RadioButton4.Checked = True Then
            '.Label292.Text = "Slab"       'Intenzitet aktivnosti na poslu (djeca)
            'End If
            'If .RadioButton5.Checked = True Then
            '.Label292.Text = "Umjeren"       'Intenzitet aktivnosti na poslu (djeca)
            'End If
            'If .RadioButton6.Checked = True Then
            '.Label292.Text = "Izražen"       'Intenzitet aktivnosti na poslu (djeca)
            'End If
            'Else    'odrasli
            If .RadioButton3.Checked = True Then
                .Label290.Text = "Izrazito slab"       'Intenzitet aktivnosti na poslu
            End If
            If .RadioButton4.Checked = True Then
                .Label290.Text = "Slab"       'Intenzitet aktivnosti na poslu
            End If
            If .RadioButton5.Checked = True Then
                .Label290.Text = "Umjeren"       'Intenzitet aktivnosti na poslu
            End If
            If .RadioButton6.Checked = True Then
                .Label290.Text = "Izražen"       'Intenzitet aktivnosti na poslu
            End If

            If .RadioButton7.Checked = True Then
                .Label291.Text = "Izrazito slab"       'Intenzitet aktivnosti izvan posla
            End If
            If .RadioButton8.Checked = True Then
                .Label291.Text = "Slab"       'Intenzitet aktivnosti izvan posla
            End If
            If .RadioButton9.Checked = True Then
                .Label291.Text = "Umjeren"       'Intenzitet aktivnosti izvan posla
            End If
            If .RadioButton10.Checked = True Then
                .Label291.Text = "Izražen"       'Intenzitet aktivnosti izvan posla
            End If

            ' End If

            .Label293.Text = .Label11.Text       'BMI
            .Label294.Text = .Label13.Text       'TEE
            .Label295.Text = .TextBox4.Text       'Dodatna energetska potrosnja
            .Label296.Text = .TextBox75.Text       'Napomena
            .Label297.Text = Date.Today       'Datum

            .BazaKorisnikaBindingSource.AddNew()

            ' .TabControl3.SelectedIndex = 0

        End With
    End Sub
End Module
