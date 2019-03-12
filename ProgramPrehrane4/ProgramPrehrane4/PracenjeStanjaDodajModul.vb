Module PracenjeStanjaDodajModul
    Sub PracenjeStanjaDodaj()
        With Form1
            .KorisniciPracenjeStanjaBindingSource.RemoveFilter()

            If .TextBox14.Text = "" Or .TextBox17.Text = "" Then GoTo kraj
            If .DataGridView10.RowCount <= 1 Then .KorisniciPracenjeStanjaBindingSource.AddNew()
            .KorisniciPracenjeStanjaBindingSource.MoveLast()

            'Provjera dali vec postoji podatak za taj datum u bazi
            Dim DGV As DataGridView = .DataGridView10
            Dim i As Integer
            For i = 0 To DGV.RowCount - 1
                If DGV.Rows(i).Cells(2).Value IsNot DBNull.Value _
                    And DGV.Rows(i).Cells(3).Value IsNot DBNull.Value _
                      And DGV.Rows(i).Cells(13).Value IsNot DBNull.Value Then

                    If DGV.Rows(i).Cells(2).Value = .TextBox14.Text _
                       And DGV.Rows(i).Cells(3).Value = .TextBox14.Text _
                         And DGV.Rows(i).Cells(13).Value = .DateTimePicker2.Value.Date Then
                        MsgBox("Podaci za " & .DateTimePicker2.Value.Date & " već postoje u bazi.")
                        Exit Sub
                    End If
                End If

            Next i

            'Bindings
            .Label308.Text = .TextBox14.Text & " " & .TextBox15.Text   'Korisnik
            .Label219.Text = .TextBox14.Text   'Ime
            .Label220.Text = .TextBox15.Text    'Prezime
            .Label307.Text = .TextBox76.Text   'Dob
            .Label221.Text = .TextBox16.Text    'Visina
            .Label222.Text = .TextBox17.Text     'Masa
            .Label223.Text = .TextBox18.Text       'Opseg struka
            .Label227.Text = .TextBox63.Text       'Opseg bokova

            ' OMJER OPSEGA STRUKA I BOKOVA
            Dim OpsegStruka As Double = Val(.TextBox18.Text)
            Dim OpsegBokova As Double = Val(.TextBox63.Text)
            Dim WHR As Double = Format(OpsegStruka / OpsegBokova, "0.00")
            .Label302.Text = WHR   'WHR

            Dim Visina As Integer = Val(.TextBox16.Text)
            Dim Masa As Integer = Val(.TextBox17.Text)
            Dim Bmi As Double = Format(Masa / ((Visina / 100) * (Visina / 100)), "0.0")
            .Label303.Text = Bmi  'BMI

            'Primjerena masa
            Dim PrimjerenaMasaMin As Integer = 18.5 * (Visina / 100) * (Visina / 100)
            Dim PrimjerenaMasaMax As Integer = 25 * (Visina / 100) * (Visina / 100)
            .Label304.Text = PrimjerenaMasaMin    'od
            .Label305.Text = PrimjerenaMasaMax    'do

            ' .Label12.Text = Format((18.5 * (Visina / 100) * (Visina / 100)), "0") & " - " & Format((25 * (Visina / 100) * (Visina / 100)), "0")

            .Label224.Text = .DateTimePicker2.Value.Date   'Datum

            '   .KorisniciBindingSource.AddNew()
            .KorisniciPracenjeStanjaBindingSource.AddNew()

            .KorisniciPracenjeStanjaBindingSource.Filter = "Korisnik='" & .TextBox14.Text & " " & .TextBox15.Text & "'"

            '  MsgBox("Spremljeno.")

        End With
        Exit Sub
kraj:
        MsgBox("Polja IME / NAZIV, PREZIME i MASA su obavezna.")

    End Sub
End Module
