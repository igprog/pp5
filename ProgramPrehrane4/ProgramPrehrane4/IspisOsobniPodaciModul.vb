Module IspisOsobniPodaciModul
    Sub IspisOsobniPodaci()
        With Form1
            .ListBox55.Items.Clear()
            .ListBox55.Items.Add("Klijent: " & .TextBox1.Text & " " & .TextBox2.Text)
            .ListBox55.Items.Add("Dob: " & .ComboBox1.Text & " god.")
            If .RadioButton1.Checked = True Then .ListBox55.Items.Add("Spol: Muškarac")
            If .RadioButton2.Checked = True Then .ListBox55.Items.Add("Spol: Žena")
            .ListBox55.Items.Add("Visina: " & .ComboBox2.Text & " cm")
            .ListBox55.Items.Add("Masa: " & .ComboBox3.Text & " kg")
            .ListBox55.Items.Add("Opseg struka: " & .ComboBox6.Text & " cm")
            .ListBox55.Items.Add("Opseg bokova: " & .ComboBox7.Text & " cm")
            If .RadioButton3.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti na poslu: izrazito slab")
            If .RadioButton4.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti na poslu: slab")
            If .RadioButton5.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti na poslu: umjeren")
            If .RadioButton6.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti na poslu: izražen")
            If .RadioButton7.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti izvan posla: izrazito slab")
            If .RadioButton8.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti izvan posla: slab")
            If .RadioButton9.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti izvan posla: umjeren")
            If .RadioButton10.Checked = True Then .ListBox55.Items.Add("Intenzitet aktivnosti izvan posla: izražen")
            .ListBox55.Items.Add("Napomena: " & .TextBox75.Text)
            .ListBox55.Items.Add("--------------------------------------------------------------------------")
            .ListBox55.Items.Add("IZRAČUN")
            .ListBox55.Items.Add("Indeks tjelesne mase (BMI): " & .Label11.Text & " kg/m2")
            .ListBox55.Items.Add("Primjerena tjelesna masa: " & .Label12.Text & " kg")
            .ListBox55.Items.Add("Ukupna energetska potrošnja: " & .Label13.Text & " kcal")
            .ListBox55.Items.Add("Omjer opsega struka i bokova (WHR): " & .Label193.Text)
            .ListBox55.Items.Add(.Label194.Text)
            .ListBox55.Items.Add(.Label195.Text)
            .ListBox55.Items.Add("Preporučeni energetski unos: " & .TextBox3.Text & " kcal")
            .ListBox55.Items.Add("Dodatna energetska potoršnja: " & .TextBox4.Text & " kcal")

            .TextBox69.Text = "PODACI O KLIJENTU" & vbCrLf
            Dim a As Integer
            For a = 0 To .ListBox55.Items.Count - 1
                If .ListBox55.Items(a).ToString <> "" Then
                    .TextBox69.Text = .TextBox69.Text & vbCrLf & vbCrLf & .ListBox55.Items(a)
                    .RichTextBoxPrintCtrl1.Text = .TextBox69.Text
                End If
            Next

        End With
    End Sub
End Module
