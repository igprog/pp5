Module NoviKorisnikModul
    Sub NoviKorisnik()
        On Error Resume Next
        With Form1
            .TabControl1.SelectedIndex = 0
            .TabControl3.SelectedIndex = 0
            .TextBox1.Text = ""
            .TextBox2.Text = ""
            .ComboBox1.Text = ""
            .RadioButton1.Checked = False
            .RadioButton2.Checked = False
            .ComboBox2.Text = ""
            .ComboBox3.Text = ""
            .ComboBox6.Text = ""
            .ComboBox7.Text = ""
            .RadioButton3.Checked = False
            .RadioButton4.Checked = False
            .RadioButton5.Checked = False
            .RadioButton6.Checked = False
            .RadioButton7.Checked = False
            .RadioButton8.Checked = False
            .RadioButton9.Checked = False
            .RadioButton10.Checked = False
            .TextBox75.Text = ""
            .TextBox1.Select()

            'Izracun ""
            .Label11.Text = ""
            .Label12.Text = ""
            .Label13.Text = ""
            .Label193.Text = ""
            .Label208.Text = ""
            .RadioButton15.Checked = False
            .RadioButton16.Checked = False
            .RadioButton17.Checked = False
            .RadioButton18.Checked = False
            .TextBox3.Text = ""
            .TextBox4.Text = ""
            .Label25.Text = ""
            .Label194.Text = ""
            .Label195.Text = ""
            .PictureBox5.Location = New Point(6, 26)

            'Briši dodatnu tjelesnu aktivnost
            Dim DGV As DataGridView = .DataGridView4
            Dim BS As BindingSource = .OdabranaDodatnaTjelesnaAktivnostBindingSource
            For i = 0 To DGV.RowCount - 1
                DGV.Rows.Remove(DGV.CurrentRow)
            Next i
            BS.AddNew()
            Dim Energ As Double = 0
            For i = 0 To DGV.RowCount - 1
                Energ = Energ + DGV.Rows(i).Cells(7).Value   'ukupna dodatna potrosnja
            Next i
            .Label176.Text = Energ
            .Label301.Text = "Ukupno: " & Energ & " kcal"   'ukupna dodatna energetska potrosnja

            DodatnaTjelesnaAktivnost()
            .ComboBox5.Text = ""  'vrijeme trajanja aktivnost

        End With
    End Sub
End Module
