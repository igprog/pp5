Module DodatnaTjelesnaAktivnostModul
    Sub DodatnaTjelesnaAktivnost()
        On Error Resume Next
        With Form1
            .Label94.Visible = True
            .Label95.Visible = True
            .Label96.Visible = True
            .Label97.Visible = True
            .Label98.Visible = True
            .Label99.Visible = True
            .Label100.Visible = True

            If Val(.Label176.Text) > Val(.TextBox10.Text) + 20 Then
                '       MsgBox("Odabrana energetska potrošnja veća od preporučene!")
                '          .Label9.Text = ""
                '         .Label1.Visible = False
                '        .Label2.Visible = False
                '       .Label3.Visible = False
                Exit Sub
            End If

            Dim DodatnaEnergetskaPotrosnja As Double = .TextBox10.Text - .Label176.Text
            Dim FaktorTjelesneAktivnosti As Double = .Label94.Text
            Dim Masa As Double = .TextBox9.Text
            Dim Vrijeme As Double
            '          DodatnaEnergetskaPotrosnja = .TextBox1.Text - .Label10.Text
            '  FaktorTjelesneAktivnosti = .Label3.Text
            '  Masa = .TextBox2.Text
            Vrijeme = (DodatnaEnergetskaPotrosnja / (FaktorTjelesneAktivnosti * Masa)) * 60
            
            .Label97.Text = "Vrijeme potrebno za potrošnju " & Format((.TextBox10.Text - .Label176.Text), "0") & " kcal:"

            .ComboBox5.Text = Format(Vrijeme, "0.0")   'minute

            .Label94.Visible = False
            .Label95.Visible = False
            ' .Label96.Visible = False
            .Label98.Visible = False
            .Label99.Visible = False
            .Label100.Visible = False

            '           .Label1.Visible = False
            '          .Label2.Visible = False
            '         .Label3.Visible = False

            '          If .Label1.Text = "Label1" Then
            '           '.Label8.Visible = False
            '          Else
            '         .Label8.Visible = True
            '        End If
            .TextBox8.Text = "Pretraži"

        End With
    End Sub
End Module
