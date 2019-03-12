Module BazaNazivaJelovnikaFilterModul
    Sub BazaNazivaJelovnikaFilter()
        On Error Resume Next
        With Form1
            'Moji jelovnici
            If .RadioButton20.Checked = True Then
                .BazaJelovnikaBindingSource.RemoveFilter()
                .BazaJelovnikaBindingSource.Filter = "Korisnik='" & .Label181.Text & "'" & _
                "AND NazivJelovnika='" & .Label186.Text & "'" & _
                "AND EnergetskaVrijednostJelovnika_kcal='" & .Label187.Text & "'"
            End If
            'Primjeri(jelovnika)
            If .RadioButton21.Checked = True Then
                .BazaPrimjeraJelovnikaBindingSource.RemoveFilter()
                .BazaPrimjeraJelovnikaBindingSource.Filter = "Korisnik='" & .Label181.Text & "'" & _
               "AND NazivJelovnika='" & .Label186.Text & "'" & _
              "AND EnergetskaVrijednostJelovnika_kcal='" & .Label187.Text & "'"
            End If
        End With

    End Sub
End Module
