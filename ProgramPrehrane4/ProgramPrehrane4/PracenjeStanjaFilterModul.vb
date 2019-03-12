Module PracenjeStanjaFilterModul
    Sub PracenjeStanjaFilter()
        On Error Resume Next
        With Form1
          .KorisniciPracenjeStanjaBindingSource.RemoveFilter()
            .KorisniciPracenjeStanjaBindingSource.Filter = "Korisnik='" & .TextBox14.Text & " " & .TextBox15.Text & "'"
        End With
    End Sub
End Module
