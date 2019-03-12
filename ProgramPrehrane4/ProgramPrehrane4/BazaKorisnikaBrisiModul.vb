Module BazaKorisnikaBrisiModul
    Sub BazaKorisnikaBrisi()
        On Error Resume Next
        With Form1
            If MessageBox.Show("Dali želite izbrisati odabranog korisnika iz baze korisnika?", "Briši korisnika", _
             MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
             = DialogResult.Yes Then

                Dim DGV As DataGridView
                DGV = .DataGridView10         'pracenje stanja
                Dim i As Integer
                For i = 0 To DGV.RowCount - 1
                    If DGV.Rows(i).Cells(1).Value = .Label210.Text Then
                        DGV.Rows.Remove(DGV.CurrentRow)
                    End If
                Next i
                .DataGridView6.Rows.Remove(.DataGridView6.CurrentRow)   'baza korisnika
                .BazaKorisnikaBindingSource.MoveLast()
                .DataGridView6.CurrentRow.Selected = False


                '      .BazaJelovnikaBindingSource.MoveLast()
                '     .DataGridView7.Rows.Remove(.DataGridView7.CurrentRow)
                '    .BazaNazivaJelovnikaBindingSource.MoveLast()

            End If
        End With
    End Sub
End Module
