Module BazaJelovnikaBrisiModul
    Sub BazaJelovnikaBrisi()
        On Error Resume Next
        With Form1
            If .RadioButton21.Checked = True Then
                Exit Sub
            End If

            If MessageBox.Show("Dali želite izbrisati odabrani jelovnik iz baze jelovnika?", "Briši jelovnik", _
       MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
       = DialogResult.Yes Then

                Dim DGV As DataGridView
                DGV = .DataGridView8
                Dim i As Integer
                For i = 0 To DGV.RowCount - 1
                    If DGV.Rows(i).Cells(1).Value = .Label181.Text _
                        And DGV.Rows(i).Cells(6).Value = .Label186.Text _
                        And DGV.Rows(i).Cells(7).Value = .Label187.Text Then
                        DGV.Rows.Remove(DGV.CurrentRow)
                    End If
                Next i

                .BazaJelovnikaBindingSource.MoveLast()
               .DataGridView7.Rows.Remove(.DataGridView7.CurrentRow)
                .BazaNazivaJelovnikaBindingSource.MoveLast()
             
            End If
        End With
    End Sub
End Module
