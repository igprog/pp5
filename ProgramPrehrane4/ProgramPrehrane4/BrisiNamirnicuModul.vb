Module BrisiNamirnicuModul
    Sub BrisiNamirnicu()
        On Error Resume Next
        With Form1
            'Dorucak
            If .TabPage7.CanFocus = True Then
                .DataGridView5.Rows.Remove(.DataGridView5.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If
            'Jutarnja uzina
            If .TabPage8.CanFocus = True Then
                .DataGridView9.Rows.Remove(.DataGridView9.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If
            'rucak
            If .TabPage9.CanFocus = True Then
                .DataGridView11.Rows.Remove(.DataGridView11.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If
            'popodnevna uzina
            If .TabPage10.CanFocus = True Then
                .DataGridView12.Rows.Remove(.DataGridView12.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If
            'vecera
            If .TabPage11.CanFocus = True Then
                .DataGridView13.Rows.Remove(.DataGridView13.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If
            'obrok pred spavanje
            If .TabPage12.CanFocus = True Then
                .DataGridView14.Rows.Remove(.DataGridView14.CurrentRow)
                .DorucakBindingSource.MoveLast()
            End If

            .DataGridView5.CurrentRow.Selected = False
            .DataGridView9.CurrentRow.Selected = False
            .DataGridView11.CurrentRow.Selected = False
            .DataGridView12.CurrentRow.Selected = False
            .DataGridView13.CurrentRow.Selected = False
            .DataGridView14.CurrentRow.Selected = False

           

        End With
    End Sub
End Module
