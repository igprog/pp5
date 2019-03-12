Module NoviJelovnikModul
    Sub NoviJelovnik()
        With Form1
            On Error Resume Next
            Dim DGV As DataGridView
            Dim BS As BindingSource
            Dim i As Integer

            .TabControl2.SelectedIndex = 0

            'briši jelovnik
            Dim b As Integer
            For b = 1 To 6

                If b = 1 Then
                    BS = .DorucakBindingSource
                    DGV = .DataGridView5 'dorucak
                End If
                If b = 2 Then
                    BS = .JutarnjaUzinaBindingSource
                    DGV = .DataGridView9 'jutarnja uzina
                End If
                If b = 3 Then
                    BS = .RucakBindingSource
                    DGV = .DataGridView11 'rucak
                End If
                If b = 4 Then
                    BS = .PopodnevnaUzinaBindingSource
                    DGV = .DataGridView12 'popodnevna uzina
                End If
                If b = 5 Then
                    BS = .VeceraBindingSource
                    DGV = .DataGridView13 'vecera
                End If
                If b = 6 Then
                    BS = .ObrokPredSpavanjeBindingSource
                    DGV = .DataGridView14 'obrok pred spavanje
                End If

                For i = 0 To DGV.RowCount - 1
                    DGV.Rows.Remove(DGV.CurrentRow)
                Next i
                BS.AddNew()

            Next b

            'Naziv i priprema jela
            .TextBox11.Text = ""
            .TextBox64.Text = ""
            .TextBox65.Text = ""
            .TextBox66.Text = ""
            .TextBox67.Text = ""
            .TextBox68.Text = ""

            .ComboBox16.Text = 1   'broj konzumenata

            '    .ProgressBar8.Value = 0
            '   .ProgressBar9.Value = 0
            '  .ProgressBar10.Value = 0

            
        End With
    End Sub
End Module
