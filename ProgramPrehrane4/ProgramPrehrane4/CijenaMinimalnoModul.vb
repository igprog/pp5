Module CijenaMinimalnoModul
    Sub CijenaMinimalno()
        With Form1
            .Label338.Text = ""
            Dim DGV As DataGridView = .DataGridView5
          
            For c = 1 To 6
                If c = 1 Then DGV = .DataGridView5 'dorucak
                If c = 2 Then DGV = .DataGridView9 'jutarnja uzina
                If c = 3 Then DGV = .DataGridView11 'rucak
                If c = 4 Then DGV = .DataGridView12 'popodnevna uzina
                If c = 5 Then DGV = .DataGridView13 'vecera
                If c = 6 Then DGV = .DataGridView14 'obrok pred spavanje

                For f = 0 To DGV.RowCount - 1
                    If DGV.Rows(f).Cells(7).Value IsNot DBNull.Value Then
                        If (DGV.Rows(f).Cells(7).Value) > 0 And Convert.ToString(DGV.Rows(f).Cells(62).Value) = "" Then
                           .Label338.Text = ">"
                             Exit Sub
                        End If
                    End If
                Next f
            Next c
        End With
    End Sub
End Module
