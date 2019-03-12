Module DgvStyleModul
    Sub DgvStyle()
        With Form1
            '.ColorDialog1.ShowDialog()
            Dim DGV As DataGridView = .DataGridView2
            ' Dim Boja As Color = Color.Lavender
            Dim Boja As Color = .ColorDialog1.Color
            Dim i As Integer = DGV.CurrentRow.Index


            'Pracenje antropometrijskih parametara
            '          DGV = .DataGridView10
            '         For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'baza klijenata
            '          DGV = .DataGridView6
            '         For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Tjalesna aktivnost detaljni izracun
            '           DGV = .DataGridView16
            '          For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Dodatna aktivnost
            '          DGV = .DataGridView3
            '         For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Vrsta dijete
            '           DGV = .DataGridView1
            '          For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Sve namirnice
            DGV = .DataGridView2
            For i = 0 To DGV.Rows.Count - 1 Step 2
                DGV.Rows(i).DefaultCellStyle.BackColor = Boja
                ' DGV.Rows(i).DefaultCellStyle.SelectionBackColor
            Next

            'Dorucak
            '          DGV = .DataGridView5
            '         For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Baza naziva jelovnika
            '         DGV = .DataGridView7
            '        For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            'Cijene
            '          DGV = .DataGridView15
            '         For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = Boja
            'Next

            My.Settings.DgvBoja = Boja
            My.Settings.Save()

        End With
    End Sub
End Module
