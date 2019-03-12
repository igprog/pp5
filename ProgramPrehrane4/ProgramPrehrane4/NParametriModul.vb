Module NParametriModul
    Sub NParametri()
        On Error Resume Next
        With Form1
            .ListBox40.Items.Clear()
            Dim a As Integer
            For a = 0 To 38
                .ListBox40.Items.Insert(a, "")
            Next a

            .ListBox41.Items.Clear()
            .ListBox42.Items.Clear()
            .ListBox43.Items.Clear()
            .ListBox44.Items.Clear()
            .ListBox47.Items.Clear()


            Dim DGV As DataGridView
            '  DGV = .DataGridView5   'Dorucak
            Dim b As Integer
            For b = 1 To 6
                If b = 1 Then DGV = .DataGridView5 'dorucak
                If b = 2 Then DGV = .DataGridView9 'jutarnja uzina
                If b = 3 Then DGV = .DataGridView11 'rucak
                If b = 4 Then DGV = .DataGridView12 'popodnevna uzina
                If b = 5 Then DGV = .DataGridView13 'vecera
                If b = 6 Then DGV = .DataGridView14 'obrok pred spavanje

                Dim j As Integer    'Kolona
                For j = 16 To 54
                    Dim i As Integer    'Red
                    For i = 0 To DGV.RowCount - 1
                        If Convert.ToString(DGV.Rows(i).Cells(j).Value) = "N" Then
                            .ListBox40.Items.Insert(j - 16, "min")
                        End If
                    Next i
                Next j

            Next b

            'Zasicene kiseline, Trans kiseline, kolesterol
            .ListBox41.Items.Insert(0, .ListBox40.Items(8))
            .ListBox41.Items.Insert(1, .ListBox40.Items(11))
            .ListBox41.Items.Insert(2, .ListBox40.Items(12))

            'natrij, kalij, klor
            .ListBox42.Items.Insert(0, .ListBox40.Items(13))
            .ListBox42.Items.Insert(1, .ListBox40.Items(14))
            .ListBox42.Items.Insert(2, .ListBox40.Items(21))

            'karoten
            .ListBox43.Items.Insert(0, .ListBox40.Items(26))

            'ostali parametri
            .ListBox44.Items.Insert(0, .ListBox40.Items(7)) 'Vlakna
            .ListBox44.Items.Insert(1, .ListBox40.Items(9))
            .ListBox44.Items.Insert(2, .ListBox40.Items(10))
            .ListBox44.Items.Insert(3, .ListBox40.Items(15))
            .ListBox44.Items.Insert(4, .ListBox40.Items(16))
            .ListBox44.Items.Insert(5, .ListBox40.Items(17))
            .ListBox44.Items.Insert(6, .ListBox40.Items(18))
            .ListBox44.Items.Insert(7, .ListBox40.Items(19))
            .ListBox44.Items.Insert(8, .ListBox40.Items(20))
            .ListBox44.Items.Insert(9, .ListBox40.Items(22))
            .ListBox44.Items.Insert(10, .ListBox40.Items(23))
            .ListBox44.Items.Insert(11, .ListBox40.Items(24))
            .ListBox44.Items.Insert(12, .ListBox40.Items(25))
            .ListBox44.Items.Insert(13, .ListBox40.Items(27))
            .ListBox44.Items.Insert(14, .ListBox40.Items(28))
            .ListBox44.Items.Insert(15, .ListBox40.Items(29))
            .ListBox44.Items.Insert(16, .ListBox40.Items(30))
            .ListBox44.Items.Insert(17, .ListBox40.Items(31))
            .ListBox44.Items.Insert(18, .ListBox40.Items(32))
            .ListBox44.Items.Insert(19, .ListBox40.Items(33))
            .ListBox44.Items.Insert(20, .ListBox40.Items(34))
            .ListBox44.Items.Insert(21, .ListBox40.Items(35))
            .ListBox44.Items.Insert(22, .ListBox40.Items(36))
            .ListBox44.Items.Insert(23, .ListBox40.Items(37))
            .ListBox44.Items.Insert(24, .ListBox40.Items(38))

            'skrob - laktoza
            .ListBox47.Items.Insert(0, .ListBox40.Items(0))
            .ListBox47.Items.Insert(1, .ListBox40.Items(1))
            .ListBox47.Items.Insert(2, .ListBox40.Items(2))
            .ListBox47.Items.Insert(3, .ListBox40.Items(3))
            .ListBox47.Items.Insert(4, .ListBox40.Items(4))
            .ListBox47.Items.Insert(5, .ListBox40.Items(5))
            .ListBox47.Items.Insert(6, .ListBox40.Items(6))


        End With
    End Sub
End Module
