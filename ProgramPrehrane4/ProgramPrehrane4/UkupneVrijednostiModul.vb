Module UkupneVrijednostiModul
    Sub UkupneVrijednosti()
        On Error Resume Next
        With Form1

            Dim DGV As DataGridView

            Dim zit As Double = 0
            Dim pov As Double = 0
            Dim voc As Double = 0
            Dim mes As Double = 0
            Dim mlij As Double = 0
            Dim mas As Double = 0

            Dim OstaleNamirnice As Double = 0

            Dim Cijena As Double = 0

            Dim uglj As Double = 0
            Dim bjel As Double = 0
            Dim mast As Double = 0
            Dim energ As Double = 0

            Dim i As Integer
            .ListBox3.Items.Clear()   'serviranja
            .ListBox10.Items.Clear()   'nutrijenti
            .ListBox11.Items.Clear()   'ukupno nutrijenti - postoci
            .ListBox13.Items.Clear()    'ukupno serviranja - postoci
            .ListBox8.Items.Clear()      'ukupno - doručak

            'DORUČAK
            ' DGV = .DataGridView5
            Dim b As Integer
            For b = 1 To 6
                If b = 1 Then DGV = .DataGridView5 'dorucak
                If b = 2 Then DGV = .DataGridView9 'jutarnja uzina
                If b = 3 Then DGV = .DataGridView11 'rucak
                If b = 4 Then DGV = .DataGridView12 'popodnevna uzina
                If b = 5 Then DGV = .DataGridView13 'vecera
                If b = 6 Then DGV = .DataGridView14 'obrok pred spavanje

                For i = 0 To DGV.RowCount - 1

                    'serviranja
                    zit = zit + DGV.Rows(i).Cells(55).Value
                    pov = pov + DGV.Rows(i).Cells(56).Value
                    voc = voc + DGV.Rows(i).Cells(57).Value
                    mes = mes + DGV.Rows(i).Cells(58).Value
                    mlij = mlij + DGV.Rows(i).Cells(59).Value
                    mas = mas + DGV.Rows(i).Cells(60).Value

                    'ostale namirnice
                    If DGV.Rows(i).Cells(61).Value.ToString <> "" Then
                        If DGV.Rows(i).Cells(61).Value > 0 Then
                            OstaleNamirnice = OstaleNamirnice + DGV.Rows(i).Cells(11).Value
                        End If
                    End If

                    'cijena
                    If DGV.Rows(i).Cells(62).Value.ToString <> "" Then
                        If DGV.Rows(i).Cells(62).Value > 0 Then
                            Cijena = Cijena + DGV.Rows(i).Cells(62).Value
                        End If
                    End If

                    'nutrijenti
                    energ = energ + DGV.Rows(i).Cells(11).Value
                    uglj = uglj + DGV.Rows(i).Cells(13).Value
                    bjel = bjel + DGV.Rows(i).Cells(14).Value
                    mast = mast + DGV.Rows(i).Cells(15).Value

                Next i

            Next b

            'serviranja - ukupno
            .ListBox3.Items.Insert(0, Format(zit, "0.0"))
            .ListBox3.Items.Insert(1, Format(pov, "0.0"))
            .ListBox3.Items.Insert(2, Format(voc, "0.0"))
            .ListBox3.Items.Insert(3, Format(mes, "0.0"))
            .ListBox3.Items.Insert(4, Format(mlij, "0.0"))
            .ListBox3.Items.Insert(5, Format(mas, "0.0"))
            'serviranja - postoci
            .ListBox13.Items.Insert(0, Format((zit / .ListBox2.Items(0)) * 100, "0") & "%")
            .ListBox13.Items.Insert(1, Format((pov / .ListBox2.Items(1)) * 100, "0") & "%")
            .ListBox13.Items.Insert(2, Format((voc / .ListBox2.Items(2)) * 100, "0") & "%")
            .ListBox13.Items.Insert(3, Format((mes / .ListBox2.Items(3)) * 100, "0") & "%")
            .ListBox13.Items.Insert(4, Format((mlij / .ListBox2.Items(4)) * 100, "0") & "%")
            .ListBox13.Items.Insert(5, Format((mas / .ListBox2.Items(5)) * 100, "0") & "%")
            'serviranja - progress bar
            .ProgressBar1.Value = Format((zit / .ListBox2.Items(0)) * 100, "0")
            If Format((zit / .ListBox2.Items(0)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar1.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If
            .ProgressBar2.Value = Format((pov / .ListBox2.Items(1)) * 100, "0")
            If Format((pov / .ListBox2.Items(1)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar2.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If
            .ProgressBar3.Value = Format((voc / .ListBox2.Items(2)) * 100, "0")
            If Format((voc / .ListBox2.Items(2)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar3.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If
            .ProgressBar4.Value = Format((mes / .ListBox2.Items(3)) * 100, "0")
            If Format((mes / .ListBox2.Items(3)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar4.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If
            .ProgressBar5.Value = Format((mlij / .ListBox2.Items(4)) * 100, "0")
            If Format((mlij / .ListBox2.Items(4)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar5.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If
            .ProgressBar6.Value = Format((mas / .ListBox2.Items(5)) * 100, "0")
            If Format((mas / .ListBox2.Items(5)) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar6.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If

            'Ostale namirnice
            .Label282.Text = Format(OstaleNamirnice, "0") & " kcal"
            .Label285.Text = Format(OstaleNamirnice, "0")
            .ProgressBar11.Value = Format((OstaleNamirnice / .Label284.Text) * 100, "0")
            If Format((OstaleNamirnice / .Label284.Text) * 100, "0") > 100 Then
                .ProgressBar11.Value = 100
                .Label283.Text = ">"
            Else
                .Label283.Text = "OK"
            End If
            .Label281.Text = Format((.Label285.Text / .Label284.Text) * 100, "0") & "%"

            'cijena ukupno
            .Label318.Text = Cijena & " " & .ComboBox26.Text
            .Label361.Text = Cijena

            'nutrijenti - ukupno
            .ListBox10.Items.Insert(0, Format(energ, "0") & " kcal")
            .ListBox10.Items.Insert(1, "")   'prazno polje
            .ListBox10.Items.Insert(2, Format(uglj, "0") & " g")
            .ListBox10.Items.Insert(3, Format(bjel, "0") & " g")
            .ListBox10.Items.Insert(4, Format(mast, "0") & " g")
            'nutrijenti - postoci
            .ListBox11.Items.Insert(0, Format((energ / .TextBox3.Text) * 100, "0") & "%")
            .ListBox11.Items.Insert(1, "")   'prazno polje
            .ListBox11.Items.Insert(2, Format((uglj * 3.75 / energ) * 100, "0") & "%")
            .ListBox11.Items.Insert(3, Format((bjel * 4 / energ) * 100, "0") & "%")
            .ListBox11.Items.Insert(4, Format((mast * 9 / energ) * 100, "0") & "%")
         
            'energija - progress bar
            .ProgressBar7.Value = Format((energ / .TextBox3.Text) * 100, "0")
            If Format((energ / .TextBox3.Text) * 100, "0") > 100 Then
                ' .PictureBox11.Visible = True
                .ProgressBar7.Value = 100
            Else
                ' .PictureBox11.Visible = False
            End If

        
            'OK, > , < - SERVIRANJA
            .ListBox37.Items.Clear()
            Dim a As Integer
            For a = 0 To 5
                If .ListBox3.Items(a) < .ListBox2.Items(a) - .ListBox2.Items(a) * 0.05 Then .ListBox37.Items.Insert(a, "<")
                ' If .ListBox3.Items(a) = .ListBox2.Items(a) Then .ListBox37.Items.Insert(a, "OK")
                If .ListBox3.Items(a) >= .ListBox2.Items(a) - .ListBox2.Items(a) * 0.05 And .ListBox3.Items(a) <= .ListBox2.Items(a) + .ListBox2.Items(a) * 0.05 Then .ListBox37.Items.Insert(a, "OK")
                If .ListBox3.Items(a) > .ListBox2.Items(a) + .ListBox2.Items(a) * 0.05 Then .ListBox37.Items.Insert(a, ">")
            Next a

            'OK, > , < - ENERGIJA I NUTRIJENTI
            .ListBox38.Items.Clear()
            'Energija
            If energ >= .TextBox3.Text - 20 And energ <= .TextBox3.Text + 20 Then
                .ListBox38.Items.Insert(0, "OK")
            Else
                If energ < .TextBox3.Text - 20 Then .ListBox38.Items.Insert(0, "<")
                If energ > .TextBox3.Text + 20 Then .ListBox38.Items.Insert(0, ">")
            End If
            'Prazno polje
            .ListBox38.Items.Insert(1, "")
            'preporučene nutritivne vrijednosti
            'Ugljikohidrati
            If Format((uglj * 3.75 / energ) * 100, "0") >= Val(.Label88.Text) And Format((uglj * 3.75 / energ) * 100, "0") <= Val(.Label91.Text) Then
                .ListBox38.Items.Insert(2, "OK")
            Else
                If Format((uglj * 3.75 / energ) * 100, "0") < Val(.Label88.Text) Then .ListBox38.Items.Insert(2, "<")
                If Format((uglj * 3.75 / energ) * 100, "0") > Val(.Label91.Text) Then .ListBox38.Items.Insert(2, ">")
            End If
            'Bjelancevine
            If Format((bjel * 4 / energ) * 100, "0") >= Val(.Label89.Text) And Format((bjel * 4 / energ) * 100, "0") <= Val(.Label92.Text) Then
                .ListBox38.Items.Insert(3, "OK")
            Else
                If Format((bjel * 4 / energ) * 100, "0") < Val(.Label89.Text) Then .ListBox38.Items.Insert(3, "<")
                If Format((bjel * 4 / energ) * 100, "0") > Val(.Label92.Text) Then .ListBox38.Items.Insert(3, ">")
            End If
            'Masti
            If Format((mast * 9 / energ) * 100, "0") >= Val(.Label90.Text) And Format((mast * 9 / energ) * 100, "0") <= Val(.Label93.Text) Then
                .ListBox38.Items.Insert(4, "OK")
            Else
                If Format((mast * 9 / energ) * 100, "0") < Val(.Label90.Text) Then .ListBox38.Items.Insert(4, "<")
                If Format((mast * 9 / energ) * 100, "0") > Val(.Label93.Text) Then .ListBox38.Items.Insert(4, ">")
            End If

            .Label175.Text = Format(energ, "0")  'Energetska vrijednost jelovnika


            'graf - pita
            If .ListBox10.Items(0) = "0 kcal" Then
                .Chart2.Series(0).Points.Clear()
                .Chart2.Series(0).Points.AddY(20)
                .Chart2.Series(0).Points.AddY(30)
                .Chart2.Series(0).Points.AddY(50)
                Exit Sub
            End If

            .Chart2.Series(0).Points.Clear()
            .Chart2.Series(0).Points.AddY((uglj * 3.75 / energ) * 100)
            .Chart2.Series(0).Points.AddY((bjel * 4 / energ) * 100)
            .Chart2.Series(0).Points.AddY((mast * 9 / energ) * 100)
            'Parametri graf
            .Chart3.Series(0).Points.Clear()
            .Chart3.Series(0).Points.AddY((uglj * 3.75 / energ) * 100)
            .Chart3.Series(0).Points.AddY((bjel * 4 / energ) * 100)
            .Chart3.Series(0).Points.AddY((mast * 9 / energ) * 100)


            .Label14.Text = energ    'Odabrana energija jelovnika


        End With
    End Sub
End Module
