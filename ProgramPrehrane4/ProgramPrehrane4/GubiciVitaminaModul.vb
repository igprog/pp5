Module GubiciVitaminaModul
    Sub GubiciVitamina()
        On Error Resume Next
        With Form1
            Dim i As Integer = .DataGridView2.CurrentRow.Index
            Dim SkupinaNamirnica As String = .DataGridView2.Rows(i).Cells(56).Value  'zadnja kolona DGV2 GubiciVitamina, zitarice, mlijeko, jaja, meso, riba, povrce, voce
            Dim Gubici_LB As ListBox = .ListBox83     'gubici, postoci
            Dim TermickaObrada_CB As ComboBox = .ComboBox11   'termicka obrada

            Gubici_LB.Items.Add(0)    'Karoten
            Gubici_LB.Items.Add(0)    'Vitamin D
            Gubici_LB.Items.Add(0)    'Vitamin E
            Gubici_LB.Items.Add(0)    'Vitamin B1
            Gubici_LB.Items.Add(0)    'Vitamin B2
            Gubici_LB.Items.Add(0)    'Vitamin B3
            Gubici_LB.Items.Add(0)    'Vitamin B6
            Gubici_LB.Items.Add(0)    'Vitamin B12
            Gubici_LB.Items.Add(0)    'Folat
            Gubici_LB.Items.Add(0)    'PantotenskaKiselina
            Gubici_LB.Items.Add(0)    'Biotin
            Gubici_LB.Items.Add(0)    'Vitamin C
            Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema

            If SkupinaNamirnica = "zitarice" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "kuhanje" Then    'kuhanje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(40)    'Vitamin B1
                    Gubici_LB.Items.Add(40)    'Vitamin B2
                    Gubici_LB.Items.Add(40)    'Vitamin B3
                    Gubici_LB.Items.Add(40)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(40)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(40)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "pečenje" Then    'pecenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(25)    'Vitamin B1
                    Gubici_LB.Items.Add(15)    'Vitamin B2
                    Gubici_LB.Items.Add(5)    'Vitamin B3
                    Gubici_LB.Items.Add(25)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(25)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "povrce" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "kuhanje" Then    'kuhanje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(35)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(30)    'Vitamin B3
                    Gubici_LB.Items.Add(40)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(40)    'Folat
                    Gubici_LB.Items.Add(0)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(45)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "prženje" Then    'przenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(0)    'Vitamin B2
                    Gubici_LB.Items.Add(0)    'Vitamin B3
                    Gubici_LB.Items.Add(25)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(55)    'Folat
                    Gubici_LB.Items.Add(0)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(30)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 2 Or TermickaObrada_CB.Text = "povrtna jela" Then    'povrtna jela
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(20)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(20)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "voce" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "pirjanje" Then    'pirjanje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(25)    'Vitamin B1
                    Gubici_LB.Items.Add(25)    'Vitamin B2
                    Gubici_LB.Items.Add(25)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(80)    'Folat
                    Gubici_LB.Items.Add(25)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(25)    'Biotin
                    Gubici_LB.Items.Add(25)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "meso" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "roštilj/prženje" Then    'roštilj (Roštilj/prženje)
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(20)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(20)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(20)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(10)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "mesna jela" Then    'Mesna jela
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(20)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(20)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(20)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(10)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "riba" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "pirjanje u vodi" Then    'Pirjanje u vodi
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(10)    'Vitamin B1
                    Gubici_LB.Items.Add(0)    'Vitamin B2
                    Gubici_LB.Items.Add(10)    'Vitamin B3
                    Gubici_LB.Items.Add(0)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(0)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(10)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "pečenje" Then    'pecenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(30)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(20)    'Vitamin B3
                    Gubici_LB.Items.Add(10)    'Vitamin B6
                    Gubici_LB.Items.Add(10)    'Vitamin B12
                    Gubici_LB.Items.Add(20)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(10)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 2 Or TermickaObrada_CB.Text = "roštilj" Then    'roštilj
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(10)    'Vitamin B1
                    Gubici_LB.Items.Add(10)    'Vitamin B2
                    Gubici_LB.Items.Add(10)    'Vitamin B3
                    Gubici_LB.Items.Add(10)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(0)    'Folat
                    Gubici_LB.Items.Add(5)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 3 Or TermickaObrada_CB.Text = "prženje" Then    'przenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(20)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(0)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(10)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "jaja" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "kajgana/omlet" Then    'umucena (kajgana/omlet)
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(5)    'Vitamin B1
                    Gubici_LB.Items.Add(20)    'Vitamin B2
                    Gubici_LB.Items.Add(5)    'Vitamin B3
                    Gubici_LB.Items.Add(15)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(30)    'Folat
                    Gubici_LB.Items.Add(15)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "pečenje" Then    'pecenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(0)    'Vitamin E
                    Gubici_LB.Items.Add(15)    'Vitamin B1
                    Gubici_LB.Items.Add(15)    'Vitamin B2
                    Gubici_LB.Items.Add(5)    'Vitamin B3
                    Gubici_LB.Items.Add(25)    'Vitamin B6
                    Gubici_LB.Items.Add(0)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(25)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(0)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            If SkupinaNamirnica = "mlijeko" Then
                If TermickaObrada_CB.Text = "" Or TermickaObrada_CB.Text = "Termička obrada" Then Exit Sub
                If TermickaObrada_CB.SelectedIndex = 0 Or TermickaObrada_CB.Text = "kuhanje" Then    'kuhanje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(20)    'Vitamin E
                    Gubici_LB.Items.Add(10)    'Vitamin B1
                    Gubici_LB.Items.Add(10)    'Vitamin B2
                    Gubici_LB.Items.Add(0)    'Vitamin B3
                    Gubici_LB.Items.Add(10)    'Vitamin B6
                    Gubici_LB.Items.Add(5)    'Vitamin B12
                    Gubici_LB.Items.Add(20)    'Folat
                    Gubici_LB.Items.Add(10)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 1 Or TermickaObrada_CB.Text = "umaci" Then    'umaci
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(20)    'Vitamin E
                    Gubici_LB.Items.Add(20)    'Vitamin B1
                    Gubici_LB.Items.Add(10)    'Vitamin B2
                    Gubici_LB.Items.Add(0)    'Vitamin B3
                    Gubici_LB.Items.Add(20)    'Vitamin B6
                    Gubici_LB.Items.Add(5)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(20)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
                If TermickaObrada_CB.SelectedIndex = 2 Or TermickaObrada_CB.Text = "pečenje" Then    'pecenje
                    Gubici_LB.Items.Clear()   'brisi list box
                    Gubici_LB.Items.Add(0)    'Karoten
                    Gubici_LB.Items.Add(0)    'Vitamin D
                    Gubici_LB.Items.Add(20)    'Vitamin E
                    Gubici_LB.Items.Add(25)    'Vitamin B1
                    Gubici_LB.Items.Add(15)    'Vitamin B2
                    Gubici_LB.Items.Add(5)    'Vitamin B3
                    Gubici_LB.Items.Add(25)    'Vitamin B6
                    Gubici_LB.Items.Add(5)    'Vitamin B12
                    Gubici_LB.Items.Add(50)    'Folat
                    Gubici_LB.Items.Add(25)    'PantotenskaKiselina
                    Gubici_LB.Items.Add(0)    'Biotin
                    Gubici_LB.Items.Add(50)    'Vitamin C
                    Gubici_LB.Items.Add(0)    'Vitamin K   'njega nema
                End If
            End If

            TermickaObrada_CB.Items.Clear()    'brisi BomboBox11

        End With
    End Sub
End Module
