Module GotovaJelaModul
    Sub GotovaJela()
        ' On Error Resume Next
        With Form1
            Dim Namirnica As String = ""
            ' Dim Serviranje As Double = 1
            ' Dim Kolicina As Double = 1
            Dim Masa As Double = 0
            ' Dim S As Double = GotovaJelaForm.ComboBox1.Text    'Broj serviranja gotovog jela
            '  Dim K As Double = GotovaJelaForm.ComboBox1.Text   'Količina gotovog jela
            Dim K As Double = .TextBox12.Text   'Količina gotovog jela

            ' Dim Cijena As Double = GotovaJelaForm.TextBox1.Text   'Cijena jela
            Dim TermickaObrada As String = ""
            Dim DGV As DataGridView = .DataGridView2
            Dim Index As Integer = DGV.CurrentRow.Index
            Dim GotovoJelo As String = DGV.Rows(Index).Cells(1).Value.ToString
            .SveNamirniceBindingSource.RemoveFilter()
            Dim j As Integer
            Dim PripremaJela As String = ""

            Dim BS As BindingSource

            Dim NazivPripremaJela As TextBox = .TextBox11
            If .TabPage7.CanFocus = True Then
                NazivPripremaJela = .TextBox11
            End If

            If .TabPage8.CanFocus = True Then
                NazivPripremaJela = .TextBox64
            End If

            If .TabPage9.CanFocus = True Then
                NazivPripremaJela = .TextBox65
            End If

            If .TabPage10.CanFocus = True Then
                NazivPripremaJela = .TextBox66
            End If

            If .TabPage11.CanFocus = True Then
                NazivPripremaJela = .TextBox67
            End If

            If .TabPage12.CanFocus = True Then
                NazivPripremaJela = .TextBox68
            End If

            'LAZANJE
            ' If GotovoJelo = "Lazanje" Then Exit Sub

            'OMLET SA SIROM
            If GotovoJelo = "Omlet sa sirom" Then
                PripremaJela = "Na zagrijanom ulju ispržiti omlet. U sredinu gotovog omleta staviti sir i preklopiti."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If

                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Jaje, cijelo"
                        Masa = 47 * K
                        TermickaObrada = "omlet"
                    End If
                    If i = 2 Then
                        Namirnica = "Sir, Edam"
                        Masa = 28 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 3 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If

                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

              
            'LEŠO CVJETAČA
            If GotovoJelo = "Cvjetača na lešo" Then
                PripremaJela = "Skuhati očišćenu i opranu cvjetaču. Ocijediti. Začiniti uljem, isjeckanim češnjakom, paprom i soli."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If

                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Cvjetača, cvjetovi"
                        Masa = 130 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, maslinovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Češnjak, izrezani"
                        Masa = 0.5 * K
                        TermickaObrada = ""
                    End If

                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'SALATA OD KRUMPIRA
            If GotovoJelo = "Salata od krumpira" Then
                PripremaJela = "Skuhati krumpir, obuliti ga i izrezati na ploške. Očistiti i izrezati luk i češnjak. Ohlađeni krumpir pomiješati s lukom i češnjakom te posoliti i popapriti po želji."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If

                For i = 1 To 4
                    If i = 1 Then
                        Namirnica = "Krumpir, stari"
                        Masa = 135 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 2 Then
                        Namirnica = "Luk"
                        Masa = 12 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Ulje, maslinovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    If i = 4 Then
                        Namirnica = "Češnjak, izrezani"
                        Masa = 1 * K
                        TermickaObrada = ""
                    End If

                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'SALATA OD ZELENE SALATE
            If GotovoJelo = "Salata od zelene salate" Then
                PripremaJela = "Očistiti i oprati salatu. Listove usitniti na manje komade. Pomiješati ulje, ocat i sol te politi po salati. Promiješati."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If

                For i = 1 To 2
                    If i = 1 Then
                        Namirnica = "Salata, zelena, prosjek"
                        Masa = 60 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If

                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'SALATA OD RAJČICE
            If GotovoJelo = "Salata od rajčice" Then
                PripremaJela = "Oprati, očistiti i na kriške izrezati rajčicu. Pomiješati ulje, ocat i sol te politi po rajčici."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 2
                    If i = 1 Then
                        Namirnica = "Rajčica"
                        Masa = 180 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SALATA OD KRASTAVACA
            If GotovoJelo = "Salata od krastavaca" Then
                PripremaJela = "Na kriške izrezati oguljeni krastavac. Pomiješati ulje, ocat i sol te politi po krastavcu."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 2
                    If i = 1 Then
                        Namirnica = "Krastavac, oguljeni"
                        Masa = 200 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SALATA OD KUPUSA
            If GotovoJelo = "Salata od kupusa" Then
                PripremaJela = "Oprati, očistiti i izribati kupus. Pomiješati ulje, ocat i sol te politi po kupusu. Promiješati."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 2
                    If i = 1 Then
                        Namirnica = "Kupus, bijeli"
                        Masa = 140 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SENDVIC 1
            If GotovoJelo = "Sendvič (bijelo pecivo) sa sirom, šunkom, jajem, majonezom i kiselim krastavcima" Then
                PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 6
                    If i = 1 Then
                        Namirnica = "Pecivo, bijelo, mekano, srednje"
                        Masa = 100 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Šunka u ovitku"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Sir, Edam"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    If i = 4 Then
                        Namirnica = "Kiseli krastavci, ocijeđeni"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 5 Then
                        Namirnica = "Majoneza, klasična"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 6 Then
                        Namirnica = "Jaje, kuhano"
                        Masa = 25 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SENDVIC 2
            If GotovoJelo = "Sendvič (bijelo pecivo) sa sirom i šunkom" Then
                PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Pecivo, bijelo, mekano, srednje"
                        Masa = 100 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Šunka u ovitku"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Sir, Edam"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'PIRE KRUMPIR
            If GotovoJelo = "Pire krumpir" Then
                PripremaJela = "U ključalu vodu staviti oguljeni krumpir. Kada je kuhan, krumpir ocijediti, usitniti i dodati mlijeko, maslac i sol. Pritisnuti kroz posirku i dobro miješati dok se smjesa ne hmogenizira."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Krumpir, stari"
                        Masa = 200 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 2 Then
                        Namirnica = "Mlijeko, djelomično obrano, prosjek"
                        Masa = 20 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 3 Then
                        Namirnica = "Maslac"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'ISTARSKA MANESTRA
            If GotovoJelo = "Istarska maneštra" Then
                PripremaJela = "Preko noći u vodi namakati grah i slanutak. Na ulju blago popržiti crveni luk, dodati oprani grah i slanutak te na kockice izrezani suhi vrat, rajčicu, kukuruz i krumpir. Začiniti po želji. Doliti vodu i kuhati. Gotovo jelo posuti peršinom."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 8
                    If i = 1 Then
                        Namirnica = "Grah, suhi"
                        Masa = 42 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 2 Then
                        Namirnica = "Slanutak, suhi"
                        Masa = 50 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 3 Then
                        Namirnica = "Kukuruz, slatki, konzervirani, podgrijani, ocijeđeni"
                        Masa = 30 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 4 Then
                        Namirnica = "Krumpir, stari"
                        Masa = 45 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 5 Then
                        Namirnica = "Rajčica"
                        Masa = 30 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 6 Then
                        Namirnica = "Luk"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 7 Then
                        Namirnica = "Svinjetina, dimljena vratina"
                        Masa = 17 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 8 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 4.3 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'BUREK SA SIROM
            If GotovoJelo = "Burek sa sirom" Then
                PripremaJela = "Od brašna, masti, soli i vode napraviti kore. Kore pouljiti i filati posoljenim svježim sirom. Peći u pouljenoj tepsiji u zagrijanoj pećnici."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 5
                    If i = 1 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Brašno, pšenično, bijelo"
                        Masa = 138 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 3 Then
                        Namirnica = "Mast, svinjska"
                        Masa = 21 * K
                        TermickaObrada = ""
                    End If
                    If i = 4 Then
                        Namirnica = "Sir, svježi, klasični"
                        Masa = 115 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 5 Then
                        Namirnica = "Sol"
                        Masa = 4 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'PIZZA CUT
            If GotovoJelo = "Pizza, miješana, cut" Then
                PripremaJela = "Od brašna, kvasca, soli i vode napraviti tijesto. Ostavitit da se diže. Premijesiti ga, razvaljati i staviti u pouljenu tepsiju. Na njega posložiti izrezanu rajčicu, šunku, sir i gljive. Peći u zagrijanoj pećnici na 180*C."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 7
                    If i = 1 Then
                        Namirnica = "Brašno, pšenično, bijelo"
                        Masa = 95 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 2 Then
                        Namirnica = "Sir, Edam"
                        Masa = 30 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 3 Then
                        Namirnica = "Rajčica"
                        Masa = 50 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 4 Then
                        Namirnica = "Šunka u ovitku"
                        Masa = 30 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 5 Then
                        Namirnica = "Gljive, šampinjoni"
                        Masa = 30 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 6 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 12 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Kvasac, svježi"
                        Masa = 17 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'PIZZA, MIJEŠANA, VELIKA
            If GotovoJelo = "Pizza, miješana, velika" Then
                PripremaJela = "Od brašna, kvasca, soli i vode napraviti tijesto. Ostavitit da se diže. Premijesiti ga, razvaljati i staviti u pouljenu tepsiju. Na njega posložiti izrezanu rajčicu, šunku, sir i gljive. Peći u zagrijanoj pećnici na 180*C."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 7
                    If i = 1 Then
                        Namirnica = "Brašno, pšenično, bijelo"
                        Masa = 389 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 2 Then
                        Namirnica = "Sir, Edam"
                        Masa = 116 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 3 Then
                        Namirnica = "Rajčica"
                        Masa = 194 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 4 Then
                        Namirnica = "Šunka u ovitku"
                        Masa = 116 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 5 Then
                        Namirnica = "Gljive, šampinjoni"
                        Masa = 116 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 6 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 47 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Kvasac, svježi"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'HAMBURGER
            If GotovoJelo = "Hamburger s pecivom" Then
                PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Hamburger, pečen na roštilju"
                        Masa = 120 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Pecivo za hamburger"
                        Masa = 120 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'PLJESKAVICA PEČENA NA ROŠTILJU
            If GotovoJelo = "Pljeskavica" Then
                PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 2
                    If i = 1 Then
                        Namirnica = "Hamburger, pečen na roštilju"
                        Masa = 120 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SALATA OD LIGANJA
            If GotovoJelo = "Salata od liganja i krumpira" Then
                PripremaJela = "Kuhane i izrezane lignje pomiješati s kuhanim i na ploške izrezanim krumpirom, rajčicom i lukom. Pomiješati ocat, maslinovo ulje, limunov sok, sol i češnjak te politi po lignjama. Posuti izrezanim peršinom."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 10
                    If i = 1 Then
                        Namirnica = "Lignja"
                        Masa = 70 * K
                        TermickaObrada = "pirjanje u vodi"   'poširanje, provjeriti
                    End If
                    If i = 2 Then
                        Namirnica = "Ocat"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Ulje, maslinovo"
                        Masa = 40 * K
                        TermickaObrada = ""
                    End If
                    If i = 4 Then
                        Namirnica = "Krumpir, mladi, kuhan u neposoljenoj vodi"
                        Masa = 50 * K
                        TermickaObrada = ""
                    End If
                    If i = 5 Then
                        Namirnica = "Rajčica"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 6 Then
                        Namirnica = "Luk"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Češnjak, izrezani"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 8 Then
                        Namirnica = "Peršin, list, usitnjeni"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 9 Then
                        Namirnica = "Limunov sok"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 10 Then
                        Namirnica = "Sol"
                        Masa = 1 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'GOVEĐI GULAŠ S KRUMPIROM
            If GotovoJelo = "Gulaš, goveđi, s krumpirom" Then
                PripremaJela = "Na ulju popržiti luk. Popržiti na kocke izrezano meso. Dodati izrezanu rajčicu, mrkvu, češnjak i krumpir i popržiti par minuta. Začiniti. Dodati vodu, smanjiti vatru i kuhati dok se meso ne skuha. Pred kraj dodati vino i posuti peršinom."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 9
                    If i = 1 Then
                        Namirnica = "Govedina, lopatica, s masti"
                        Masa = 100 * K
                        TermickaObrada = "mesna jela"   'poširanje, provjeriti
                    End If
                    If i = 2 Then
                        Namirnica = "Krumpir, stari"
                        Masa = 170 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 3 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 4 Then
                        Namirnica = "Luk"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 5 Then
                        Namirnica = "Rajčica"
                        Masa = 15 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 6 Then
                        Namirnica = "Češnjak"
                        Masa = 2 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Vino, bijelo, prosjek"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 8 Then
                        Namirnica = "Mrkva"
                        Masa = 10 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 9 Then
                        Namirnica = "Sol"
                        Masa = 2 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            'SARMA   ' greška
            If GotovoJelo = "Sarma" Then
                PripremaJela = "Na ulju popržen luk pomiješati s mljevenim mesom, rižom i začinima. Dobro isprati kiseli kupus. Smjesu stavljati na listove i zarolati. Sarme poredati u posudu i prekriti izrezanim kupusom. Prekriti vodom i kuhati na umjerenoj vatri."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 5
                    If i = 1 Then
                        Namirnica = "Kupus, bijeli"
                        Masa = 450 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Govedina, mljevena"    'provjeriti
                        Masa = 100 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 3 Then
                        Namirnica = "Riža, bijela"   'provjeriti
                        Masa = 10 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 4 Then
                        Namirnica = "Luk"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 5 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'PUNJENA PAPRIKA
            If GotovoJelo = "Punjena paprika" Then
                PripremaJela = "Na ulju popržen luk pomiješati s mljevenim mesom, rižom i začinima. Smjesom puniti oprane i očišćene paprike. Na vrhu svake paprike staviti komadić rajčice. Paprike poredati u posudu. Prekriti vodom i kuhati na umjerenoj vatri."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 7
                    If i = 1 Then
                        Namirnica = "Govedina, mljevena"
                        Masa = 117 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 2 Then
                        Namirnica = "Paprika, zelena, izrezana"
                        Masa = 150 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 3 Then
                        Namirnica = "Luk"
                        Masa = 50 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 4 Then
                        Namirnica = "Riža, bijela"
                        Masa = 20 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 5 Then
                        Namirnica = "Rajčica"
                        Masa = 60 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 6 Then
                        Namirnica = "Sol"
                        Masa = 1 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'ŠPAGETI BOLONJEZ
            If GotovoJelo = "Špageti bolonjez" Then
                PripremaJela = "Na ulju popržiti luk. Dodati mljeveno meso i kratko popršiti. Dodati izrezanu rajčicu i mrkvu. Začiniti. Smanjiti vatru. Uz dolijevanje vode miješati i pirjati dok se meso ne skuha. Kuhanu tjesteninu preliti umakom i posuti parmezanom."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 7
                    If i = 1 Then
                        Namirnica = "Mrkva, stara"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 4.3 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Govedina, mljevena"
                        Masa = 100 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 4 Then
                        Namirnica = "Rajčica"
                        Masa = 25 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 5 Then
                        Namirnica = "Tjestenina, špageti, bijela"
                        Masa = 100 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 6 Then
                        Namirnica = "Sir, Parmezan"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Luk"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'ZAPEČENA TJESTENINA S MESOM
            If GotovoJelo = "Zapečena tjestenina s mesom" Then
                PripremaJela = "Na ulju popržiti luk. Dodati meso i kratko popržiti. Dodati izrezanu rajčicu i mrkvu. Začiniti. Na dno tepsije staviti tjesteninu, pa meso i ponovo tjesteninu. Peći u zagrijanoj pećnici. Pred kraj preliti umućenim jajem i vrhnjem. Zapeći."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 9
                    If i = 1 Then
                        Namirnica = "Jaje, cijelo"
                        Masa = 25 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 2 Then
                        Namirnica = "Vrhnje za kuhanje, s 19% m.m."
                        Masa = 50 * K
                        TermickaObrada = "pečenje"
                    End If
                    If i = 3 Then
                        Namirnica = "Mrkva, stara"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 4 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 4.3 * K
                        TermickaObrada = ""
                    End If
                    If i = 5 Then
                        Namirnica = "Govedina, mljevena"
                        Masa = 100 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 6 Then
                        Namirnica = "Rajčica"
                        Masa = 25 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 7 Then
                        Namirnica = "Tjestenina, makaroni"
                        Masa = 100 * K
                        TermickaObrada = "kuhanje"
                    End If
                    If i = 8 Then
                        Namirnica = "Sir, Parmezan"
                        Masa = 20 * K
                        TermickaObrada = ""
                    End If
                    If i = 9 Then
                        Namirnica = "Luk"
                        Masa = 20 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'SVINJSKO PEČENJE
            If GotovoJelo = "Svinjsko pečenje" Then
                PripremaJela = "U vatrostalnoj posudi staviti meso, ulje te očišćenu i na krupne komade izrezanu mrkvu i luk. Začiniti i poškropiti vodom. Poklopiti i peći u zagrijanoj pećnici. Pred kraj pečenja otklopiti da meso porumeni."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 4
                    If i = 1 Then
                        Namirnica = "Svinjetina, odrezak, s masti"
                        Masa = 115 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 2 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 8.6 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Luk"
                        Masa = 10 * K
                        TermickaObrada = "prženje"
                    End If
                    If i = 4 Then
                        Namirnica = "Mrkva, stara"
                        Masa = 10 * K
                        TermickaObrada = "prženje"
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

            
            'HOT DOG
            If GotovoJelo = "Hot-dog (pecivo, hrenovka i senf)" Then
                PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 3
                    If i = 1 Then
                        Namirnica = "Pecivo, bijelo, mekano, srednje"
                        Masa = 100 * K
                        TermickaObrada = ""
                    End If
                    If i = 2 Then
                        Namirnica = "Hrenovka"
                        Masa = 70 * K
                        TermickaObrada = ""
                    End If
                    If i = 3 Then
                        Namirnica = "Senf"
                        Masa = 8 * K
                        TermickaObrada = ""
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If


            'PAŠTICADA
            If GotovoJelo = "Pašticada" Then
                PripremaJela = "Na ulju popržiti luk. Meso premazano senfom i slaninu kratko popržiti. Dodati izrezano povrće, šljive i vino. Začiniti. Smanjiti vatru. Uz dolijevanje vode pirjati dok se meso ne skuha. Umak propasirati i njime preliti meso."
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
                For i = 1 To 13
                    If i = 1 Then
                        Namirnica = "Govedina, but"
                        Masa = 115 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 2 Then
                        Namirnica = "Kiseli krastavci, ocijeđeni"
                        Masa = 5 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 3 Then
                        Namirnica = "Rajčica"
                        Masa = 30 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 4 Then
                        Namirnica = "Senf"
                        Masa = 2 * K
                        TermickaObrada = ""
                    End If
                    If i = 5 Then
                        Namirnica = "Ulje, maslinovo"
                        Masa = 5 * K
                        TermickaObrada = ""
                    End If
                    If i = 6 Then
                        Namirnica = "Ulje, suncokretovo"
                        Masa = 10 * K
                        TermickaObrada = ""
                    End If
                    If i = 7 Then
                        Namirnica = "Luk"
                        Masa = 50 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 8 Then
                        Namirnica = "Češnjak, izrezani"
                        Masa = 2 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 9 Then
                        Namirnica = "Mrkva, stara"
                        Masa = 25 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    If i = 10 Then
                        Namirnica = "Sol"
                        Masa = 2 * K
                        TermickaObrada = ""
                    End If
                    If i = 11 Then
                        Namirnica = "Slanina, samo mast"
                        Masa = 10 * K
                        TermickaObrada = "mesna jela"
                    End If
                    If i = 12 Then
                        Namirnica = "Šljiva, suha"
                        Masa = 5 * K
                        TermickaObrada = "pirjanje"
                    End If
                    If i = 13 Then
                        Namirnica = "Vino, crno, prosjek"
                        Masa = 15 * K
                        TermickaObrada = "povrtna jela"
                    End If
                    For j = 0 To DGV.RowCount - 1
                        If DGV.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV.Rows(j).Cells(1).Value = Namirnica Then
                                DGV.CurrentCell = DGV.Rows(j).Cells(1)
                                OdaberiNamirnicu()
                                GubiciVitamina()
                                .ComboBox11.Text = TermickaObrada   'Termicka obrada
                                .RadioButton14.Checked = True   'Kolicina Radio button
                                .TextBox6.Text = Masa   'Masa
                                .RadioButton13.Checked = True   'Serviranje Radio button
                                .TextBox81.Text = 1000   'Kolicina g
                                PrebaciNamirnicu()
                                UkupneVrijednosti()
                                ObrociNutrijentiUkupno()
                            End If
                        End If
                    Next j
                Next i
            End If

                DGV.CurrentRow.Selected = False
                '  OdaberiNamirnicu()
            .TextBox85.Text = ""     'Serviranja
                .TextBox5.Text = ""   'Naziv Namirnice
                .TextBox12.Text = ""   'Kolicina
                .Label229.Text = ""   'Kolicina
                .Label103.Text = ""   'Mjera
                .Label230.Text = ""   'Mjera
                .TextBox6.Text = ""   'Masa_g
                .Label228.Text = ""   'Masa_g
                .ListBox83.Items.Clear()   'gubici vitamina
            .ComboBox11.Text = "Termička obrada"   'termicka obrada
                .TextBox77.Text = ""  'cijena namirnice
            .RadioButton14.Checked = True   'Masa

     
        End With
    End Sub
End Module
