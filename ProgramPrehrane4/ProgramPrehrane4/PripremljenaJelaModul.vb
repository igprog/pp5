Module PripremljenaJelaModul
    Sub PripremljenaJela()
        PripremljenaJelaForm.Show()

        With PripremljenaJelaForm

            Dim DGV As DataGridView = .DataGridView1
            Dim DGV1 As DataGridView = Form1.DataGridView2
            Dim Index As Integer = DGV1.CurrentRow.Index
            Dim GotovoJelo As String = DGV1.Rows(Index).Cells(1).Value.ToString
            Dim K As Double = Form1.TextBox12.Text   'Količina gotovog jela

            .PripremljenaJelaBindingSource.RemoveFilter()
            .PripremljenaJelaBindingSource.Filter = "NazivJela='" & Form1.TextBox5.Text & "'"

            'If DGV.Rows(0).Cells(3).Value Is DBNull.Value Then DGV.Rows(0).Cells(3).Value = "" 'Ako nema pripreme jela

            Dim Namirnica As String = ""
            Dim Masa As Double = 0
            Dim TermickaObrada As String = ""
            Dim PripremaJela As String = DGV.Rows(0).Cells(3).Value.ToString
            Dim i As Integer = 0
            Dim j As Integer = 0

            With Form1
                .SveNamirniceBindingSource.RemoveFilter()

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

                If PripremaJela = "N" Then PripremaJela = ""
                If NazivPripremaJela.Text = "" Then
                    NazivPripremaJela.Text = GotovoJelo & vbCrLf & PripremaJela
                Else
                    NazivPripremaJela.Text = NazivPripremaJela.Text & vbCrLf & GotovoJelo & vbCrLf & PripremaJela
                End If
            End With

            For i = 0 To DGV.RowCount - 1
                If DGV.Rows(i).Cells(4).Value = "" Then
                    PripremljenaJelaForm.Close()
                    Exit Sub
                End If

                Namirnica = DGV.Rows(i).Cells(4).Value
                Masa = DGV.Rows(i).Cells(5).Value * K
                TermickaObrada = DGV.Rows(i).Cells(6).Value
                PripremaJela = DGV.Rows(i).Cells(3).Value


                With Form1

                    For j = 0 To DGV1.RowCount - 1
                        If DGV1.Rows(j).Cells(1).Value IsNot DBNull.Value Then
                            If DGV1.Rows(j).Cells(1).Value = Namirnica Then
                                DGV1.CurrentCell = DGV1.Rows(j).Cells(1)
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
                    Next
                End With

            Next

        End With
        PripremljenaJelaForm.Close()
    End Sub
End Module
