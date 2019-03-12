Module EditiranjeNamirniceModul
    Sub EditiranjeNamirnica()
        ' On Error Resume Next
        With EditiranjeNamirniceForma
            Dim DGV As DataGridView = Form1.DataGridView5
            '   Dim StaroServiranje As Label = .Label8   
           
            If Form1.TabPage7.CanFocus = True Then
                DGV = Form1.DataGridView5
            End If
            If Form1.TabPage8.CanFocus = True Then
                DGV = Form1.DataGridView9
            End If
            If Form1.TabPage9.CanFocus = True Then
                DGV = Form1.DataGridView11
            End If
            If Form1.TabPage10.CanFocus = True Then
                DGV = Form1.DataGridView12
            End If
            If Form1.TabPage11.CanFocus = True Then
                DGV = Form1.DataGridView13
            End If
            If Form1.TabPage12.CanFocus = True Then
                DGV = Form1.DataGridView14
            End If

            Dim NovoServiranje As Label = .Label4
            Dim NovaMasa As Label = .Label6
            Dim NovaKolicina As Label = .Label7
            Dim NovaMjera As Label = .Label14

            Dim i As Integer = DGV.CurrentRow.Index
            Dim Koeficjent As Double = NovaMasa.Text / DGV.Rows(i).Cells(7).Value

            If .RadioButton1.Checked = True Then
                Koeficjent = NovoServiranje.Text / DGV.Rows(i).Cells(7).Value
            End If
            If .RadioButton2.Checked = True Then
                Koeficjent = NovaMasa.Text / DGV.Rows(i).Cells(10).Value
            End If
            If .RadioButton3.Checked = True Then
                Koeficjent = NovaKolicina.Text / DGV.Rows(i).Cells(8).Value
            End If


            ' DGV.Rows(i).Cells(7).Value = Format((DGV.Rows(i).Cells(7).Value * Koeficjent), "0.000")   'serviranje
            ' DGV.Rows(i).Cells(8).Value = Format((DGV.Rows(i).Cells(8).Value * Koeficjent), "0.000")  'količina
            ' DGV.Rows(i).Cells(10).Value = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0.000")   'masa

            DGV.Rows(i).Cells(7).Value = (DGV.Rows(i).Cells(7).Value * Koeficjent)   'serviranje
            DGV.Rows(i).Cells(8).Value = (DGV.Rows(i).Cells(8).Value * Koeficjent) 'količina
            DGV.Rows(i).Cells(9).Value = NovaMjera.Text   'novamjera
            DGV.Rows(i).Cells(10).Value = (DGV.Rows(i).Cells(10).Value * Koeficjent)  'masa


            '  NovaKolicina.Text = Format((DGV.Rows(i).Cells(8).Value), "0.00")' & " " & DGV.Rows(i).Cells(9).Value 'količina i mjera
            '    If DGV.Rows(i).Cells(10).Value < 10 And DGV.Rows(i).Cells(10).Value >= 1 Then
            'DGV.Rows(i).Cells(10).Value = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0.0")   'masa
            '       End If
            '      If DGV.Rows(i).Cells(10).Value < 1 Then
            'DGV.Rows(i).Cells(10).Value = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0.00")   'masa
            '     End If
            '    If DGV.Rows(i).Cells(10).Value >= 10 Then
            'DGV.Rows(i).Cells(10).Value = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0")   'masa
            '    End If
            '  NovaMasa.Text = Format((DGV.Rows(i).Cells(10).Value), "0.00") '& " g"  'masa

            NovoServiranje.Text = DGV.Rows(i).Cells(7).Value  'serviranje

            'NovaKolicina.Text = DGV.Rows(i).Cells(8).Value & " " & DGV.Rows(i).Cells(9).Value 'količina i mjera
            NovaKolicina.Text = DGV.Rows(i).Cells(8).Value 'količina 

            NovaMjera.Text = DGV.Rows(i).Cells(9).Value.ToString 'mjera 

            'NovaMasa.Text = DGV.Rows(i).Cells(10).Value & " g"  'masa
            NovaMasa.Text = DGV.Rows(i).Cells(10).Value  'masa


            Dim j As Integer = 11
            For j = 11 To 62
                If DGV.Rows(i).Cells(j).Value IsNot DBNull.Value Then
                    If DGV.Rows(i).Cells(j).Value.ToString <> "N" Then
                        ' DGV.Rows(i).Cells(j).Value = Format((DGV.Rows(i).Cells(j).Value * Koeficjent), "0.000")  'Energija kcal...
                        DGV.Rows(i).Cells(j).Value = DGV.Rows(i).Cells(j).Value * Koeficjent  'Energija kcal...

                    End If
                End If
            Next


        End With
    End Sub
End Module
