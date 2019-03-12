Public Class EditiranjeNamirniceForma

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        '     On Error Resume Next
        '      EditiranjeNamirnica()
        '     UkupneVrijednosti()
        '    ObrociNutrijentiUkupno()
   


        Me.Close()


        '      Namirnica = DGV.Rows(i).Cells(5).Value
        '      Serviranje = DGV.Rows(i).Cells(7).Value
        '        Label3.Text = DGV.Rows(i).Cells(7).Value
    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'On Error Resume Next
        With Me
            Dim DGV As DataGridView = Form1.DataGridView5

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

            Dim StaroServiranje As Label = .Label8
            'Dim NovoServiranje As Label = .Label4
            Dim StaraMasa As Label = .Label3
            Dim StaraKolicina As Label = .Label5

            Dim i As Integer = DGV.CurrentRow.Index

            '    Dim Koeficjent As Double = NovoServiranje.Text / DGV.Rows(i).Cells(7).Value
            Dim Koeficjent As Double = StaroServiranje.Text / DGV.Rows(i).Cells(7).Value
            If .RadioButton1.Checked = True Then
                Koeficjent = StaroServiranje.Text / DGV.Rows(i).Cells(7).Value
            End If
            If .RadioButton2.Checked = True Then
                Koeficjent = StaraMasa.Text / DGV.Rows(i).Cells(10).Value
            End If
            If .RadioButton3.Checked = True Then
                Koeficjent = StaraKolicina.Text / DGV.Rows(i).Cells(8).Value
            End If

        
            DGV.Rows(i).Cells(7).Value = Format((DGV.Rows(i).Cells(7).Value * Koeficjent), "0.000")   'serviranje
            DGV.Rows(i).Cells(8).Value = Format((DGV.Rows(i).Cells(8).Value * Koeficjent), "0.000")  'količina
            '      .Label6.Text = Format((DGV.Rows(i).Cells(8).Value * Koeficjent), "0.0") & DGV.Rows(i).Cells(9).Value 'količina i mjera
            DGV.Rows(i).Cells(10).Value = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0.000")   'masa
            '     .Label7.Text = Format((DGV.Rows(i).Cells(10).Value * Koeficjent), "0.0") & " g"  'masa

            Dim j As Integer = 11
            For j = 11 To 62
                If DGV.Rows(i).Cells(j).Value IsNot DBNull.Value Then
                    If DGV.Rows(i).Cells(j).Value.ToString <> "N" Then
                        DGV.Rows(i).Cells(j).Value = Format((DGV.Rows(i).Cells(j).Value * Koeficjent), "0.000")  'Energija kcal...
                    End If
                End If
            Next

        End With

        UkupneVrijednosti()
        ObrociNutrijentiUkupno()
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As System.EventArgs)


    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub VScrollBar1_Scroll(sender As System.Object, e As System.Windows.Forms.ScrollEventArgs)


    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub EditiranjeNamirniceForma_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        

    End Sub

    Private Sub EditiranjeNamirniceForma_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub HScrollBar1_DragOver(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles HScrollBar1.DragOver

    End Sub

    Private Sub HScrollBar1_Move(sender As Object, e As System.EventArgs) Handles HScrollBar1.Move


    End Sub

    Private Sub HScrollBar1_Scroll(sender As System.Object, e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        'On Error Resume Next
        ' Dim Pomak As Integer = 0
        Me.RadioButton1.Checked = True
        Dim AktivniLabel As Label = Me.Label4


        '       If Me.Label9.Text < Me.HScrollBar1.Value Then
        'Pomak = Pomak + Me.HScrollBar1.Value / 10
        If AktivniLabel.Text = "" Then Exit Sub
        If AktivniLabel.Text > 9.8 Then
            AktivniLabel.Text = 9.8
            Exit Sub
        End If
        If AktivniLabel.Text < 0.2 Then
            AktivniLabel.Text = 0.2
            Exit Sub
        End If

        AktivniLabel.Text = Me.HScrollBar1.Value / 10

        MjeraEditiranjeNamirnice()
        EditiranjeNamirnica()

        UkupneVrijednosti()
        ObrociNutrijentiUkupno()

        '  Me.Label9.Text = Me.HScrollBar1.Value

    End Sub

   

    Private Sub HScrollBar2_Scroll(sender As System.Object, e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar2.Scroll
        'On Error Resume Next
        ' Dim Pomak As Integer = 0
        Me.RadioButton2.Checked = True
        Dim AktivniLabel As Label = Me.Label6   'masa


        '       If Me.Label9.Text < Me.HScrollBar1.Value Then
        'Pomak = Pomak + Me.HScrollBar1.Value / 10
        If AktivniLabel.Text = "" Then Exit Sub
        If AktivniLabel.Text > 999.8 Then
            AktivniLabel.Text = 999.8
            Exit Sub
        End If
        If AktivniLabel.Text < 0.2 Then
            AktivniLabel.Text = 0.2
            Exit Sub
        End If
        If Me.Label7.Text < 0.002 Then   'kolicina
            Me.Label7.Text = 0.002
            Exit Sub
        End If
        If Me.Label4.Text < 0.002 Then    'serviranje
            Me.Label4.Text = 0.002
            Exit Sub
        End If


        AktivniLabel.Text = Me.HScrollBar2.Value / 10

        MjeraEditiranjeNamirnice()
        EditiranjeNamirnica()

        UkupneVrijednosti()
        ObrociNutrijentiUkupno()


    End Sub

    Private Sub Label4_Click(sender As System.Object, e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub HScrollBar3_Scroll(sender As System.Object, e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar3.Scroll
        'On Error Resume Next

        Me.RadioButton3.Checked = True
        Dim AktivniLabel As Label = Me.Label7   'kolicina
       
        If AktivniLabel.Text = "" Then Exit Sub
        If AktivniLabel.Text > 9.8 Then
            AktivniLabel.Text = 9.8
            Exit Sub
        End If
        If AktivniLabel.Text < 0.2 Then
            AktivniLabel.Text = 0.2
            Exit Sub
        End If
     

        AktivniLabel.Text = Me.HScrollBar3.Value / 10

        MjeraEditiranjeNamirnice()
        EditiranjeNamirnica()

        UkupneVrijednosti()
        ObrociNutrijentiUkupno()
    End Sub

    

    Private Sub Label7_TextChanged(sender As Object, e As System.EventArgs) Handles Label7.TextChanged

    End Sub

    Private Sub Label7_Click(sender As System.Object, e As System.EventArgs) Handles Label7.Click

    End Sub
End Class