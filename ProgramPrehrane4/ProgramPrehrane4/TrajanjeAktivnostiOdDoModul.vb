Module TrajanjeAktivnostiOdDoModul
    Sub TrajanjeAktivnostiOdDo()
        With Form1
            On Error Resume Next
            If Val(.ComboBox19.Text) < Val(.TextBox83.Text) Or Val(.ComboBox19.Text) + (Val(.ComboBox20.Text) / 60) > 24 Then Exit Sub

            '     If Val(.ComboBox19.Text) < Val(.TextBox83.Text) Then
            'MsgBox("Greška. " & Val(.ComboBox19.Text) - Val(.TextBox83.Text))
            'Exit Sub
            'Else

            Dim i As Integer
            Dim DGV As DataGridView
            DGV = .DataGridView17
            Dim Min As Double = 0
            Dim Energ As Double = 0
            For i = 0 To DGV.RowCount - 1
                Min = Min + DGV.Rows(i).Cells(9).Value
            Next i
            If Min = 60 * 24 Then
                .RadioButton3.Checked = False
                .RadioButton4.Checked = False
                .RadioButton5.Checked = False
                .RadioButton6.Checked = False
                .RadioButton7.Checked = False
                .RadioButton8.Checked = False
                .RadioButton9.Checked = False
                .RadioButton10.Checked = False
            End If
            If Min > 60 * 24 Then
                MsgBox("Error. >24h")

                '   .ComboBox14.SelectedIndex = .ComboBox14.SelectedIndex + 1
                Exit Sub
            End If
            Dim TrajanjeAktivnostiOd As Integer = Val(.TextBox83.Text) * 60 + Val(.TextBox84.Text) 'minut od
            Dim TrajanjeAktivnostiDo As Integer = Val(.ComboBox19.Text) * 60 + Val(.ComboBox20.Text)  'minut do
            .Label336.Text = TrajanjeAktivnostiDo - TrajanjeAktivnostiOd   'trajanje aktivnosti (min)
            .Label362.Text = .TextBox83.Text & ":" & .TextBox84.Text & " h"

            '  End If
        End With
    End Sub
End Module
