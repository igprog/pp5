Module TermickeObradeModule
    Sub TermickeObrade()
        On Error Resume Next
        With Form1
            Dim i As Integer = .DataGridView2.CurrentRow.Index
            Dim SkupinaNamirnica As String = .DataGridView2.Rows(i).Cells(56).Value  'zadnja kolona DGV2 GubiciVitamina, zitarice, mlijeko, jaja, meso, riba, povrce, voce
            Dim TermickeObrade_CB As ComboBox = .ComboBox11   'termicke obrade, kuhanje, pecenje...)
            TermickeObrade_CB.Items.Clear()

            If SkupinaNamirnica = "zitarice" Then
                TermickeObrade_CB.Items.Add("kuhanje")
                TermickeObrade_CB.Items.Add("pečenje")
            End If

            If SkupinaNamirnica = "povrce" Then
                TermickeObrade_CB.Items.Add("kuhanje")
                TermickeObrade_CB.Items.Add("prženje")
                TermickeObrade_CB.Items.Add("povrtna jela")
            End If

            If SkupinaNamirnica = "voce" Then
                TermickeObrade_CB.Items.Add("pirjanje")
            End If

            If SkupinaNamirnica = "meso" Then
                TermickeObrade_CB.Items.Add("roštilj/prženje")
                ' TermickeObrade_CB.Items.Add("prženje")
                TermickeObrade_CB.Items.Add("mesna jela")
            End If

            If SkupinaNamirnica = "riba" Then
                TermickeObrade_CB.Items.Add("pirjanje u vodi")
                TermickeObrade_CB.Items.Add("pečenje")
                TermickeObrade_CB.Items.Add("roštilj")
                TermickeObrade_CB.Items.Add("prženje")
            End If

            If SkupinaNamirnica = "jaja" Then
                TermickeObrade_CB.Items.Add("kajgana/omlet")
                'TermickeObrade_CB.Items.Add("omlet")
                TermickeObrade_CB.Items.Add("pečenje")
            End If

            If SkupinaNamirnica = "mlijeko" Then
                TermickeObrade_CB.Items.Add("kuhanje")
                TermickeObrade_CB.Items.Add("umaci")
                TermickeObrade_CB.Items.Add("pečenje")
            End If

            If SkupinaNamirnica = "N" Then
                TermickeObrade_CB.Enabled = False
            Else
                TermickeObrade_CB.Enabled = True
            End If

        End With
    End Sub
End Module
