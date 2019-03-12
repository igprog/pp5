Module ObrociModul
    Sub Obroci()
        On Error Resume Next
        With Form1
            'obavezni obroci: dorucak, rucak, vecera
            If .CheckBox2.Checked = False Or .CheckBox4.Checked = False Or .CheckBox6.Checked = False Then
                MsgBox("Obroci " & .ComboBox15.Text & ", " & .ComboBox22.Text & " i " & .ComboBox24.Text & " su obavezni.")
                .TabControl1.SelectedIndex = 4
                Exit Sub
            End If

            'obrok pred spavanje je dozvoljen samo u kombinaciji sa svim ostalim obrocima
            If .CheckBox7.Checked = True Then
                If .CheckBox2.Checked = False Or .CheckBox3.Checked = False Or .CheckBox4.Checked = False Or .CheckBox5.Checked = False Or .CheckBox6.Checked = False Then
                    MsgBox("Odabrana kombinacija obroka u jelovniku nije dozvoljena." & vbCrLf & .ComboBox25.Text & " u ovoj kombinaciji mora biti isključen.")
                    .TabControl1.SelectedIndex = 4
                    Exit Sub
                End If
            End If

            .ListBox84.Items.Clear()
            If .ComboBox15.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox15.Text)
            End If
            If .ComboBox21.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox21.Text)
            End If
            If .ComboBox22.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox22.Text)
            End If
            If .ComboBox23.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox23.Text)
            End If
            If .ComboBox24.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox24.Text)
            End If
            If .ComboBox25.Enabled = True Then
                .ListBox84.Items.Add(.ComboBox25.Text)
            End If

            Dim i As Integer = 0

            'Obrok 1 (doručak)
            .TabPage7.Text = .ComboBox15.Text
            .GroupBox23.Text = .ComboBox15.Text
            If .CheckBox2.Checked = True Then
                If .Label347.Text = 0 Then   'ako nema obroka
                    .ComboBox15.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox15.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage7) 'vrati obrok
                            .Label347.Text = 1 'vracen obrok
                        Else
                            .ComboBox15.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage7)
                .ComboBox15.Enabled = False
                .Label347.Text = 0   'nema obroka
            End If

            'Obrok 2 (jutaranj uzina)
            .TabPage8.Text = .ComboBox21.Text
            .GroupBox26.Text = .ComboBox21.Text
            If .CheckBox3.Checked = True Then
                If .Label348.Text = 0 Then   'ako nema obroka
                    .ComboBox21.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox21.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage8) 'vrati obrok
                            .Label348.Text = 1 'vracen obrok
                        Else
                            .ComboBox21.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage8)
                .ComboBox21.Enabled = False
                .Label348.Text = 0   'nema obroka
            End If

            'Obrok 3 (rucak)
            .TabPage9.Text = .ComboBox22.Text
            .GroupBox30.Text = .ComboBox22.Text
            If .CheckBox4.Checked = True Then
                If .Label349.Text = 0 Then   'ako nema obroka
                    .ComboBox22.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox22.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage9) 'vrati obrok
                            .Label349.Text = 1 'vracen obrok
                        Else
                            .ComboBox22.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage9)
                .ComboBox22.Enabled = False
                .Label349.Text = 0   'nema obroka
            End If

            'Obrok 4 (popodnevna uzina)
            .TabPage10.Text = .ComboBox23.Text
            .GroupBox34.Text = .ComboBox23.Text
            If .CheckBox5.Checked = True Then
                If .Label350.Text = 0 Then   'ako nema obroka
                    .ComboBox23.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox23.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage10) 'vrati obrok
                            .Label350.Text = 1 'vracen obrok
                        Else
                            .ComboBox23.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage10)
                .ComboBox23.Enabled = False
                .Label350.Text = 0   'nema obroka
            End If

            'Obrok 5 (vecera)
            .TabPage11.Text = .ComboBox24.Text
            .GroupBox38.Text = .ComboBox24.Text
            If .CheckBox6.Checked = True Then
                If .Label351.Text = 0 Then   'ako nema obroka
                    .ComboBox24.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox24.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage11) 'vrati obrok
                            .Label351.Text = 1 'vracen obrok
                        Else
                            .ComboBox24.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage11)
                .ComboBox24.Enabled = False
                .Label351.Text = 0   'nema obroka
            End If

            'Obrok 6 (obrok pred spavanje)
            .TabPage12.Text = .ComboBox25.Text
            .GroupBox42.Text = .ComboBox25.Text
            If .CheckBox7.Checked = True Then
                If .Label352.Text = 0 Then   'ako nema obroka
                    .ComboBox25.Enabled = True
                    For i = 0 To .ListBox84.Items.Count - 1
                        If .ListBox84.Items(i).ToString = .ComboBox25.Text Then
                            .TabControl2.TabPages.Insert(i, .TabPage12) 'vrati obrok
                            .Label352.Text = 1 'vracen obrok
                        Else
                            .ComboBox25.Enabled = True
                        End If
                    Next
                End If
            Else
                .TabControl2.TabPages.Remove(.TabPage12)
                .ComboBox25.Enabled = False
                .Label352.Text = 0   'nema obroka
            End If

        End With
    End Sub
End Module
