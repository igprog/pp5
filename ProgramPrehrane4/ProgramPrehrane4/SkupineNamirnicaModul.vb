Module SkupineNamirnicaModul
    Sub SkupineNamirnica()
        With Form1
            If .Label15.Text = "Moje namirnice" Or .Label15.Text = "Favoriti" Then
                OdaberiNamirnicu()
                .DataGridView2.DataSource = .SveNamirniceBindingSource
                .Label15.Text = "Sve namirnice"
                .DataGridView2.CurrentRow.Selected = False  'Namirnice
                '    .TabControl1.SelectedIndex = 3
                '  .TabControl1.SelectedIndex = 4
            End If

            Dim SN As BindingSource = .SveNamirniceBindingSource   'Sve namirnice
            Dim DGV As DataGridView = .DataGridView2
            SN.RemoveFilter()
            .Label15.Text = .TreeView1.SelectedNode.Text
            If DGV.CurrentRow IsNot Nothing Then
                DGV.CurrentRow.Selected = False
            End If

            DGV.DataSource = SN  'Sve namirnice
            SN.RemoveFilter()

            '      .Label15.Text = "Sve namirnice"

            If .TreeView1.SelectedNode.Name = "Node0" Then
                DGV.DataSource = SN  'Sve namirnice
                SN.RemoveFilter()
                'DGV.Rows(1)
            End If

            If .TreeView1.SelectedNode.Name = "Node1" Then
                '        .DataGridView2.DataSource = .ZitariceBindingSource
                SN.RemoveFilter()
                ' SN.Filter = "Zitarice='" & "1" & "'"
                SN.Filter = "SkupinaNamirnica='" & "Zitarice" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node2" Then
                '          .DataGridView2.DataSource = .PovrceBindingSource
                'SN.Filter = "Povrce='" & "1" & "'"
                SN.Filter = "SkupinaNamirnica='" & "Povrce" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node3" Then
                '          .DataGridView2.DataSource = .VoceBindingSource
                'DGV.Rows(i).Cells(7).Value 
                SN.Filter = "SkupinaNamirnica='" & "Voce" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node4" Then
                '          .DataGridView2.DataSource = .MesoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "IzrazitoNemasnoMeso" & "'" & _
                "OR SkupinaNamirnica='" & "NemasnoMeso" & "'" & _
                "OR SkupinaNamirnica='" & "SrednjeMasnoMeso" & "'" & _
                "OR SkupinaNamirnica='" & "MasnoMeso" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node5" Then
                '         .DataGridView2.DataSource = .IzrazitoNemasnoMesoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "IzrazitoNemasnoMeso" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node6" Then
                '        .DataGridView2.DataSource = .NemasnoMesoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "NemasnoMeso" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node7" Then
                '         .DataGridView2.DataSource = .SrednjeMasnoMesoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "SrednjeMasnoMeso" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node8" Then
                '         .DataGridView2.DataSource = .MasnoMesoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "MasnoMeso" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node9" Then
                '         .DataGridView2.DataSource = .MlijekoBindingSource
                'SN.Filter = "SkupinaNamirnica='" & "Ml" & "'"
                '/// TREBA DODATI ILI ISPRAVITI ObranoMlijko-----///
                SN.Filter = "SkupinaNamirnica='" & "ObranoMlijeko" & "'" & _
          "OR SkupinaNamirnica='" & "DjelomicnoObranoMlijeko" & "'" & _
          "OR SkupinaNamirnica='" & "PunomasnoMlijeko" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node10" Then
                '           .DataGridView2.DataSource = .NemasnoINiskomasnoMlijekoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "ObranoMlijeko" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node11" Then
                '           .DataGridView2.DataSource = .DjelomicnoObranoMlijekoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "DjelomicnoObranoMlijeko" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node12" Then
                '         .DataGridView2.DataSource = .PunomasnoMlijekoBindingSource
                SN.Filter = "SkupinaNamirnica='" & "PunomasnoMlijeko" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node13" Then
                '         .DataGridView2.DataSource = .MastiBindingSource
                ' SN.Filter = "SkupinaNamirnica='" & "PunomasnoMlijeko" & "'"
                SN.Filter = "SkupinaNamirnica='" & "ZasiceneMasti" & "'" & _
           "OR SkupinaNamirnica='" & "VisestrukoNezasiceneMasti" & "'" & _
          "OR SkupinaNamirnica='" & "JednostrukoNezasiceneMasti" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node14" Then
                '         .DataGridView2.DataSource = .ZasiceneMastiBindingSource
                SN.Filter = "SkupinaNamirnica='" & "ZasiceneMasti" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node15" Then
                '        .DataGridView2.DataSource = .VisestrukoNezasiceneMastiBindingSource
                SN.Filter = "SkupinaNamirnica='" & "VisestrukoNezasiceneMasti" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node16" Then
                '          .DataGridView2.DataSource = .JednostrukoNezasiceneMastiBindingSource
                SN.Filter = "SkupinaNamirnica='" & "JednostrukoNezasiceneMasti" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node17" Then
                '          .DataGridView2.DataSource = .MjesoviteNamirniceBindingSource
                SN.Filter = "SkupinaNamirnica='" & "MjesoviteNamirnice" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node18" Then
                '           .DataGridView2.DataSource = .OstaleNamirniceBindingSource
                SN.Filter = "SkupinaNamirnica='" & "OstaleNamirnice" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node21" Then
                '           .DataGridView2.DataSource = .OstaleNamirniceBindingSource
                SN.Filter = "SkupinaNamirnica='" & "Jela" & "'"
            End If

            If .TreeView1.SelectedNode.Name = "Node19" Then
                DGV.DataSource = .MojeNamirniceBindingSource
            End If

            If .Label15.Text = "Moje namirnice" Then   'Napomena
                .Label311.Visible = True
            Else
                .Label311.Visible = False
            End If

            If .TreeView1.SelectedNode.Name = "Node20" Then
                DGV.DataSource = .FavoritiBindingSource
            End If

            'OdabirNamirnica()    'TreeView1

        End With
    End Sub
End Module
