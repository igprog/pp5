Imports System.Globalization
Imports System.Windows.Forms.DataVisualization    'print chart
Public Class Form1

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
            MsgBox("Odaberite dob.")
            TabControl3.SelectedIndex = 0
            Exit Sub
        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("Odaberite spol.")
            TabControl3.SelectedIndex = 0
            Exit Sub
        End If
        If ComboBox2.Text = "" Then
            MsgBox("Odaberite visinu u cm.")
            TabControl3.SelectedIndex = 0
            Exit Sub
        End If
        If ComboBox3.Text = "" Then
            MsgBox("Odaberite masu u kg.")
            TabControl3.SelectedIndex = 0
            Exit Sub
        End If

        If Me.Label315.Text = 0 Then
            If Val(ComboBox1.Text) < 18 And Val(ComboBox1.Text) > 15 Then
                If RadioButton3.Checked = False And RadioButton4.Checked = False _
                           And RadioButton5.Checked = False And RadioButton6.Checked = False Then
                    MsgBox("Odaberite intenzitet tjelesne aktivnosti.")
                    TabControl3.SelectedIndex = 0
                    Exit Sub
                End If
            End If
            If Val(ComboBox1.Text) > 18 Then
                If RadioButton3.Checked = False And RadioButton4.Checked = False _
                    And RadioButton5.Checked = False And RadioButton6.Checked = False Then
                    MsgBox("Odaberite intenzitet tjelesne aktivnosti na poslu.")
                    TabControl3.SelectedIndex = 0
                    Exit Sub
                End If
                If RadioButton7.Checked = False And RadioButton8.Checked = False _
                    And RadioButton9.Checked = False And RadioButton10.Checked = False Then
                    MsgBox("Odaberite intenzitet tjelesne aktivnosti izvan posla.")
                    TabControl3.SelectedIndex = 0
                    Exit Sub
                End If
            End If
        End If

        If Val(ComboBox1.Text) < 18 And Val(ComboBox1.Text) >= 15 And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then
            If MessageBox.Show("Izračun i preporuke za osobe od 15 do 18 godina je dostupan samo u PREMIUM verziji programa." _
                                     & vbCrLf & "Želite li naručiti aktivacijski kod za pokretanje PREMIUM verzije?", "Program Prehrane 5.0", _
                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                   = DialogResult.Yes Then
                System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
            End If
            ' MsgBox("Program ne daje izračun energetske potrošnje za djecu mlađu od 9 godina!")
            TextBox3.Text = ""
            TextBox4.Text = ""
            Exit Sub
        End If
        If Val(ComboBox1.Text) < 15 Then
            MsgBox("Program ne daje izračun energetske potrošnje za djecu mlađu od 15 godina!")
            TextBox3.Text = ""
            TextBox4.Text = ""
            Exit Sub
        End If
        BMI()   'IZRACUN

        'Spremi podatke u bazu korisnika

        'Provjera dali vec postoji podatak za taj datum u bazi
        Dim DGV As DataGridView = Me.DataGridView6
        Dim i As Integer
        For i = 0 To DGV.RowCount - 1
            If DGV.Rows(i).Cells(2).Value IsNot DBNull.Value _
              And DGV.Rows(i).Cells(3).Value IsNot DBNull.Value _
              And DGV.Rows(i).Cells(13).Value IsNot DBNull.Value Then

                If DGV.Rows(i).Cells(2).Value = Me.TextBox1.Text _
                     And DGV.Rows(i).Cells(3).Value = Me.TextBox2.Text Then
                    'And DGV.Rows(i).Cells(13).Value = Me.DateTimePicker2.Value.Date Then
                    '   MsgBox("Podaci za " & Me.DateTimePicker2.Value.Date & " već postoje u bazi.")
                    TabControl1.SelectedIndex = 1   'izračun
                    Exit Sub
                Else
                    'Pitanje
                    If My.Settings.PP5PremiumAktivacija = "Da" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Da" Then
                        If MessageBox.Show("Želite li spremiti unesene podatke u bazu klijenata?", "Program Prehrane 5.0", _
                                                   MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                                   = DialogResult.Yes Then
                            TabControl3.SelectedIndex = 3   'Baza klijenata
                            SpremiKorisnika()
                            TabControl3.SelectedIndex = 0

                            TabControl3.SelectedIndex = 2
                            TabControl4.SelectedIndex = 0   'Pracenje antropometrijskih parametara
                            PracenjeStanjaDodaj()
                            TabControl3.SelectedIndex = 0

                            TabControl1.SelectedIndex = 1

                            'Brisi detaljni izracun energetske potrosnje
                            On Error Resume Next
                            With Me

                                Dim DGV1 As DataGridView = .DataGridView17
                                Dim BS As BindingSource = BazaEnergetskePotrosnjeBindingSource
                                Dim j As Integer

                                'briši odabrane aktivnosti

                                For j = 0 To DGV.RowCount - 1
                                    DGV1.Rows.Remove(DGV1.CurrentRow)
                                Next j
                                BS.AddNew()

                                .DataGridView17.CurrentRow.Selected = False

                                .TextBox83.Text = 0
                                .TextBox84.Text = 0
                                .ComboBox19.Text = 0
                                .ComboBox20.Text = 0
                                .Label315.Text = 0

                            End With

                            TrajanjeAktivnostiOdDo()
                            EnergetskaPotrosnjaUkupno()

                        End If

                    End If
                End If
            End If
        Next i
        TabControl1.SelectedIndex = 1   'izračun
        Me.Label315.Text = 0   'TEE (detaljni izracun)


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Val(TextBox4.Text) = 0 Then
            TabControl1.SelectedIndex = 3
        Else
            TabControl1.SelectedIndex = 2
        End If

        VrstaPrehranePreporuka()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TabControl1.SelectedIndex = 0

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TabControl1.SelectedIndex = 4

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If Val(TextBox4.Text) = 0 Then
            TabControl1.SelectedIndex = 1
        Else
            TabControl1.SelectedIndex = 2
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TabControl1.SelectedIndex = 3
        Me.DataGridView1.CurrentRow.Selected = False   'vrsta prehrane/dijete

        VrstaPrehranePreporuka()

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        TabControl1.SelectedIndex = 1

    End Sub

    Private Sub TabPage2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage2.Click

    End Sub

    Private Sub TabPage2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage2.Enter

    End Sub

    Private Sub TabPage4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage4.Click

    End Sub

    Private Sub TabPage4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage4.Enter
        On Error Resume Next
        If Val(Me.TextBox3.Text) > 500 And Val(Me.TextBox3.Text) < 20000 Then
            If Me.Label179.Text < 1 Or Me.Label179.Text > 50 Then
                MsgBox("Odaberite vrstu prehrane / dijete.")
                TabControl1.SelectedIndex = 3
                Exit Sub
            End If
        End If

        PreporuceniBrojServiranja()   'preporuceni broj serviranja - normalna prehrana i ostale dijete
        Dim BrojDijete As Integer = Me.Label179.Text
        If BrojDijete = 9 Then PrepBrojServLagana() 'preporuceni broj serviranja - lagana dijeta
        If BrojDijete = 14 Then PrepBrojServJetra() 'preporuceni broj serviranja - kronicne bolesti jetre
        If BrojDijete = 15 Then PrepBrojServCrijeva() 'preporuceni broj serviranja - upalne bolesti crijeva
        If BrojDijete >= 18 And BrojDijete <= 22 Then PrepBrojServDijabet() 'preporuceni broj serviranja - dijabeticke dijete
        If BrojDijete = 8 Then PrepBrojServGlikogen() 'preporuceni broj serviranja - zavrsa faza deponiranja glikogena
        If BrojDijete = 23 Then PrepBrojServVege() 'preporuceni broj serviranja - laktoovo vegetarijanska dijeta

        Obroci()
        OstaleNamirnicePreporuka()
        NutrijentiPreporuceniPostoci()
        UkupneVrijednosti()
        ObrociNutrijentiUkupno()
        OdaberiNamirnicu()
        Me.DataGridView2.DataSource = Me.SveNamirniceBindingSource
        Me.Label15.Text = "Sve namirnice"
        Me.Label311.Visible = False
        '  Me.TreeView1.Nodes(0).IsSelected = True
        ' me.treeView1.Nodes[0].Selected=true
        'TreeView1.Nodes[4].Selected = true
        DataGridView2.CurrentRow.Selected = False  'Namirnice

        'Obroci
        With Me


            'graf - pita
            If .ListBox10.Items(0) = "0 kcal" Then
                .Chart2.Series(0).Points.Clear()
                .Chart2.Series(0).Points.AddY(20)
                .Chart2.Series(0).Points.AddY(30)
                .Chart2.Series(0).Points.AddY(50)
            End If


            If Val(.TextBox3.Text) <= 0 Then
                .ListBox13.Items.Clear()
                .ListBox37.Items.Clear()
                .Label281.Text = ""
                .Label283.Text = ""
                .ListBox11.Items.Clear()
                .ListBox38.Items.Clear()
                'obroci
                .ListBox12.Items.Clear()
                .ListBox59.Items.Clear()
                .ListBox64.Items.Clear()
                .ListBox69.Items.Clear()
                .ListBox74.Items.Clear()
                .ListBox79.Items.Clear()
                .ListBox39.Items.Clear()
                .ListBox60.Items.Clear()
                .ListBox65.Items.Clear()
                .ListBox70.Items.Clear()
                .ListBox75.Items.Clear()
                .ListBox80.Items.Clear()
            End If

            'BOJA
            '         Dim DGV As DataGridView = Me.DataGridView2
            '        Dim i As Integer = DGV.CurrentRow.Index
            'Namirnice
            '       For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.BackColor = My.Settings.DgvBoja
            'Next
            'BOJA-SELEKTIRANI RED
            'Sve namirnice
            '   For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = My.Settings.DgvBojaSelect
            'Next

        End With

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        DanUTjednu1()

    End Sub

    Private Sub Form1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

      


        '  On Error Resume Next
        ' With Me
        ''BOJA
        '     Dim DGV As DataGridView = .DataGridView2
        '    Dim i As Integer = DGV.CurrentRow.Index
        '   'Sve Namirnice
        '  DGV = .DataGridView2
        ' For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.BackColor = My.Settings.DgvBoja
        '' DGV.Rows(i).DefaultCellStyle.BackColor = SystemColors.GradientInactiveCaption
        'Next

        '  'BOJA-SELEKTIRANI RED
        ''Sve namirnice
        '    For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = My.Settings.DgvBojaSelect
        'Next



        '   End With
    End Sub



    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        On Error Resume Next
        With Me
            .BazaKorisnikaBindingSource.EndEdit()
            .BazaKorisnikaTableAdapter.Update(.PP5aDataSet)

            .KorisniciPracenjeStanjaBindingSource.EndEdit()
            .KorisniciPracenjeStanjaTableAdapter.Update(.PP5aDataSet)

            .BazaNazivaJelovnikaBindingSource.EndEdit()
            .BazaNazivaJelovnikaTableAdapter.Update(.PP5aDataSet)

            .BazaJelovnikaBindingSource.EndEdit()
            .BazaJelovnikaTableAdapter.Update(.PP5aDataSet)

            .MojeNamirniceBindingSource.EndEdit()
            .MojeNamirniceTableAdapter.Update(.PP5aDataSet)

            .FavoritiBindingSource.EndEdit()
            .FavoritiTableAdapter.Update(.PP5aDataSet)

            .CijeneBindingSource.EndEdit()
            .CijeneTableAdapter.Update(.PP5aDataSet)

            If Me.Label315.Text > 0 Then
                .BazaEnergetskePotrosnjeBindingSource.EndEdit()
                .BazaEnergetskePotrosnjeTableAdapter.Update(.PP5aDataSet)
            End If

        End With
        End
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PP5aDataSet.KorisniciPracenjeStanja' table. You can move, or remove it, as needed.
        Me.KorisniciPracenjeStanjaTableAdapter.Fill(Me.PP5aDataSet.KorisniciPracenjeStanja)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaKorisnika' table. You can move, or remove it, as needed.
        Me.BazaKorisnikaTableAdapter.Fill(Me.PP5aDataSet.BazaKorisnika)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaEnergetskePotrosnje' table. You can move, or remove it, as needed.
        Me.BazaEnergetskePotrosnjeTableAdapter.Fill(Me.PP5aDataSet.BazaEnergetskePotrosnje)
        'TODO: This line of code loads data into the 'PP5DataSet.SveTjelesneAktivnosti' table. You can move, or remove it, as needed.
        Me.SveTjelesneAktivnostiTableAdapter.Fill(Me.PP5DataSet.SveTjelesneAktivnosti)
        'TODO: This line of code loads data into the 'PP5DataSet.OdabranaDodatnaTjelesnaAktivnost' table. You can move, or remove it, as needed.
        Me.OdabranaDodatnaTjelesnaAktivnostTableAdapter.Fill(Me.PP5DataSet.OdabranaDodatnaTjelesnaAktivnost)
        'TODO: This line of code loads data into the 'PP5DataSet.SportskeAktivnosti' table. You can move, or remove it, as needed.
        Me.SportskeAktivnostiTableAdapter.Fill(Me.PP5DataSet.SportskeAktivnosti)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaJelovnika' table. You can move, or remove it, as needed.
        Me.BazaJelovnikaTableAdapter.Fill(Me.PP5aDataSet.BazaJelovnika)
        'TODO: This line of code loads data into the 'PP5DataSet.BazaPrimjeraJelovnika' table. You can move, or remove it, as needed.
        Me.BazaPrimjeraJelovnikaTableAdapter.Fill(Me.PP5DataSet.BazaPrimjeraJelovnika)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaNazivaJelovnika' table. You can move, or remove it, as needed.
        Me.BazaNazivaJelovnikaTableAdapter.Fill(Me.PP5aDataSet.BazaNazivaJelovnika)
        'TODO: This line of code loads data into the 'PP5DataSet.BazaNazivaPrimjeraJelovnika' table. You can move, or remove it, as needed.
        Me.BazaNazivaPrimjeraJelovnikaTableAdapter.Fill(Me.PP5DataSet.BazaNazivaPrimjeraJelovnika)
        'TODO: This line of code loads data into the 'PP5DataSet.SveNamirnice' table. You can move, or remove it, as needed.
        Me.SveNamirniceTableAdapter.Fill(Me.PP5DataSet.SveNamirnice)
        'TODO: This line of code loads data into the 'PP5aDataSet.Favoriti' table. You can move, or remove it, as needed.
        Me.FavoritiTableAdapter.Fill(Me.PP5aDataSet.Favoriti)
        'TODO: This line of code loads data into the 'PP5aDataSet.MojeNamirnice' table. You can move, or remove it, as needed.
        Me.MojeNamirniceTableAdapter.Fill(Me.PP5aDataSet.MojeNamirnice)
        'TODO: This line of code loads data into the 'PP5aDataSet.Cijene' table. You can move, or remove it, as needed.
        Me.CijeneTableAdapter.Fill(Me.PP5aDataSet.Cijene)
        'TODO: This line of code loads data into the 'PP5aDataSet.Favoriti' table. You can move, or remove it, as needed.
        Me.FavoritiTableAdapter.Fill(Me.PP5aDataSet.Favoriti)
        'TODO: This line of code loads data into the 'PP5aDataSet.MojeNamirnice' table. You can move, or remove it, as needed.
        Me.MojeNamirniceTableAdapter.Fill(Me.PP5aDataSet.MojeNamirnice)
        'TODO: This line of code loads data into the 'PP5aDataSet.KorisniciPracenjeStanja' table. You can move, or remove it, as needed.
        Me.KorisniciPracenjeStanjaTableAdapter.Fill(Me.PP5aDataSet.KorisniciPracenjeStanja)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaKorisnika' table. You can move, or remove it, as needed.
        Me.BazaKorisnikaTableAdapter.Fill(Me.PP5aDataSet.BazaKorisnika)
        'TODO: This line of code loads data into the 'PP5DataSet.SportskeAktivnosti' table. You can move, or remove it, as needed.
        Me.SportskeAktivnostiTableAdapter.Fill(Me.PP5DataSet.SportskeAktivnosti)
        'TODO: This line of code loads data into the 'PP5DataSet.VrstaDijete' table. You can move, or remove it, as needed.
        Me.VrstaDijeteTableAdapter.Fill(Me.PP5DataSet.VrstaDijete)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaJelovnika' table. You can move, or remove it, as needed.
        Me.BazaJelovnikaTableAdapter.Fill(Me.PP5aDataSet.BazaJelovnika)
        'TODO: This line of code loads data into the 'PP5aDataSet.BazaNazivaJelovnika' table. You can move, or remove it, as needed.
        Me.BazaNazivaJelovnikaTableAdapter.Fill(Me.PP5aDataSet.BazaNazivaJelovnika)
        'TODO: This line of code loads data into the 'PP5DataSet.ObrokPredSpavanje' table. You can move, or remove it, as needed.
        Me.ObrokPredSpavanjeTableAdapter.Fill(Me.PP5DataSet.ObrokPredSpavanje)
        'TODO: This line of code loads data into the 'PP5DataSet.Vecera' table. You can move, or remove it, as needed.
        Me.VeceraTableAdapter.Fill(Me.PP5DataSet.Vecera)
        'TODO: This line of code loads data into the 'PP5DataSet.PopodnevnaUzina' table. You can move, or remove it, as needed.
        Me.PopodnevnaUzinaTableAdapter.Fill(Me.PP5DataSet.PopodnevnaUzina)
        'TODO: This line of code loads data into the 'PP5DataSet.Rucak' table. You can move, or remove it, as needed.
        Me.RucakTableAdapter.Fill(Me.PP5DataSet.Rucak)
        'TODO: This line of code loads data into the 'PP5DataSet.JutarnjaUzina' table. You can move, or remove it, as needed.
        Me.JutarnjaUzinaTableAdapter.Fill(Me.PP5DataSet.JutarnjaUzina)
        'TODO: This line of code loads data into the 'PP5DataSet.Dorucak' table. You can move, or remove it, as needed.
        Me.DorucakTableAdapter.Fill(Me.PP5DataSet.Dorucak)
        'TODO: This line of code loads data into the 'PP5DataSet.SveNamirnice' table. You can move, or remove it, as needed.
        Me.SveNamirniceTableAdapter.Fill(Me.PP5DataSet.SveNamirnice)

        Application.CurrentCulture = New CultureInfo("hr")  'Hrvatske regionalne postavke



        Me.Label176.Text = 0    'Dodatna energetska potrošnja

        DanUTjednu()    'Dan u tjednu

        Dijeta_BindingSource()   'Modul  - binding source - dijeta

        BazaJelovnika_BindingSource()   'Modul - binding source  - baza jelovnika

        PracenjeStanja_BindingSource()   'Modul - binding source - pracenje stanja

        BazaKorisnika_BindingSource()    'Modul - binding source - baza korisnika

        AktivnostiBindingSource()        'Modul - binding source - aktivnosti

        MojeNamirnice_BindingSource()      'Modul - binding source - moje namirnice

        EnergetskaPotrosnjaBindingSource()    'Modul - binding source - energetska potrosnja

        ' DgvStyle()    'Modul - DataGridView Style

        'NAZIVI OBROKA
        With Me
            .ComboBox15.Text = My.Settings.Obrok1.ToString
            .ComboBox21.Text = My.Settings.Obrok2.ToString
            .ComboBox22.Text = My.Settings.Obrok3.ToString
            .ComboBox23.Text = My.Settings.Obrok4.ToString
            .ComboBox24.Text = My.Settings.Obrok5.ToString
            .ComboBox25.Text = My.Settings.Obrok6.ToString

            'VALUTA
            .ComboBox26.Text = My.Settings.Valuta.ToString   'Postavke
            .Label322.Text = My.Settings.Valuta.ToString & " /"   'Cijene
            .Label317.Text = My.Settings.Valuta.ToString & " /"    'Jelovnik
        End With

        With Me
            .Label11.Text = ""
            .Label25.Text = ""
            .Label12.Text = ""
            .Label13.Text = ""
            .Label193.Text = ""
            .Label194.Text = ""
            .Label208.Text = ""
            .Label195.Text = ""
            .Label196.Text = ""
            .Label197.Text = ""
            .Label287.Text = ""
            .Label20.Text = ""   'opis tjelesne aktivnosti
            .ComboBox5.Text = ""   'trajanje aktivnosti (min)
            .Label23.Text = ""   'korisnik (pracenje stanja tablicni prikaz)
            'postoci
            .Label83.Text = ""   'uglj
            .Label84.Text = ""   'bjel
            .Label85.Text = ""   'masti

            .Label21.Text = ""   'vrsta prehrane

            'Obroci
            .Label347.Text = 1
            .Label348.Text = 1
            .Label349.Text = 1
            .Label350.Text = 1
            .Label351.Text = 1
            .Label352.Text = 1

            .Label315.Text = 0   'detaljni izracun energetske potrosnje


        End With

        Me.VrstaDijeteBindingSource.Position = 2   'normalna prehrana

        'VERZIJE
        If My.Settings.PP5PremiumAktivacija.ToString = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then
            Demo()
        End If
        If My.Settings.PP5StartAktivacija.ToString = "Da" Or My.Settings.PP5StartTrajnaLicencaAktivacija.ToString = "Da" Then
            Start()
        End If
        If My.Settings.PP5PremiumAktivacija.ToString = "Da" Or My.Settings.PP5PremiumTrajnaLicencaAktivacija.ToString = "Da" Then
            Premium()
        End If


        NevidljiviLabeli()    'Nevidljivi Labeli

        DataGridView3.CurrentRow.Selected = False    'Tjelesne aktivnosti

        Me.TabControl1.SelectedIndex = 5  'Izrada jelovnika
        NoviJelovnik()
        Me.TabControl1.SelectedIndex = 0   'Ulazni podaci

        Me.TextBox1.Select()


    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        If Me.TextBox5.Text = "" Then
            MsgBox("Odaberite namirnicu.")
            Exit Sub
        End If

        If Val(ComboBox1.Text) < 9 Then
            MsgBox("Program nije namijenjen izradi jelovnika za osobe mlađe od 9 godina.")
            Exit Sub
        End If

        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then   'demo
            Dim BrojOdabranihNamirnica As Integer
            BrojOdabranihNamirnica = Me.DataGridView5.RowCount - 1 + Me.DataGridView9.RowCount - 1 _
                + Me.DataGridView11.RowCount - 1 + Me.DataGridView12.RowCount - 1 + Me.DataGridView13.RowCount - 1 _
                + Me.DataGridView14.RowCount - 1
            If BrojOdabranihNamirnica > 12 Then GoTo kraj
            'If Label80.Text > 1500 Then GoTo Kraj
        End If

        If Val(TextBox3.Text) > 500 And Val(TextBox3.Text) < 20000 Then

            'gotova jela
            Dim DGV As DataGridView = Me.DataGridView2
            Dim Index As Integer = DGV.CurrentRow.Index
            Dim GotovoJelo As String = DGV.Rows(Index).Cells(57).Value.ToString
            '  OdaberiNamirnicu()
            'Za lazanje nama podatke o namirnicama
            If DGV.Rows(Index).Cells(2).Value.ToString <> "Lazanje" And GotovoJelo = "Jela" Then
                '  OdaberiNamirnicu()
                ' Exit Sub
                'GotovaJela()
                PripremljenaJela()
                Exit Sub
            End If
            'If GotovoJelo = "Jela" Then
            ''

            GubiciVitamina()
            PrebaciNamirnicu()
            UkupneVrijednosti()
            ObrociNutrijentiUkupno()
            '          FavoritiSpremiForm.Show()

            If Me.CheckBox1.Checked = True Then
                '    TabControl1.SelectedIndex = 9
                '  Me.TextBox78.Text = Me.TextBox5.Text  'naziv namirnice
                FavoritiSpremi()
                TabControl1.SelectedIndex = 5
                '       FavoritiSpremiForm.Close()
            End If

            DataGridView2.CurrentRow.Selected = False
            '  OdaberiNamirnicu()
            Me.TextBox85.Text = 1     'Serviranja
            Me.TextBox5.Text = ""   'Naziv Namirnice
            Me.TextBox12.Text = ""   'Kolicina
            Me.Label229.Text = ""   'Kolicina
            Me.Label103.Text = ""   'Mjera
            Me.Label230.Text = ""   'Mjera
            Me.TextBox6.Text = ""   'Masa_g
            Me.Label228.Text = ""   'Masa_g
            Me.ListBox83.Items.Clear()   'gubici vitamina
            Me.ComboBox11.Text = "Termička obrada"   'termicka obrada
            Me.TextBox77.Text = ""  'cijena namirnice
        Else
            MsgBox("Prije izrade jelovnika izračunajte preporučeni energetski unos.")
            TabControl1.SelectedIndex = 0
            TabControl3.SelectedIndex = 0
        End If

        Exit Sub
Kraj:
        If MessageBox.Show("U Demo verziji broj namirnica u jelovniku je ograničen." _
                           & vbCrLf & "Želite li aktivirati punu verziju programa?", "Program Prehrane 5.0", _
                         MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                         = DialogResult.Yes Then
            System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
        End If


    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        On Error Resume Next
        '  Dim DGV As DataGridView = Me.DataGridView2
        '  Dim Index As Integer = DGV.CurrentRow.Index
        '   Dim GotovoJelo As String = DGV.Rows(Index).Cells(57).Value.ToString
        OdaberiNamirnicu()
        'Za lazanje nama podatke o namirnicama
        '      If DGV.Rows(Index).Cells(2).Value.ToString = "Lazanje" Then
        '  OdaberiNamirnicu()
        ' Exit Sub

        'End If
        'If GotovoJelo = "Jela" Then
        ''       If MessageBox.Show("Dali želite pojedinačne namirnice iz Jela uvrstiti u jelovnik?" _
        ''                 , "Program Prehrane 5.0", _
        ' '             MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
        '  '           = DialogResult.No Then
        '   'OdaberiNamirnicu()
        ' '  Exit Sub
        '' Else

        'GotovaJelaForm.Show()
        'Exit Sub
        ''End If
        ' End If
        '' OdaberiNamirnicu()

    End Sub

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView2.KeyUp
        OdaberiNamirnicu()

    End Sub

    Private Sub TabPage3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage3.Click

    End Sub



    Private Sub DataGridView3_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        DodatnaTjelesnaAktivnost()

    End Sub

    Private Sub DataGridView3_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView3.KeyUp
        ' DodatnaTjelesnaAktivnost()

    End Sub

    Private Sub TabPage6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage6.Click

    End Sub

    Private Sub TabPage6_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage6.Enter
        On Error Resume Next
        AktivnostiBindingSource()
        TextBox9.Text = ComboBox3.Text  'masa
        TextBox10.Text = TextBox4.Text  'dodatna energetska potrošnja
        Me.DataGridView3.CurrentRow.Selected = False

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OdabirAktivnosti()

    End Sub

    Private Sub TabPage5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage5.Click

    End Sub

    Private Sub TabPage5_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage5.Enter
        On Error Resume Next
        Parametri()
        ParametriPreporuceneVrijednosti()
        ParametriProgressBar()
        NParametri()
        With Me
            'graf - pita
            If .ListBox10.Items(0) = "0 kcal" Then
                .Chart3.Series(0).Points.Clear()
                .Chart3.Series(0).Points.AddY(20)
                .Chart3.Series(0).Points.AddY(30)
                .Chart3.Series(0).Points.AddY(50)
            End If
        End With
        If Val(ComboBox1.Text) < 9 Then
            MsgBox("Program ne daje podatke za osobe mlađe od 9 godina.")
            Exit Sub
        End If
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label112_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label112.Click

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        BrisiNamirnicu()
        UkupneVrijednosti()
        ObrociNutrijentiUkupno()
        With Me
            'graf - pita
            If .ListBox10.Items(0) = "0 kcal" Then
                .Chart2.Series(0).Points.Clear()
                .Chart2.Series(0).Points.AddY(20)
                .Chart2.Series(0).Points.AddY(30)
                .Chart2.Series(0).Points.AddY(50)
            End If
        End With

    End Sub

    Private Sub RadioButton15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton15.CheckedChanged
        'smanjenje tjelesna mase
        Dim Dob As Double = Val(Me.ComboBox1.Text)

        If RadioButton15.Checked = True Then
            'djeca
            If Dob >= 9 And Dob < 18 Then
                TextBox3.Text = Val(Label13.Text)
                TextBox4.Text = 200
            Else
                'punoljetni
                TextBox3.Text = Val(Label13.Text) - Val(Label197.Text)
                TextBox4.Text = Val(Label196.Text)
            End If
        End If
    End Sub

    Private Sub RadioButton16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton16.CheckedChanged
        'povecanje tjelesna mase
        If RadioButton16.Checked = True Then
            TextBox3.Text = Val(Label13.Text) + Val(Label197.Text) + Val(Label196.Text)
            TextBox4.Text = Val(Label196.Text)
        End If
    End Sub

    Private Sub RadioButton17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton17.CheckedChanged
        'zadrzavanje postojece tjelesna mase
        If RadioButton17.Checked = True Then
            TextBox3.Text = Val(Label13.Text) + Val(Label196.Text)
            TextBox4.Text = Val(Label196.Text)
        End If
    End Sub

    Private Sub RadioButton18_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton18.CheckedChanged
        'povecanje misicne mase
        If RadioButton18.Checked = True Then
            TextBox3.Text = Val(Label13.Text) + Val(Label197.Text) + Val(Label196.Text) + 200
            TextBox4.Text = Val(Label196.Text) + 200
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        On Error Resume Next

        If MessageBox.Show("Spremi jelovnik u BAZU JELOVNIKA?" _
                       , "Program Prehrane 5.0", _
                      MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                      = DialogResult.No Then
            Exit Sub
        Else

            Me.RadioButton20.Checked = True
            TabControl1.SelectedIndex = 8
            'Provjera dali vec postoji isti jelovnik u bazi
            Dim DGV As DataGridView = Me.DataGridView7
            Dim i As Integer
            For i = 0 To DGV.RowCount - 1
                If DGV.Rows(i).Cells(1).Value IsNot DBNull.Value _
                  And DGV.Rows(i).Cells(6).Value IsNot DBNull.Value _
                  And DGV.Rows(i).Cells(7).Value IsNot DBNull.Value Then
                    If DGV.Rows(i).Cells(1).Value = Me.TextBox1.Text & " " & Me.TextBox2.Text _
                         And DGV.Rows(i).Cells(6).Value = Me.TextBox13.Text _
                          Then
                        'And DGV.Rows(i).Cells(7).Value = Me.Label175.Text
                        '        MsgBox("Jelovnik već postoji u bazi jeovnika." _
                        '              & vbCrLf & "(" & Me.TextBox1.Text & " " & Me.TextBox2.Text _
                        '            & "/" & Me.TextBox13.Text & "/" & Me.Label175.Text & "kcal)" _
                        '            & vbCrLf & "Promijenite naziv jelovnika.")
                        ' DGV.Rows(i).Selected = True
                        DGV.CurrentCell = DGV.Rows(i).Cells(0)
                        If MessageBox.Show("U BAZI JELOVNIKA već postoji jelovnik s istim nazivom." _
                              & vbCrLf & "Spremi izmjene?", "Program Prehrane 5.0", _
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                            = DialogResult.No Then
                            TabControl1.SelectedIndex = 5
                            Exit Sub
                        Else
                            With Me
                                .RadioButton20.Checked = True   'Moji jelovnici
                                Dim DGV1 As DataGridView = .DataGridView8
                                Dim a As Integer
                                For a = 0 To DGV1.RowCount - 1
                                    If DGV1.Rows(a).Cells(1).Value IsNot DBNull.Value _
                                   And DGV1.Rows(a).Cells(6).Value IsNot DBNull.Value Then
                                        If DGV1.Rows(a).Cells(1).Value = .Label181.Text _
                                            And DGV1.Rows(a).Cells(6).Value = .Label186.Text Then _
                                          '   And DGV1.Rows(a).Cells(7).Value = .Label187.Text 
                                            DGV1.Rows.Remove(DGV1.CurrentRow)
                                        End If
                                    End If
                                Next a

                                .BazaJelovnikaBindingSource.MoveLast()
                                DGV.Rows.Remove(DGV.CurrentRow)
                                .BazaNazivaJelovnikaBindingSource.MoveLast()
                            End With


                            BazaNazivaJelovnikaSpremi()
                            BazaJelovnikaSpremi()
                            TabControl1.SelectedIndex = 5
                            MsgBox("Jelovnik je spremljen u BAZU JELOVNIKA.")
                        End If


                        TabControl1.SelectedIndex = 5
                        Me.TextBox13.Select()
                        Exit Sub
                    End If
                End If
            Next i
            BazaNazivaJelovnikaSpremi()
            BazaJelovnikaSpremi()
            TabControl1.SelectedIndex = 5
            MsgBox("Jelovnik je spremljen u BAZU JELOVNIKA.")

        End If

    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DataGridView7_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellClick

        On Error Resume Next
        Dim DGV As DataGridView = Me.DataGridView7
        'start
        If My.Settings.PP5StartAktivacija = "Da" Then
            Start()
            Dim i As Integer = DGV.CurrentRow.Index
            If i > 0 Then
                DGV.Rows(0).Selected = True   'vrati na prvi jelovnik 
                '   .VrstaDijeteBindingSource.Position = 2
                BazaNazivaPrimjeraJelovnikaBindingSource.Position = 0
                Exit Sub
            End If
        End If
        'demo
        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" Then
            Demo()
            Dim j As Integer = DGV.CurrentRow.Index
            If j > 0 Then
                DGV.Rows(0).Selected = True   'vrati na prvi jelovnik 
                BazaNazivaPrimjeraJelovnikaBindingSource.Position = 0
                Exit Sub
            End If
        End If

        BazaNazivaJelovnikaFilter()

    End Sub

    Private Sub DataGridView7_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

    End Sub

    Private Sub TabPage16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage16.Click

    End Sub

    Private Sub TabPage18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage18.Click

    End Sub

    Private Sub TabPage18_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage18.Enter
        PracenjeStanja()
        PracenjeStanjaFilter()
        PracenjeStanjaGraf()
        Me.DateTimePicker2.Value = Today
        TabControl4.SelectedIndex = 0

    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        PracenjeStanjaDodaj()
        PracenjeStanjaGraf()

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click

        With Me
            If .TextBox1.Text = "" Or .ComboBox3.Text = "" Then
                MsgBox("Polja NAZIV/IME i MASA su obavezni.")
                Exit Sub
            End If

        End With
        TabControl3.SelectedIndex = 3   'Baza klijenata
        SpremiKorisnika()
        TabControl3.SelectedIndex = 0

        TabControl3.SelectedIndex = 2
        TabControl4.SelectedIndex = 0   'Pracenje antropometrijskih parametara
        PracenjeStanjaDodaj()
        TabControl3.SelectedIndex = 0

    End Sub

    Private Sub ListBox32_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox32.SelectedIndexChanged

    End Sub

    Private Sub Label87_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label87.Click

    End Sub

    Private Sub Label87_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label87.TextChanged

    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged
        'Djeca
        '        If Val(ComboBox1.Text) >= 9 And Val(ComboBox1.Text) < 18 Then
        'GroupBox2.Text = "Intenzitet tjelesne aktivnosti"
        '      GroupBox3.Visible = False
        '     Me.Label188.Visible = False
        '    Me.ComboBox6.Visible = False
        '   Me.Label190.Visible = False
        '  Me.Label189.Visible = False
        ' Me.ComboBox7.Visible = False
        '      Me.Label191.Visible = False
        '     Me.PictureBox9.Visible = False
        '    Exit Sub
        '   Else
        '  GroupBox2.Text = "Intenzitet tjelesne aktivnosti na poslu"
        '      GroupBox3.Visible = True
        '     Me.Label188.Visible = True
        '    Me.ComboBox6.Visible = True
        '   Me.Label190.Visible = True
        '  Me.Label189.Visible = True
        ' Me.ComboBox7.Visible = True
        '     Me.Label191.Visible = True
        '    Me.PictureBox9.Visible = True
        '   End If

    End Sub

    Private Sub TextBox8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox8.Click
        TextBox8.Text = ""

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        On Error Resume Next
        If TextBox8.Text = "Pretraži" Then Exit Sub
        '        If RadioButton11.Checked = True Then
        If TextBox8.Text = "" Then SportskeAktivnostiBindingSource.RemoveFilter()
        SportskeAktivnostiBindingSource.RemoveFilter()
        SportskeAktivnostiBindingSource.Filter = "OpisTjelesneAktivnosti Like'%" & TextBox8.Text & "%'"
        DataGridView1.CurrentRow.Selected = False
        ' End If

        '     If RadioButton12.Checked = True Then
        'If TextBox8.Text = "" Then SportskeAktivnostiBindingSource.RemoveFilter()
        '        SportskeAktivnostiBindingSource.RemoveFilter()
        '       SportskeAktivnostiBindingSource.Filter = "OpisTjelesneAktivnosti Like'%" & TextBox8.Text & "%'"
        '      DataGridView1.CurrentRow.Selected = False
        '      End If
    End Sub

    Private Sub RadioButton11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '       Label20.Visible = True
        '      Label95.Visible = True
        '     Label96.Visible = True

        'If RadioButton11.Checked = True Then
        'DataGridView3.DataSource = SportskeAktivnostiBindingSource
        '        Label20.DataBindings.Clear()
        '       Label96.DataBindings.Clear()
        '      Label94.DataBindings.Clear()
        '     Label20.DataBindings.Add(New Binding("Text", SportskeAktivnostiBindingSource, "OpisTjelesneAktivnosti"))
        '    ' Label96.DataBindings.Add(New Binding("Text", TjelesneAktivnostiBindingSource, "FaktorTjelesneAktivnostikJ"))
        '  Label94.DataBindings.Add(New Binding("Text", SportskeAktivnostiBindingSource, "FaktorTjelesneAktivnostiKcal"))
        '   End If

        '  If RadioButton12.Checked = True Then
        '     DataGridView3.DataSource = SportskeAktivnostiBindingSource
        '    Label20.DataBindings.Clear()
        '   Label96.DataBindings.Clear()
        '  Label94.DataBindings.Clear()
        ' Label20.DataBindings.Add(New Binding("Text", SportskeAktivnostiBindingSource, "OpisTjelesneAktivnosti"))
        ' Label96.DataBindings.Add(New Binding("Text", SportskeAktivnostiBindingSource, "FaktorTjelesneAktivnostikJ"))
        '      Label94.DataBindings.Add(New Binding("Text", SportskeAktivnostiBindingSource, "FaktorTjelesneAktivnostiKcal"))
        '     End If

        '   Label20.Visible = False
        '  Label95.Visible = False
        ' Label96.Visible = False

    End Sub

    Private Sub RadioButton12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub RadioButton13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton13.CheckedChanged
        If RadioButton13.Checked = True Then
            TextBox85.Enabled = True
            TextBox6.Enabled = False
            TextBox12.Enabled = False
        End If
    End Sub

    Private Sub RadioButton14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton14.CheckedChanged
        If RadioButton14.Checked = True Then
            TextBox85.Enabled = False
            TextBox6.Enabled = True
            TextBox12.Enabled = False
        End If
    End Sub

    Private Sub RadioButton19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton19.CheckedChanged
        If RadioButton19.Checked = True Then
            TextBox85.Enabled = False
            TextBox6.Enabled = False
            TextBox12.Enabled = True
        End If
    End Sub

    Private Sub ComboBox4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)


    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub ComboBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox6_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        'samo brojevi
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "," Then
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox6.TextChanged
        On Error Resume Next
        '  If TextBox6.Text < 0.00001 Then Exit Sub
        Dim Serv As Double = (TextBox6.Text / Label228.Text) * Label363.Text
        Dim Kol As Double = (TextBox6.Text / Label228.Text) * Label229.Text
        If RadioButton14.Checked = True Then
            If Serv < 0.1 Then
                TextBox85.Text = Format(Serv, "0.00")   'Serviranja
                TextBox12.Text = Format(Kol, "0.00")   'Kolicina
            Else
                TextBox85.Text = Format(Serv, "0.0")   'Serviranja
                TextBox12.Text = Format(Kol, "0.00")   'Kolicina
            End If
            '   TextBox85.Text = Format((TextBox6.Text / Label228.Text) * Label363.Text, "0.0")   'Serviranja
            '  TextBox12.Text = Format((TextBox6.Text / Label228.Text) * Label229.Text, "0.00")   'Kolicina
        End If

        Mjera()

    End Sub

    Private Sub TextBox12_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        'samo brojevi
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "," Then
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        On Error Resume Next
        If TextBox12.Text < 0.00001 Then Exit Sub

        Dim Serv As Double = (TextBox12.Text / Label229.Text) * Label363.Text
        Dim Mas As Double = (TextBox12.Text / Label229.Text) * Label228.Text
        If RadioButton19.Checked = True Then
            If Serv < 0.1 Then
                TextBox85.Text = Format(Serv, "0.00")   'Serviranja
                TextBox6.Text = Format(Mas, "0.0")   'Masa
            Else
                TextBox85.Text = Format(Serv, "0.0")   'Serviranja
                TextBox6.Text = Format(Mas, "0.0")   'Masa
            End If
        End If

        '   If RadioButton19.Checked = True Then
        'TextBox85.Text = Format((TextBox12.Text / Label229.Text) * Label363.Text, "0.0")   'Serviranja
        '  TextBox6.Text = Format((TextBox12.Text / Label229.Text) * Label228.Text, "0")   'Masa
        '    End If

        Mjera()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        On Error Resume Next

        'start
        If My.Settings.PP5StartAktivacija = "Da" Or My.Settings.PP5StartTrajnaLicencaAktivacija = "Da" Then
            Start()
        End If
        'demo
        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then
            Demo()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        On Error Resume Next
        DataGridView4.Rows.Remove(DataGridView4.CurrentRow)
        OdabranaDodatnaTjelesnaAktivnostBindingSource.MoveLast()

        Dim i As Integer
        Dim DGV As DataGridView
        DGV = Me.DataGridView4
        Dim Energ As Double = 0
        For i = 0 To DGV.RowCount - 1
            Energ = Energ + DGV.Rows(i).Cells(7).Value   'ukupna dodatna potrosnja
        Next i
        Me.Label176.Text = Energ
        Me.Label301.Text = "Ukupno: " & Energ & " kcal"   'ukupna dodatna energetska potrosnja

        DodatnaTjelesnaAktivnost()

    End Sub

    Private Sub Label96_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label96.Click

    End Sub

    Private Sub DataGridView7_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView7.CellDoubleClick
        UzmiBazaJelovnika()

    End Sub

    Private Sub DataGridView7_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView7.KeyUp
        On Error Resume Next
        Dim DGV As DataGridView = Me.DataGridView7
        'start
        If My.Settings.PP5StartAktivacija = "Da" Then
            Start()
            Dim i As Integer = DGV.CurrentRow.Index
            If i > 0 Then
                DGV.Rows(0).Selected = True   'vrati na prvi jelovnik 
                BazaNazivaPrimjeraJelovnikaBindingSource.Position = 0
                Exit Sub
            End If
        End If
        'demo
        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" Then
            Demo()
            Dim j As Integer = DGV.CurrentRow.Index
            If j > 0 Then
                DGV.Rows(0).Selected = True   'vrati na prvi jelovnik 
                BazaNazivaPrimjeraJelovnikaBindingSource.Position = 0
                Exit Sub
            End If
        End If

        BazaNazivaJelovnikaFilter()

    End Sub

    Private Sub Button18_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        On Error Resume Next
        Me.SveNamirniceBindingSource.RemoveFilter()
        If Me.DataGridView2.CurrentRow IsNot Nothing Then
            Me.DataGridView2.CurrentRow.Selected = False
        End If

        SkupineNamirnica()
        OdaberiNamirnicu()

        '       'BOJA
        '      Dim DGV As DataGridView = Me.DataGridView2
        '     Dim i As Integer = DGV.CurrentRow.Index
        '    'Namirnice
        '   For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.BackColor = My.Settings.DgvBoja
        '     Next
        '    'BOJA-SELEKTIRANI RED
        '   'Sve namirnice
        '  For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = My.Settings.DgvBojaSelect
        '   Next

    End Sub

    Private Sub Label258_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label258.Click

    End Sub

    Private Sub TextBox7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox7.Click
        TextBox7.Text = ""

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        On Error Resume Next
        If TextBox7.Text = "Pretraži" Then Exit Sub
        '  Label15.Text = "Sve namirnice"
        If Me.Label15.Text <> "Moje namirnice" And Me.Label15.Text <> "Favoriti" Then
            DataGridView2.DataSource = SveNamirniceBindingSource
            SveNamirniceBindingSource.RemoveFilter()
            SveNamirniceBindingSource.Filter = "NazivNamirnice Like'%" & TextBox7.Text & "%'"
        End If
        If Me.Label15.Text = "Moje namirnice" Then
            DataGridView2.DataSource = MojeNamirniceBindingSource
            MojeNamirniceBindingSource.RemoveFilter()
            MojeNamirniceBindingSource.Filter = "NazivNamirnice Like'%" & TextBox7.Text & "%'"
        End If
        If Me.Label15.Text = "Favoriti" Then
            DataGridView2.DataSource = FavoritiBindingSource
            FavoritiBindingSource.RemoveFilter()
            FavoritiBindingSource.Filter = "NazivNamirnice Like'%" & TextBox7.Text & "%'"
        End If


        DataGridView2.CurrentRow.Selected = False

        '      'BOJA
        '     Dim DGV As DataGridView = Me.DataGridView2
        '    Dim i As Integer = DGV.CurrentRow.Index
        '   'Namirnice
        '  For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.BackColor = My.Settings.DgvBoja
        '     Next
        '    'BOJA-SELEKTIRANI RED
        '   'Sve namirnice
        '  For i = 0 To DGV.Rows.Count - 1 Step 2
        'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = My.Settings.DgvBojaSelect
        '  Next

    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub ListBox10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox10.SelectedIndexChanged

    End Sub

    Private Sub TabPage1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub TabPage1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage1.Enter
        TabControl3.SelectedIndex = 0

    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        MojeNamirniceSpremi()

    End Sub

    Private Sub TabPage17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage17.Click

    End Sub

    Private Sub TabPage17_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage17.Enter
        On Error Resume Next
        Dim BS As BindingSource = Me.MojeNamirniceBindingSource
        Me.DataGridView2.DataSource = BS
        If Me.DataGridView2.RowCount <= 1 Then
            BS.AddNew()
            BS.MovePrevious()
            DataGridView2.CurrentRow.Selected = False
        End If
        BS.MoveLast()
        Me.Label15.Text = "Moje namirnice"

    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TabPage3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage3.Enter

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        NoviJelovnik()
        NoviJelovnik()
        UkupneVrijednosti()
        ObrociNutrijentiUkupno()
        With Me
            'graf - pita
            .Chart2.Series(0).Points.Clear()
            .Chart2.Series(0).Points.AddY(20)
            .Chart2.Series(0).Points.AddY(30)
            .Chart2.Series(0).Points.AddY(50)
        End With
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Button28_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        'Clipboard.SetText(Me.TextBox69.Text, TextDataFormat.Text)
        Clipboard.SetText(Me.RichTextBoxPrintCtrl1.Rtf, TextDataFormat.Rtf)

    End Sub

    Private Sub ListBox60_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox60.SelectedIndexChanged

    End Sub

    Private Sub GroupBox35_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox35.Enter

    End Sub

    Private Sub GroupBox45_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox45.Enter

    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        NoviKorisnik()
        BazaKorisnikaUzmi()

    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        BazaKorisnikaBrisi()

    End Sub

    Private Sub TabPage15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage15.Click

    End Sub

    Private Sub TabPage15_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage15.Enter
        On Error Resume Next
        Me.DataGridView6.CurrentRow.Selected = False

    End Sub

    Private Sub Button24_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        On Error Resume Next

        If Me.DataGridView7.CurrentRow.Selected = False Then
            MsgBox("Odaberite jelovnik.")
            Exit Sub
        End If

        ' PrebaciNamirnicu()
        BazaNazivaJelovnikaFilter()
        BazaJelovnikaBindingClear()
        BazaJelovnika_BindingSource()
        UzmiBazaJelovnika()

    End Sub

    Private Sub TabPage7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage7.Click

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs)

    End Sub

    Private Sub MenuStrip1_ItemClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub OProgramuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OProgramuToolStripMenuItem.Click
        On Error Resume Next
        OProgramuForma.Close()
        OProgramuForma.Show()

    End Sub

    Private Sub TabPage7_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage7.Enter

    End Sub

    Private Sub TabPage8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage8.Click

    End Sub

    Private Sub TabPage8_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage8.Enter

    End Sub
    Dim StringToPrint As String

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        With Me
            '   StringToPrint = .TextBox69.Text
            'StringToPrint = .RichTextBox1.Text
            StringToPrint = .RichTextBoxPrintCtrl1.Text

            .PrintDialog1.Document = .PrintDocument1
            .PageSetupDialog1.Document = .PrintDocument1
            If .PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If .PageSetupDialog1.ShowDialog = DialogResult.OK Then
                    '.PrintPreviewDialog1.Document = .PrintDocument1
                    '.PrintPreviewDialog1.ShowDialog()
                    PrintDocument1.Print()
                End If
            End If
        End With
    End Sub

    Private Sub Button18_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click

        TabControl1.SelectedIndex = 7

        '      Ispis()
        '     With Me
        ''StringToPrint = .TextBox69.Text
        ''StringToPrint = .RichTextBox1.Text
        '     StringToPrint = .RichTextBoxPrintCtrl1.Text

        '.PrintDialog1.Document = .PrintDocument1
        '   .PageSetupDialog1.Document = .PrintDocument1
        '  If .PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
        'If .PageSetupDialog1.ShowDialog = DialogResult.OK Then
        ''.PrintPreviewDialog1.Document = .PrintDocument1
        ''.PrintPreviewDialog1.ShowDialog()
        '    PrintDocument1.Print()
        '   End If
        '  End If
        ' End With
    End Sub

    Private Sub Button19_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        IspisOsobniPodaci()
        '    With Me
        'StringToPrint = .TextBox69.Text
        '     .PageSetupDialog1.Document = .PrintDocument1
        '    If .PageSetupDialog1.ShowDialog = DialogResult.OK Then
        '.PrintPreviewDialog1.Document = .PrintDocument1
        '    .PrintPreviewDialog1.ShowDialog()
        '   End If
        '  End With

        With Me
            '   StringToPrint = .TextBox69.Text
            'StringToPrint = .RichTextBox1.Text
            StringToPrint = .RichTextBoxPrintCtrl1.Text

            .PrintDialog1.Document = .PrintDocument1
            .PageSetupDialog1.Document = .PrintDocument1
            If .PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If .PageSetupDialog1.ShowDialog = DialogResult.OK Then
                    '.PrintPreviewDialog1.Document = .PrintDocument1
                    '.PrintPreviewDialog1.ShowDialog()
                    PrintDocument1.Print()
                End If
            End If
        End With
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        SavePrikazi()

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        '      Dim numChars As Integer
        '     Dim numLines As Integer
        '    Dim stringForPage As String
        '   Dim strFormat As New StringFormat()
        '  Dim PrintFont As Font
        ' PrintFont = TextBox69.Font
        ' PrintFont = RichTextBox1.Font
        '       Dim rectDraw As New RectangleF(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height)
        '      Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont.GetHeight(e.Graphics))
        '     strFormat.Trimming = StringTrimming.Word
        '    e.Graphics.MeasureString(StringToPrint, PrintFont, sizeMeasure, strFormat, numChars, numLines)
        '   stringForPage = StringToPrint.Substring(0, numChars)
        '  e.Graphics.DrawString(stringForPage, PrintFont, Brushes.Black, rectDraw, strFormat)
        ' If numChars < StringToPrint.Length Then
        'StringToPrint = StringToPrint.Substring(numChars)
        '      e.HasMorePages = True
        '     Else
        '    e.HasMorePages = False
        '   End If


        ' Print the content of the RichTextBox. Store the last character printed.
        Dim checkPrint As Integer
        checkPrint = RichTextBoxPrintCtrl1.Print(checkPrint, RichTextBoxPrintCtrl1.TextLength, e)

        ' Look for more pages
        If checkPrint < RichTextBoxPrintCtrl1.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If



    End Sub

    Private Sub BrišiNamirnicuIzSkupineMojeNamirniceToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GroupBox13_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox13.Enter

    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        NoviKorisnik()

        On Error Resume Next
        With Me

            Dim DGV As DataGridView = .DataGridView17
            Dim BS As BindingSource = BazaEnergetskePotrosnjeBindingSource
            Dim i As Integer

            'briši odabrane aktivnosti

            For i = 0 To DGV.RowCount - 1
                DGV.Rows.Remove(DGV.CurrentRow)
            Next i
            BS.AddNew()

            .DataGridView17.CurrentRow.Selected = False

            .TextBox83.Text = 0
            .TextBox84.Text = 0
            .ComboBox19.Text = 0
            .ComboBox20.Text = 0
            .Label315.Text = 0

        End With

        TrajanjeAktivnostiOdDo()
        EnergetskaPotrosnjaUkupno()

    End Sub

    Private Sub NoviKorisnikToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoviKorisnikToolStripMenuItem.Click
        NoviKorisnik()

        On Error Resume Next
        With Me

            Dim DGV As DataGridView = .DataGridView17
            Dim BS As BindingSource = BazaEnergetskePotrosnjeBindingSource
            Dim i As Integer

            'briši odabrane aktivnosti

            For i = 0 To DGV.RowCount - 1
                DGV.Rows.Remove(DGV.CurrentRow)
            Next i
            BS.AddNew()

            .DataGridView17.CurrentRow.Selected = False

            .TextBox83.Text = 0
            .TextBox84.Text = 0
            .ComboBox19.Text = 0
            .ComboBox20.Text = 0
            .Label315.Text = 0

        End With

        TrajanjeAktivnostiOdDo()
        EnergetskaPotrosnjaUkupno()

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        BazaJelovnikaBrisi()

    End Sub

    Private Sub UputaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UputaToolStripMenuItem.Click
        On Error Resume Next
        System.Diagnostics.Process.Start("http://www.programprehrane.com/4da35544-91f6-4206-9c33-47ad44379a3b/PP5Uputa.pdf")

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox9.SelectedIndexChanged
        PracenjeStanjaGraf()

    End Sub

    Private Sub DataGridView10_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView10.CellContentClick

    End Sub

    Private Sub Button23_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        PracenjeStanjaFilter()
        PracenjeStanjaGraf()

    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox15.TextChanged
        PracenjeStanjaFilter()
        PracenjeStanjaGraf()

    End Sub

    Private Sub PriručnikToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PriručnikToolStripMenuItem1.Click
        On Error Resume Next
        System.Diagnostics.Process.Start("http://www.programprehrane.com/4da35544-91f6-4206-9c33-47ad44379a3b/SamSvojNutricionist.pdf")

    End Sub

    Private Sub Button23_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        On Error Resume Next
        With Me
            Dim DGV As DataGridView = .DataGridView10
            Dim i As Integer = DGV.CurrentRow.Index
            If DGV.Rows(i).Cells(2).Value.ToString <> "" Then
                '      MsgBox("~" & DGV.Rows(i).Cells(2).Value.ToString & "~")
                DGV.Rows.Remove(DGV.CurrentRow)   'brisi unos - pracenje stanja
                .KorisniciPracenjeStanjaBindingSource.MoveLast()
                DGV.CurrentRow.Selected = False
            End If
        End With
        PracenjeStanjaGraf()

    End Sub

    Private Sub IzlazIzProgramaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IzlazIzProgramaToolStripMenuItem.Click
        On Error Resume Next
        With Me
            .BazaKorisnikaBindingSource.EndEdit()
            .BazaKorisnikaTableAdapter.Update(.PP5aDataSet)

            .KorisniciPracenjeStanjaBindingSource.EndEdit()
            .KorisniciPracenjeStanjaTableAdapter.Update(.PP5aDataSet)

            .BazaNazivaJelovnikaBindingSource.EndEdit()
            .BazaNazivaJelovnikaTableAdapter.Update(.PP5aDataSet)

            .BazaJelovnikaBindingSource.EndEdit()
            .BazaJelovnikaTableAdapter.Update(.PP5aDataSet)

            .MojeNamirniceBindingSource.EndEdit()
            .MojeNamirniceTableAdapter.Update(.PP5aDataSet)

            .FavoritiBindingSource.EndEdit()
            .FavoritiTableAdapter.Update(PP5aDataSet)

            .CijeneBindingSource.EndEdit()
            .CijeneTableAdapter.Update(.PP5aDataSet)

        End With
        End

    End Sub

    Private Sub AktivacijaPuneVerzijeProgramaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AktivacijaPuneVerzijeProgramaToolStripMenuItem.Click

        '  UlazPPForma.Show()

        '      Dim Kljuc = InputBox("Unesite aktivacijski ključ", MsgBoxStyle.Information)
        '     If Kljuc = "ERH4K297GT" Then   'Akticacijski kljuc
        'My.Settings.AktivacijskiKljuc = "Da"
        '        My.Settings.Save()
        '       Me.Text = "Program Prehrane 5.0 - Copyright (c) 2013, IG PROG."
        '      Me.AktivacijaPuneVerzijeProgramaToolStripMenuItem.Visible = False
        '     Me.PriručnikToolStripMenuItem1.Enabled = True
        '    Me.Button39.Enabled = True
        '   System.Diagnostics.Process.Start("http://www.programprehrane.com/pp4aktivacija.htm")
        '  Else
        ' '  If Kljuc <> "" Then
        'If MessageBox.Show("Pogrešan unos." _
        '                          & vbCrLf & "Dali želite naručiti aktivacijski ključ za pokretanje pune verzije Programa Prehrane 5.0?", "Program Prehrane 5.0 Demo", _
        '                       MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
        '                      = DialogResult.Yes Then
        '     System.Diagnostics.Process.Start("http://www.programprehrane.com/pp4kljuc.htm")
        '    End If
        'End If
        '   End If
    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub TabPage14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage14.Click

    End Sub

    Private Sub TabPage14_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage14.Enter

    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        On Error Resume Next
        PrimjeriAktivnostiNaPosluForma.Close()
        PrimjeriAktivnostiNaPosluForma.Show()

    End Sub

    Private Sub PictureBox9_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox9.MouseMove

    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub TabPage13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage13.Click

    End Sub

    Private Sub TabPage13_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage13.Enter
        '  Me.TextBox69.Text = ""   'ispis briši
        Ispis()
        Ispis()
        Me.ComboBox17.SelectedIndex = 0  'Ispis Jelovnika

    End Sub

    Private Sub Label15_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.TextChanged

        If Me.Label15.Text = "Moje namirnice" Or Me.Label15.Text = "Favoriti" Then
            Me.DataGridView2.AllowUserToDeleteRows = True
            Me.DataGridView2.ReadOnly = False
        Else
            Me.DataGridView2.AllowUserToDeleteRows = False
            Me.DataGridView2.ReadOnly = True
        End If


    End Sub

    Private Sub TabPage21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabPage21.Click

    End Sub

    Private Sub TabPage21_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage21.Enter
        Me.Label23.Text = Me.TextBox14.Text & " " & Me.TextBox15.Text

    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        On Error Resume Next
        PricekajteTrenutakDemoForma.Close()
        PricekajteTrenutakStartForma.Close()
        PricekajteTrenutakPremiumForma.Close()

    End Sub

    Private Sub RadioButton20_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton20.CheckedChanged
        On Error Resume Next
        With Me
            If .RadioButton20.Checked = True Then
                DataGridView7.DataSource = BazaNazivaJelovnikaBindingSource
                DataGridView8.DataSource = BazaJelovnikaBindingSource
                .DataGridView7.CurrentRow.Selected = False
                .DataGridView8.CurrentRow.Selected = False
            End If
        End With
        BazaJelovnikaBindingClear()
        BazaJelovnika_BindingSource()


    End Sub

    Private Sub RadioButton21_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton21.CheckedChanged
        On Error Resume Next
        With Me
            If .RadioButton21.Checked = True Then
                DataGridView7.DataSource = BazaNazivaPrimjeraJelovnikaBindingSource
                DataGridView8.DataSource = BazaPrimjeraJelovnikaBindingSource
                .DataGridView7.CurrentRow.Selected = False
                .DataGridView8.CurrentRow.Selected = False
            End If
        End With
        BazaJelovnikaBindingClear()
        BazaJelovnika_BindingSource()

    End Sub

    Private Sub Button31_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox75_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox75.TextChanged
        If Len(Me.TextBox75.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If

    End Sub

    Private Sub TextBox11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox11.TextChanged
        If Len(Me.TextBox11.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If
    End Sub

    Private Sub TextBox64_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox64.TextChanged
        If Len(Me.TextBox64.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If
    End Sub

    Private Sub TextBox65_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox65.TextChanged
        If Len(Me.TextBox65.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If

    End Sub

    Private Sub TextBox66_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox66.TextChanged
        If Len(Me.TextBox66.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If

    End Sub

    Private Sub TextBox67_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox67.TextChanged
        If Len(Me.TextBox67.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If

    End Sub

    Private Sub TextBox68_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox68.TextChanged
        If Len(Me.TextBox68.Text) = 1000 Then
            MsgBox("Maksimalni broj znakova je 1000.")
        End If

    End Sub

    Private Sub TabPage16_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabPage16.Enter
        On Error Resume Next
        With Me
            If .RadioButton20.Checked = True Then
                DataGridView7.DataSource = BazaNazivaJelovnikaBindingSource
                DataGridView8.DataSource = BazaJelovnikaBindingSource
                .DataGridView7.CurrentRow.Selected = False
                .DataGridView8.CurrentRow.Selected = False
            End If
        End With
        BazaJelovnikaBindingClear()
        BazaJelovnika_BindingSource()

    End Sub

    Private Sub Button31_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        On Error Resume Next
        With Me
            Dim Korak As Double = 1
            If .ComboBox9.Text = "WHR (omjer opsega struka i bokova)" Then
                '   Korak = 0.1
                '.Chart1.ChartAreas(0).AxisY.IntervalOffset = 0.1
                '    If .Chart1.ChartAreas(0).AxisY.Minimum < 0.5 Then
                '.Chart1.ChartAreas(0).AxisY.Minimum = 0.5
                Exit Sub
                '    End If
            Else
                Korak = 1
                '   .Chart1.ChartAreas(0).AxisY.IntervalOffset = 1
                If .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Minimum = 5
                    Exit Sub
                End If
                If .Chart1.ChartAreas(0).AxisY.Maximum - .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Maximum = .Chart1.ChartAreas(0).AxisY.Maximum + Korak
                    .Chart1.ChartAreas(0).AxisY.Minimum = .Chart1.ChartAreas(0).AxisY.Minimum - Korak
                    Exit Sub
                End If
            End If

            Dim Min As Integer = .Chart1.ChartAreas(0).AxisY.Minimum - Korak
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisY.Maximum + Korak
            .Chart1.ChartAreas(0).AxisY.Minimum = Min
            .Chart1.ChartAreas(0).AxisY.Maximum = Max
        End With


    End Sub

    Private Sub Button32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button32.Click
        On Error Resume Next
        With Me
            Dim Korak As Double = 0.1
            If .ComboBox9.Text = "WHR (omjer opsega struka i bokova)" Then
                '      Korak = 0.1
                '  .Chart1.ChartAreas(0).AxisY.IntervalOffset = 0.1
                '         If .Chart1.ChartAreas(0).AxisY.Minimum < 0.5 Then
                '   .Chart1.ChartAreas(0).AxisY.Minimum = 0.5
                Exit Sub
                '  End If
            Else
                Korak = 1
                '   .Chart1.ChartAreas(0).AxisY.IntervalOffset = 1
                If .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Minimum = 5
                    Exit Sub
                End If
                If .Chart1.ChartAreas(0).AxisY.Maximum - .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Maximum = .Chart1.ChartAreas(0).AxisY.Maximum + 1
                    .Chart1.ChartAreas(0).AxisY.Minimum = .Chart1.ChartAreas(0).AxisY.Minimum - 1
                    Exit Sub
                End If
            End If

            Dim Min As Integer = .Chart1.ChartAreas(0).AxisY.Minimum + Korak
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisY.Maximum - Korak
            .Chart1.ChartAreas(0).AxisY.Minimum = Min
            .Chart1.ChartAreas(0).AxisY.Maximum = Max
        End With


    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        On Error Resume Next
        With Me
            If .Chart1.ChartAreas(0).AxisX.Minimum < 0 Or .Chart1.ChartAreas(0).AxisX.Maximum - .Chart1.ChartAreas(0).AxisX.Minimum < 2 Then
                Exit Sub
            End If
            Dim Min As Integer = .Chart1.ChartAreas(0).AxisX.Minimum - 0.5
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisX.Maximum + 0.5
            .Chart1.ChartAreas(0).AxisX.Minimum = Min
            .Chart1.ChartAreas(0).AxisX.Maximum = Max
        End With

    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        On Error Resume Next
        With Me
            If .Chart1.ChartAreas(0).AxisX.Minimum < 0 Or .Chart1.ChartAreas(0).AxisX.Maximum - .Chart1.ChartAreas(0).AxisX.Minimum < 2 Then
                Exit Sub
            End If
            Dim Min As Integer = .Chart1.ChartAreas(0).AxisX.Minimum + 0.5
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisX.Maximum - 0.5
            .Chart1.ChartAreas(0).AxisX.Minimum = Min
            .Chart1.ChartAreas(0).AxisX.Maximum = Max
        End With

    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        On Error Resume Next
        With Me
            Dim Korak As Double = 1
            If .ComboBox9.Text = "WHR (omjer opsega struka i bokova)" Then
                '  Korak = 0.1
                '    .Chart1.ChartAreas(0).AxisY.IntervalOffset = 0.1
                '   If .Chart1.ChartAreas(0).AxisY.Minimum < 0.5 Then
                '.Chart1.ChartAreas(0).AxisY.Minimum = 0.5
                Exit Sub
                ' End If
            Else
                Korak = 1
                '    .Chart1.ChartAreas(0).AxisY.IntervalOffset = 1
                If .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Minimum = 5
                    Exit Sub
                End If
            End If

            Dim Min As Integer = .Chart1.ChartAreas(0).AxisY.Minimum + Korak
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisY.Maximum + Korak
            .Chart1.ChartAreas(0).AxisY.Minimum = Min
            .Chart1.ChartAreas(0).AxisY.Maximum = Max
        End With

    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click
        On Error Resume Next
        With Me
            Dim Korak As Double = 1
            If .ComboBox9.Text = "WHR (omjer opsega struka i bokova)" Then
                '    Korak = 0.1
                '   .Chart1.ChartAreas(0).AxisY.Interval = 0.1
                '   If .Chart1.ChartAreas(0).AxisY.Minimum < 0.5 Then
                '.Chart1.ChartAreas(0).AxisY.Minimum = 0.5
                Exit Sub
                'End If
            Else
                Korak = 1
                '   .Chart1.ChartAreas(0).AxisY.Interval = 1
                If .Chart1.ChartAreas(0).AxisY.Minimum < 5 Then
                    .Chart1.ChartAreas(0).AxisY.Minimum = 5
                    Exit Sub
                End If
            End If

            Dim Min As Integer = .Chart1.ChartAreas(0).AxisY.Minimum - Korak
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisY.Maximum - Korak
            .Chart1.ChartAreas(0).AxisY.Minimum = Min
            .Chart1.ChartAreas(0).AxisY.Maximum = Max
        End With

    End Sub

    Private Sub Button37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button37.Click
        On Error Resume Next
        With Me
            Dim Min As Integer = .Chart1.ChartAreas(0).AxisX.Minimum - 1
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisX.Maximum - 1
            .Chart1.ChartAreas(0).AxisX.Minimum = Min
            .Chart1.ChartAreas(0).AxisX.Maximum = Max
        End With

    End Sub

    Private Sub Button38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button38.Click
        On Error Resume Next
        With Me
            Dim Min As Integer = .Chart1.ChartAreas(0).AxisX.Minimum + 1
            Dim Max As Integer = .Chart1.ChartAreas(0).AxisX.Maximum + 1
            .Chart1.ChartAreas(0).AxisX.Minimum = Min
            .Chart1.ChartAreas(0).AxisX.Maximum = Max
        End With

    End Sub

    Private Sub DataGridView6_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView6.CellDoubleClick
        BazaKorisnikaUzmi()

    End Sub

    Private Sub Button39_Click(sender As System.Object, e As System.EventArgs) Handles Button39.Click
        On Error Resume Next
        System.Diagnostics.Process.Start("http://www.programprehrane.com/4da35544-91f6-4206-9c33-47ad44379a3b/SamSvojNutricionist.pdf")

    End Sub

    Private Sub Button40_Click(sender As System.Object, e As System.EventArgs) Handles Button40.Click
        On Error Resume Next
        System.Diagnostics.Process.Start("http://www.programprehrane.com/4da35544-91f6-4206-9c33-47ad44379a3b/PP5Uputa.pdf")


    End Sub

    Private Sub PrintPreviewDialog1_Load(sender As System.Object, e As System.EventArgs) Handles PrintPreviewDialog1.Load

    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk

    End Sub

    Private Sub UpdateToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles UpdateToolStripMenuItem1.Click
        System.Diagnostics.Process.Start("http://www.programprehrane.com/download/hr/ProgramPrehrane5.exe")

    End Sub

    Private Sub Button41_Click(sender As System.Object, e As System.EventArgs) Handles Button41.Click
        On Error Resume Next
        SortiranjeIspisaForm.Close()
        SortiranjeIspisaForm.Show()

    End Sub

    Private Sub Button42_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button42_Click_1(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub Button42_Click_2(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TextBox13_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TabPage22_Click(sender As System.Object, e As System.EventArgs) Handles TabPage22.Click

    End Sub

    Private Sub TabPage22_Enter(sender As Object, e As System.EventArgs) Handles TabPage22.Enter
        Me.TextBox78.Text = Me.TextBox5.Text   'naziv namirnice

        DataGridView2.DataSource = SveNamirniceBindingSource
        SveNamirniceBindingSource.RemoveFilter()

    End Sub

    Private Sub Button43_Click(sender As System.Object, e As System.EventArgs) Handles Button43.Click
        With Me
            If .ComboBox12.SelectedValue = "Jela" Then
                MsgBox("Nije moguće spremiti cijenu pripremljenih jela.")
                Exit Sub
            End If
            If .CheckBox1.Checked = False Then
                .CheckBox1.Checked = True
                TabControl1.SelectedIndex = 5
                .Label331.Text = .Label325.Text   'jedinicna cijena
                FavoritiSpremi()
                .CheckBox1.Checked = False
            Else
                TabControl1.SelectedIndex = 5
                .Label331.Text = .Label325.Text   'jedinicna cijena
                FavoritiSpremi()
            End If
            .TabControl1.SelectedIndex = 10
        End With
    End Sub

    Private Sub TextBox79_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox79.TextChanged
        On Error Resume Next
        Dim Cijena As Double = Me.TextBox79.Text
        Dim Kolicina As Double = Me.TextBox80.Text
        Dim Valuta As String = Me.ComboBox26.Text
        Dim JedinicnaCijena As Double = Format(Cijena / (Kolicina / 1000), "0.00")
        Me.Label324.Text = "Jedinična cijena: " & JedinicnaCijena & " " & Valuta & " / 1 kg"
        Me.Label325.Text = JedinicnaCijena
        Me.TextBox77.Text = JedinicnaCijena
        Me.TextBox81.Text = Me.TextBox80.Text
        Me.Label324.Visible = True

    End Sub

    Private Sub TextBox80_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox80.TextChanged
        On Error Resume Next
        Dim Cijena As Double = Me.TextBox79.Text
        Dim Kolicina As Double = Me.TextBox80.Text
        Dim Valuta As String = Me.ComboBox26.Text
        Dim JedinicnaCijena As Double = Format(Cijena / (Kolicina / 1000), "0.00")
        Me.Label324.Text = "Jedinična cijena: " & JedinicnaCijena & " " & Valuta & " / 1 kg"
        Me.Label325.Text = JedinicnaCijena
        Me.TextBox77.Text = JedinicnaCijena
        Me.TextBox81.Text = Me.TextBox80.Text
        Me.Label324.Visible = True

    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox11.SelectedIndexChanged


    End Sub

    Private Sub TextBox77_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox77.TextChanged
        On Error Resume Next
        Dim Cijena As Double = Me.TextBox77.Text
        Dim Kolicina As Double = Me.TextBox81.Text
        Dim JedinicnaCijena As Double = Format(Cijena / (Kolicina / 1000), "0.00")
        Me.Label331.Text = JedinicnaCijena

    End Sub

    Private Sub TextBox81_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox81.TextChanged
        On Error Resume Next
        Dim Cijena As Double = Me.TextBox77.Text
        Dim Kolicina As Double = Me.TextBox81.Text
        Dim JedinicnaCijena As Double = Format(Cijena / (Kolicina / 1000), "0.00")
        Me.Label331.Text = JedinicnaCijena

    End Sub

    Private Sub ComboBox15_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub TabPage24_Click(sender As System.Object, e As System.EventArgs) Handles TabPage24.Click

    End Sub

    Private Sub TabPage24_Enter(sender As Object, e As System.EventArgs) Handles TabPage24.Enter
        On Error Resume Next
        '  Label334.Text = TextBox1.Text & " " & TextBox2.Text  'Klijent
        ' EnergetskaPotrosnjaBindingSource()
        With Me
            If ComboBox1.Text = "" Then
                MsgBox("Odaberite dob.")
                TabControl3.SelectedIndex = 0
                Exit Sub
            End If
            If RadioButton1.Checked = False And RadioButton2.Checked = False Then
                MsgBox("Odaberite spol.")
                TabControl3.SelectedIndex = 0
                Exit Sub
            End If
            If ComboBox2.Text = "" Then
                MsgBox("Odaberite visinu u cm.")
                TabControl3.SelectedIndex = 0
                Exit Sub
            End If
            If ComboBox3.Text = "" Then
                MsgBox("Odaberite masu u kg.")
                TabControl3.SelectedIndex = 0
                Exit Sub
            End If
            '         .BazaEnergetskePotrosnjeBindingSource.RemoveFilter()
            '        BazaEnergetskePotrosnjeBindingSource.Filter = "Korisnik='" & .TextBox1.Text & " " & .TextBox2.Text & "'"
        End With
        TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub Button42_Click_3(sender As System.Object, e As System.EventArgs) Handles Button42.Click
        TabControl3.SelectedIndex = 1

    End Sub

    Private Sub Button44_Click(sender As System.Object, e As System.EventArgs) Handles Button44.Click
        On Error Resume Next
        With Me
            If Val(.ComboBox19.Text) < Val(.TextBox83.Text) Or Val(.ComboBox19.Text) + (Val(.ComboBox20.Text) / 60) > 24 Then Exit Sub
            '      If Val(.ComboBox19.Text) < Val(.TextBox83.Text) Then
            'MsgBox("Greška. " & Val(.ComboBox19.Text) - Val(.TextBox83.Text))
            '     Exit Sub
            '     Else

            Dim i As Integer
            Dim DGV As DataGridView
            DGV = .DataGridView17
            Dim Min As Double = 0
            Dim Energ As Double = 0
            For i = 0 To DGV.RowCount - 1
                Min = Min + DGV.Rows(i).Cells(9).Value
            Next i
            If Min > 60 * 24 Then
                MsgBox("Error.")
                '   .ComboBox14.SelectedIndex = .ComboBox14.SelectedIndex + 1
                Exit Sub
            End If
            '   End If
        End With
        EnergetskaPotrosnja()
        TrajanjeAktivnostiOdDo()
        EnergetskaPotrosnjaUkupno()


    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)
        '  On Error Resume Next
        '  TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox16_TextChanged(sender As Object, e As System.EventArgs)
        '   On Error Resume Next
        '  TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox17_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)
        '  On Error Resume Next
        '  TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox19.SelectedIndexChanged
        On Error Resume Next
        TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox20.SelectedIndexChanged
        On Error Resume Next
        TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox17_TextChanged(sender As Object, e As System.EventArgs)
        '   On Error Resume Next
        '  TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox20_TextChanged(sender As Object, e As System.EventArgs) Handles ComboBox20.TextChanged
        On Error Resume Next
        TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub ComboBox19_TextChanged(sender As Object, e As System.EventArgs) Handles ComboBox19.TextChanged
        On Error Resume Next
        TrajanjeAktivnostiOdDo()

    End Sub

    Private Sub Button45_Click(sender As System.Object, e As System.EventArgs) Handles Button45.Click
        On Error Resume Next
        EditiranjeNamirniceForma.Close()
        EditiranjeNamirniceForma.Show()
        With EditiranjeNamirniceForma
            Dim DGV As DataGridView = Me.DataGridView5

            If Me.TabPage7.CanFocus = True Then
                DGV = Me.DataGridView5
            End If
            If Me.TabPage8.CanFocus = True Then
                DGV = Me.DataGridView9
            End If
            If Me.TabPage9.CanFocus = True Then
                DGV = Me.DataGridView11
            End If
            If Me.TabPage10.CanFocus = True Then
                DGV = Me.DataGridView12
            End If
            If Me.TabPage11.CanFocus = True Then
                DGV = Me.DataGridView13
            End If
            If Me.TabPage12.CanFocus = True Then
                DGV = Me.DataGridView14
            End If

            Dim Namirnica As TextBox = .TextBox1
            Dim Serviranje As Label = .Label4
            Dim Masa As Label = .Label6
            Dim Kolicina As Label = .Label7
            Dim StaroServiranje As Label = .Label8
            Dim StaraMasa As Label = .Label3
            Dim StaraKolicina As Label = .Label5
            Dim StaraMjera As Label = .Label15
            Dim i As Integer = DGV.CurrentRow.Index

            Namirnica.Text = DGV.Rows(i).Cells(5).Value.ToString
            Serviranje.Text = DGV.Rows(i).Cells(7).Value
            StaroServiranje.Text = DGV.Rows(i).Cells(7).Value

            ' Masa.Text = DGV.Rows(i).Cells(10).Value & " g"
            Masa.Text = DGV.Rows(i).Cells(10).Value

            StaraMasa.Text = DGV.Rows(i).Cells(10).Value

            'Kolicina.Text = DGV.Rows(i).Cells(8).Value & " " & DGV.Rows(i).Cells(9).Value 'količina i mjera
            .Label14.Text = DGV.Rows(i).Cells(9).Value   'kolicina (mjera)
            Kolicina.Text = DGV.Rows(i).Cells(8).Value
            StaraKolicina.Text = DGV.Rows(i).Cells(8).Value
            StaraMjera.Text = DGV.Rows(i).Cells(9).Value.ToString

            .HScrollBar1.Value = StaroServiranje.Text * 10
            .HScrollBar2.Value = StaraMasa.Text * 10
            .HScrollBar3.Value = StaraKolicina.Text * 10

            If Namirnica.Text = "" Then
                MsgBox("Odaberite namirnicu iz obroka koju želite urediti.")
                EditiranjeNamirniceForma.Close()
                Exit Sub
            End If

        End With


    End Sub

    Private Sub Button46_Click(sender As System.Object, e As System.EventArgs) Handles Button46.Click
        On Error Resume Next

        'OBROCI
        If MessageBox.Show("Dali želite spremiti promjene u postavke programa?", "Program Prehrane 5.0", _
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                    = DialogResult.Yes Then
            My.Settings.Obrok1 = Me.ComboBox15.Text.ToString
            My.Settings.Obrok2 = Me.ComboBox21.Text.ToString
            My.Settings.Obrok3 = Me.ComboBox22.Text.ToString
            My.Settings.Obrok4 = Me.ComboBox23.Text.ToString
            My.Settings.Obrok5 = Me.ComboBox24.Text.ToString
            My.Settings.Obrok6 = Me.ComboBox25.Text.ToString
            'My.Settings.Save()

            'VALUTA
            Me.Label322.Text = Me.ComboBox26.Text & " /"    'valuta
            Me.Label317.Text = Me.ComboBox26.Text & " /"   'valuta
            My.Settings.Valuta = Me.ComboBox26.Text.ToString

            My.Settings.Save()

            End If

    End Sub

    Private Sub Button47_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox2.CheckedChanged
        On Error Resume Next
        If Me.CheckBox2.Checked = True Then
            Me.ComboBox15.Enabled = True
        Else
            Me.ComboBox15.Enabled = False
        End If
        '    Obroci()

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox3.CheckedChanged
        On Error Resume Next
        If Me.CheckBox3.Checked = True Then
            Me.ComboBox21.Enabled = True
        Else
            Me.ComboBox21.Enabled = False
        End If
        '   Obroci()

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox4.CheckedChanged
        On Error Resume Next
        If Me.CheckBox4.Checked = True Then
            Me.ComboBox22.Enabled = True
        Else
            Me.ComboBox22.Enabled = False
        End If
        '   Obroci()

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox5.CheckedChanged
        On Error Resume Next
        If Me.CheckBox5.Checked = True Then
            Me.ComboBox23.Enabled = True
        Else
            Me.ComboBox23.Enabled = False
        End If
        '   Obroci()

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox6.CheckedChanged
        On Error Resume Next
        If Me.CheckBox6.Checked = True Then
            Me.ComboBox24.Enabled = True
        Else
            Me.ComboBox24.Enabled = False
        End If
        '   Obroci()

    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox7.CheckedChanged
        On Error Resume Next
        If Me.CheckBox7.Checked = True Then
            Me.ComboBox25.Enabled = True
        Else
            Me.ComboBox25.Enabled = False
        End If
        ' Obroci()

    End Sub

    Private Sub Button48_Click(sender As System.Object, e As System.EventArgs)



    End Sub

    Private Sub TextBox82_Click(sender As Object, e As System.EventArgs) Handles TextBox82.Click
        Me.TextBox82.Text = ""
    End Sub

    Private Sub TextBox82_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox82.TextChanged
        On Error Resume Next
        '       If TextBox7.Text = "Pretraži" Then Exit Sub
        '      DataGridView2.DataSource = SveNamirniceBindingSource
        '     Label15.Text = "Sve namirnice"
        '    SveNamirniceBindingSource.RemoveFilter()
        '   SveNamirniceBindingSource.Filter = "NazivNamirnice Like'" & TextBox7.Text & "*'"
        '  DataGridView2.CurrentRow.Selected = False
        If TextBox82.Text = "Pretraži" Then Exit Sub
        If Me.TextBox82.Text = "" Then Me.SveTjelesneAktivnostiBindingSource.RemoveFilter()

        'DataGridView16.DataSource = Me.SportskeAktivnostiBindingSource
        SveTjelesneAktivnostiBindingSource.RemoveFilter()
        SveTjelesneAktivnostiBindingSource.Filter = "OpisTjelesneAktivnosti Like'%" & TextBox82.Text & "%'"
        DataGridView16.CurrentRow.Selected = False

        ' The Filter string can include Boolean expressions.
        '  source1.Filter = "artist = 'Dave Matthews' OR cd = 'Tigerlily'"

      
    End Sub

    Private Sub Button50_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub Button52_Click(sender As System.Object, e As System.EventArgs)



    End Sub

    Private Sub Button53_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub ComboBox14_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox14.SelectedIndexChanged

    End Sub

    Private Sub Button51_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox18.SelectedIndexChanged
        On Error Resume Next
        With Me
            .BazaEnergetskePotrosnjeBindingSource.RemoveFilter()
            BazaEnergetskePotrosnjeBindingSource.Filter = "Dan='" & .ComboBox18.Text & "'"
            If .ComboBox18.Text = "Svi dani" Then .BazaEnergetskePotrosnjeBindingSource.RemoveFilter()
            .DataGridView17.CurrentRow.Selected = False
        End With
        EnergetskaPotrosnjaUkupno()

    End Sub

    Private Sub Button50_Click_1(sender As System.Object, e As System.EventArgs) Handles Button50.Click
        On Error Resume Next
        With Me

            Dim DGV As DataGridView = .DataGridView17
            Dim BS As BindingSource = BazaEnergetskePotrosnjeBindingSource
            Dim i As Integer

            'briši odabrane aktivnosti

                For i = 0 To DGV.RowCount - 1
                    DGV.Rows.Remove(DGV.CurrentRow)
                Next i
            BS.AddNew()

            .DataGridView17.CurrentRow.Selected = False

            .TextBox83.Text = 0
            .TextBox84.Text = 0
            .ComboBox19.Text = 0
            .ComboBox20.Text = 0
            .Label315.Text = 0

        End With

        TrajanjeAktivnostiOdDo()
        EnergetskaPotrosnjaUkupno()


    End Sub

    Private Sub TextBox83_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox83.TextChanged

    End Sub

    Private Sub ComboBox16_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox16.SelectedIndexChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub ComboBox16_TextChanged1(sender As Object, e As System.EventArgs) Handles ComboBox16.TextChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub ComboBox17_SelectedIndexChanged_1(sender As System.Object, e As System.EventArgs) Handles ComboBox17.SelectedIndexChanged
        On Error Resume Next
        With Me
            If .ComboBox17.SelectedIndex = 0 Then
                Ispis()
                Ispis()
                .GroupBox56.Visible = True
                .GroupBox57.Visible = True
            Else
                .GroupBox56.Visible = False
                .GroupBox57.Visible = False
            End If
            If .ComboBox17.SelectedIndex = 1 Then
                IspisOsobniPodaci()
            End If
            If .ComboBox17.SelectedIndex = 2 Then
                IspisParametri()
            End If

        End With
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox8.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox9.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox10.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox11.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox12.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox13.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox14.CheckedChanged
        On Error Resume Next
        Ispis()

    End Sub

    Private Sub Button9_Click_1(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub Button49_Click(sender As System.Object, e As System.EventArgs)



    End Sub

    Private Sub Button9_Click_2(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button10_Click_1(sender As System.Object, e As System.EventArgs)



    End Sub

    Private Sub Button9_Click_3(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        Me.ColorDialog1.ShowDialog()
        DgvStyle()

    End Sub

    Private Sub Button10_Click_2(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        With Me
            .ColorDialog1.ShowDialog()

            Dim DGV As DataGridView = .DataGridView2
            ' Dim Boja As Color = Color.Lavender
            Dim Boja As Color = .ColorDialog1.Color
            Dim i As Integer = DGV.CurrentRow.Index


            'Pracenje antropometrijskih parametara
            '           DGV = .DataGridView10
            '          For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja

            'Next

            'baza klijenata
            'DGV = .DataGridView6
            '      For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Tjalesna aktivnost detaljni izracun
            'DGV = .DataGridView16
            'For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Dodatna aktivnost
            '       DGV = .DataGridView3
            '      For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Vrsta dijete
            '     DGV = .DataGridView1
            '    For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Sve namirnice
            DGV = .DataGridView2
            For i = 0 To DGV.Rows.Count - 1 Step 2
                DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            Next

            My.Settings.DgvBojaSelect = Boja
            My.Settings.Save()

            'Dorucak
            '            DGV = .DataGridView5
            '           For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Baza naziva jelovnika
            '     DGV = .DataGridView7
            '    For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next

            'Cijene
            '     DGV = .DataGridView15
            '    For i = 0 To DGV.Rows.Count - 1 Step 2
            'DGV.Rows(i).DefaultCellStyle.SelectionBackColor = Boja
            'Next


        End With
    End Sub

    Private Sub Button11_Click_1(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub Button11_Click_2(sender As System.Object, e As System.EventArgs) Handles Button11.Click
         On Error Resume Next
        '       'BOJA
        Dim DGV As DataGridView = Me.DataGridView2
        Dim i As Integer = DGV.CurrentRow.Index
        Dim Boja As Color = SystemColors.ControlLight    'boja pozadine baze namirnica
        Dim SelektBoja As Color = SystemColors.GradientInactiveCaption   'boja selektiranog reda baze namirnica
        'Sve Namirnice
        DGV = Me.DataGridView2
        For i = 0 To DGV.Rows.Count - 1 Step 2
            ' DGV.Rows(i).DefaultCellStyle.BackColor = My.Settings.DgvBoja
            DGV.Rows(i).DefaultCellStyle.BackColor = Boja
        Next
        My.Settings.DgvBoja = Boja
        My.Settings.Save()
        'BOJA-SELEKTIRANI RED
        'Sve namirnice
        For i = 0 To DGV.Rows.Count - 1 Step 2
            DGV.Rows(i).DefaultCellStyle.SelectionBackColor = SelektBoja
        Next
        My.Settings.DgvBojaSelect = SelektBoja
        My.Settings.Save()

        'OBROCI
        With Me
            .ComboBox15.Text = "Doručak"
            .ComboBox21.Text = "Jutarnja užina"
            .ComboBox22.Text = "Ručak"
            .ComboBox23.Text = "Popodnevna užina"
            .ComboBox24.Text = "Večera"
            .ComboBox25.Text = "Obrok pred spavanje"

            My.Settings.Obrok1 = .ComboBox15.Text
            My.Settings.Obrok2 = .ComboBox21.Text
            My.Settings.Obrok3 = .ComboBox22.Text
            My.Settings.Obrok4 = .ComboBox23.Text
            My.Settings.Obrok5 = .ComboBox24.Text
            My.Settings.Obrok6 = .ComboBox25.Text
            My.Settings.Save()
        End With

        'VALUTA
        Me.ComboBox26.Text = "HRK"
        Me.Label322.Text = Me.ComboBox26.Text & " /"    'valuta
        Me.Label317.Text = Me.ComboBox26.Text & " /"   'valuta

        My.Settings.Valuta = Me.ComboBox26.Text.ToString
        My.Settings.Save()


    End Sub

    Private Sub ComboBox22_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox22.SelectedIndexChanged

    End Sub

    Private Sub GroupBox54_Enter(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TabPage23_Click(sender As System.Object, e As System.EventArgs) Handles TabPage23.Click

    End Sub

    Private Sub Button49_Click_1(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Label21_Click(sender As System.Object, e As System.EventArgs) Handles Label21.Click

    End Sub

    Private Sub Label21_TextChanged(sender As Object, e As System.EventArgs) Handles Label21.TextChanged
      
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyUp
        On Error Resume Next

        'start
        If My.Settings.PP5StartAktivacija = "Da" Or My.Settings.PP5StartTrajnaLicencaAktivacija = "Da" Then
            Start()
        End If
        'demo
        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then
            Demo()
        End If

      
    End Sub

    Private Sub TextBox85_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox85.KeyPress
        'samo brojevi
        If Not Char.IsDigit(e.KeyChar) And Not Char.IsControl(e.KeyChar) And Not e.KeyChar = "," Then
            e.Handled = True
        End If

    End Sub

    Private Sub TextBox85_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox85.TextChanged
        On Error Resume Next
        ' If TextBox85.Text < 0.00001 Then Exit Sub
        Dim Masa As Double = (TextBox85.Text / Label363.Text) * Label228.Text
        Dim Kolicina As Double = (TextBox85.Text / Label363.Text) * Label229.Text
        If RadioButton13.Checked = True Then
            If Masa < 100 Then
                TextBox6.Text = Format(Masa, "0.0")   'Masa
                TextBox12.Text = Format(Kolicina, "0.00")   'Kolicina
            Else
                TextBox6.Text = Format(Masa, "0")   'Masa
                TextBox12.Text = Format(Kolicina, "0.00")   'Kolicina
            End If

        End If

        Mjera()

    End Sub

    Private Sub Button49_Click_2(sender As System.Object, e As System.EventArgs) Handles Button49.Click
        TabControl1.SelectedIndex = 5
        Dim BrojDijete As Integer = Me.Label179.Text
        If BrojDijete > 2 Then
            'If Me.Label21.Text <> "Normalna prehrana (za osobe od 9 do 14 god.)" Or Me.Label21.Text <> "Normalna prehrana (za osobe od 14 do 18 god.)" Then
            If Val(Me.ComboBox1.Text) < 18 And Val(Me.ComboBox1.Text) >= 9 Then   'djeca
                MsgBox("Za osobe mlađe od 18 godina, " & Me.Label21.Text & " smije se provoditi isključivo pod nadzorom stručne osobe uz liječničko dopuštenje i praćenje!")
            End If
        End If

    End Sub

    Private Sub Button51_Click_1(sender As System.Object, e As System.EventArgs) Handles Button51.Click
        TabControl1.SelectedIndex = 3

    End Sub

    Private Sub TabPage19_Click(sender As System.Object, e As System.EventArgs) Handles TabPage19.Click

    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox12.SelectedIndexChanged
        Me.TextBox78.Text = Me.ComboBox12.Text

    End Sub

    Private Sub Button52_Click_1(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button53_Click_1(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button52_Click_2(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button53_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button53.Click
        'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click 
        ' This call is required by the designer.
        ' InitializeComponent()

        Dim PM As Charting.PrintingManager
        PM = Me.Chart1.Printing

        '     PM.PrintPreview()
        'PM.PrintDocument.Print()
        ' PM.PageSetup()
        '    PrintDocument2.Print()
        With Me
            '   StringToPrint = .RichTextBoxPrintCtrl1.Text


            .PrintDialog1.Document = .PrintDocument1
            .PageSetupDialog1.Document = .PrintDocument1
            If .PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                If .PageSetupDialog1.ShowDialog = DialogResult.OK Then
                    '.PrintPreviewDialog1.Document = .PrintDocument1
                    '.PrintPreviewDialog1.ShowDialog()
                    PrintDocument1.Print()
                End If
            End If
        End With

        '      .PrintDialog1.Document = .PrintDocument1
        '     .PageSetupDialog1.Document = .PrintDocument1
        ' If Me.PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
        '  If Me.PageSetupDialog1.ShowDialog = DialogResult.OK Then
        ''.PrintPreviewDialog1.Document = .PrintDocument1
        ''.PrintPreviewDialog1.ShowDialog()
        '   PrintDocument1.Print()
        'PM.PrintPreview()
        ' End If
        ' End If


    End Sub

   
    Private Sub Button54_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub ComboBox26_Click(sender As Object, e As System.EventArgs) Handles ComboBox26.Click
     

    End Sub

    Private Sub ComboBox26_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox26.SelectedIndexChanged
        'VALUTA
        Me.Label322.Text = Me.ComboBox26.Text & " /"    'valuta
        Me.Label317.Text = Me.ComboBox26.Text & " /"   'valuta
        My.Settings.Valuta = Me.ComboBox26.Text.ToString

        My.Settings.Save()

    End Sub

    Private Sub ComboBox26_TextChanged(sender As Object, e As System.EventArgs) Handles ComboBox26.TextChanged
        'VALUTA
        Me.Label322.Text = Me.ComboBox26.Text & " /"    'valuta
        Me.Label317.Text = Me.ComboBox26.Text & " /"   'valuta
        My.Settings.Valuta = Me.ComboBox26.Text.ToString

        My.Settings.Save()

    End Sub
End Class
