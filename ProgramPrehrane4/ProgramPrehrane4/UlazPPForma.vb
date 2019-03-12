Public Class UlazPPForma

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.programprehrane.com")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        PricekajteTrenutakDemoForma.Show()
        Form1.Show()
        Form1.Text = "Program Prehrane 5.0 Demo"
        ' Form1.AktivacijaPuneVerzijeProgramaToolStripMenuItem.Visible = True
        '    Form1.PriručnikToolStripMenuItem1.Enabled = False
        Demo()

    End Sub

    Private Sub UlazPP4Forma_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        On Error Resume Next
        Me.MaskedTextBox1.Select()

        If My.Settings.Instalacija = "Ne" Then     'Thank you for installing
            My.Settings.Instalacija = "Da"
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5Instalacija.aspx")
        End If

        'DEMO
        If My.Settings.PP5StartAktivacija = "Ne" And _
            My.Settings.PP5PremiumAktivacija = "Ne" And _
            My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" And _
            My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" And _
            My.Settings.PP5Premium2MjAktivacija = "Ne" Then
            Demo()
        End If
        '   If My.Settings.AktivacijskiKljuc = "Da" Then
        'Me.Hide()
        '    PricekajteTrenutakProForma.Show()
        '   Form1.Show()
        '  Me.Text = "Program Prehrane 5.0 Pro"
        ' Form1.AktivacijaPuneVerzijeProgramaToolStripMenuItem.Visible = False
        'Form1.PriručnikToolStripMenuItem1.Enabled = True
        '    Form1.Button39.Enabled = True
        '   End If

    End Sub

    Private Sub UlazPP4Forma_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'provjeri rezoluciju monitora
        Dim intX As Integer = Screen.PrimaryScreen.Bounds.Width
        Dim intY As Integer = Screen.PrimaryScreen.Bounds.Height
        If intX < 1024 Or intY < 768 Then
            MsgBox("Za pravilan prikaz programa minimalna rezolucija monitora treba biti 1024x768.", MsgBoxStyle.Information)
        End If


        'TRAJNA LICENCA
        'START
        If My.Settings.PP5StartTrajnaLicencaAktivacija = "Da" Then
            My.Settings.PP5StartAktivacija = "Da"
            Me.GroupBox1.Enabled = False  'Demo
            Me.GroupBox2.Enabled = True   'Start
            Me.Button3.Enabled = False
            Me.MaskedTextBox1.Enabled = False
            Me.GroupBox3.Enabled = True   'Premium
            Me.Label6.Enabled = False
            Exit Sub
        End If

        'PREMIUM
        If My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Da" Then
            My.Settings.PP5PremiumAktivacija = "Da"
            Me.GroupBox1.Enabled = False  'Demo
            Me.GroupBox2.Enabled = False   'Start
            Me.GroupBox3.Enabled = True   'Premium
            Me.Button5.Enabled = False
            Me.MaskedTextBox2.Enabled = False
            Me.Label7.Enabled = False
            Exit Sub
        End If


        'START
        If My.Settings.PP5StartAktivacija = "Da" Then
            If DateDiff(DateInterval.Day, My.Settings.PP5StartAktivacijaDatum, Now) > 365 Then '365 is the expiration date
                If MessageBox.Show("Licenca za korištenje PROGRAM PREHRANE 5.0 START je istekla." _
                                    & vbCrLf & "Želite li obnoviti licencu?", "Program Prehrane 5.0", _
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                  = DialogResult.Yes Then
                    System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
                End If
                '     Me.Label8.Enabled = True
                Me.GroupBox1.Enabled = True   'Demo
                Me.GroupBox2.Enabled = True   'Start
                Me.GroupBox3.Enabled = True   'Premium
                '     Me.Button2.Enabled = True
                '   Me.Label6.Enabled = True
                '    Me.MaskedTextBox1.Enabled = True
                '   Me.Button3.Enabled = True
                My.Settings.PP5StartAktivacija = "Ne"
                '         My.Settings.PP5StartKod1 = My.Settings.PP5StartKod1.ToString + 72
                '        My.Settings.PP5StartKod2 = My.Settings.PP5StartKod2.ToString + 53
                '       My.Settings.PP5StartKod3 = My.Settings.PP5StartKod3.ToString + 23
                '      My.Settings.PP5StartKod4 = My.Settings.PP5StartKod4.ToString + 51
                My.Settings.Save()
            Else
                '        Me.Label8.Enabled = False
                'ako jos traje licenca za START
                Me.GroupBox1.Enabled = False  'Demo
                Me.GroupBox2.Enabled = True   'Start
                Me.Button3.Enabled = False
                Me.MaskedTextBox1.Enabled = False
                Me.GroupBox3.Enabled = True   'Premium
                Me.Label6.Enabled = False

                '     Me.Button2.Enabled = False
                '    Me.Label6.Enabled = False
                '    Me.MaskedTextBox1.Enabled = False
                '   Me.Button3.Enabled = False
            End If
        End If

        'PREMIUM
        If My.Settings.PP5PremiumAktivacija = "Da" Then
            If DateDiff(DateInterval.Day, My.Settings.PP5PremiumAktivacijaDatum, Now) > 365 Then '365 is the expiration date
                If MessageBox.Show("Licenca za korištenje PROGRAM PREHRANE 5.0 PREMIUM je istekla." _
                                    & vbCrLf & "Želite li obnoviti licencu?", "Program Prehrane 5.0", _
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                  = DialogResult.Yes Then
                    System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
                End If
                'premium
                Me.GroupBox1.Enabled = True   'Demo
                Me.GroupBox2.Enabled = True   'Start
                Me.GroupBox3.Enabled = True   'Premium

                '   Me.Label7.Enabled = True
                '   Me.Button5.Enabled = True
                '   Me.MaskedTextBox2.Enabled = True
                'start
                '     Me.Label8.Enabled = True
                '   Me.GroupBox2.Enabled = True   'Start


                '  Me.Button1.Enabled = True
                '  Me.Button2.Enabled = True
                '  Me.Label6.Enabled = True
                ' Me.MaskedTextBox1.Enabled = True
                ' Me.Button3.Enabled = True
                My.Settings.PP5PremiumAktivacija = "Ne"
                '           My.Settings.PP5PremiumKod1 = My.Settings.PP5PremiumKod1.ToString + 43
                '          My.Settings.PP5PremiumKod2 = My.Settings.PP5PremiumKod2.ToString + 62
                '         My.Settings.PP5PremiumKod3 = My.Settings.PP5PremiumKod3.ToString + 54
                '        My.Settings.PP5PremiumKod4 = My.Settings.PP5PremiumKod4.ToString + 68
                My.Settings.Save()
            Else
                'premium
                'ako jos traje licenca za PREMIUM
                Me.GroupBox1.Enabled = False  'Demo
                Me.GroupBox2.Enabled = False   'Start
                Me.GroupBox3.Enabled = True   'Premium
                Me.Button5.Enabled = False
                Me.MaskedTextBox2.Enabled = False
                Me.Label7.Enabled = False
                '  Me.Label7.Enabled = False
                '    Me.Button5.Enabled = False
                '   Me.MaskedTextBox2.Enabled = False
                'start
                '      Me.Label8.Enabled = False

                '    Me.Button2.Enabled = False
                '    Me.Button1.Enabled = False
                '   Me.Label6.Enabled = False
                '  Me.MaskedTextBox1.Enabled = False
                '  Me.Button3.Enabled = False
            End If
        End If
        '   Dim Istek As Integer = My.Settings.PP5StartIstekLicence - Date.Today.DayOfYear.ToString
        'If Istek=>

        '   If My.Settings.PP5StartIstekLicence.ToString > 0 Then



        'PREMIUM 2 MJESECNA LICENCA
        If My.Settings.PP5Premium2MjAktivacija = "Da" Then
            'provjeri dali je istekla dvomjesecna licenca
            If DateDiff(DateInterval.Day, My.Settings.PP5Premium2MjAktivacijaDatum, Now) > 61 Then '61 is the expiration date
                If MessageBox.Show("Licenca za korištenje PROGRAM PREHRANE 5.0 PREMIUM je istekla." _
                                    & vbCrLf & "Želite li obnoviti licencu?", "Program Prehrane 5.0", _
                                     MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                  = DialogResult.Yes Then
                    System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
                End If
                'premium
                Me.GroupBox1.Enabled = True   'Demo
                Me.GroupBox2.Enabled = True   'Start
                Me.GroupBox3.Enabled = True   'Premium

                My.Settings.PP5Premium2MjAktivacija = "Ne"

                My.Settings.Save()
            Else
                'premium
                'ako jos traje licenca za PREMIUM
                Me.GroupBox1.Enabled = False  'Demo
                Me.GroupBox2.Enabled = False   'Start
                Me.GroupBox3.Enabled = True   'Premium
                Me.Button5.Enabled = False
                Me.MaskedTextBox2.Enabled = False
                Me.Label7.Enabled = False

            End If
        End If




    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        On Error Resume Next

        'trajna licenca
        If My.Settings.PP5StartTrajnaLicencaAktivacija.ToString = "Da" Then
            My.Settings.PP5StartAktivacija = "Da"
            My.Settings.Save()
            Me.Hide()
            PricekajteTrenutakStartForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Start"
            Start()
            Exit Sub
        End If

        If Me.MaskedTextBox1.Text = My.Settings.PP5StartTrajnaLicencaKod Then
            My.Settings.PP5PremiumAktivacija = "Ne"   'Disabled Premium verzija
            My.Settings.PP5Premium2MjAktivacija = "Ne"  'Disabled Premium 2 mj
            My.Settings.PP5StartAktivacija = "Da"   'uspjesna aktivacija
            My.Settings.PP5StartTrajnaLicencaAktivacija = "Da"
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5StartAktivacija.aspx")
            Me.Hide()
            PricekajteTrenutakStartForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Start"
            Start()
            Exit Sub
        End If


        'godisnja licenca
        If My.Settings.PP5StartAktivacija.ToString = "Da" Then
            Me.Hide()
            PricekajteTrenutakStartForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Start"
            '   Form1.PriručnikToolStripMenuItem1.Enabled = True
            '  Form1.Button39.Enabled = True
            Start()
            Exit Sub
        End If

        If Me.MaskedTextBox1.Text = My.Settings.PP5StartKod1 & "-" & My.Settings.PP5StartKod2 & "-" & My.Settings.PP5StartKod3 & "-" & My.Settings.PP5StartKod4 Then

            My.Settings.PP5PremiumAktivacija = "Ne"   'Disabled Premium verzija
            My.Settings.PP5Premium2MjAktivacija = "Ne"  'Disabled Premium 2 mj
            My.Settings.PP5StartAktivacija = "Da"   'uspjesna aktivacija
            My.Settings.PP5StartAktivacijaDatum = Today.Date    'datum aktivacije

            My.Settings.PP5StartKod1 = My.Settings.PP5StartKod1.ToString + 72
            My.Settings.PP5StartKod2 = My.Settings.PP5StartKod2.ToString + 53
            My.Settings.PP5StartKod3 = My.Settings.PP5StartKod3.ToString + 23
            My.Settings.PP5StartKod4 = My.Settings.PP5StartKod4.ToString + 51
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5StartAktivacija.aspx")
            Me.Hide()
            PricekajteTrenutakStartForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Start"
            '    Form1.PriručnikToolStripMenuItem1.Enabled = True
            '   Form1.Button39.Enabled = True
            Start()
        Else
            If MessageBox.Show("Pogrešan unos." _
                                      & vbCrLf & "Želite li naručiti aktivacijski kod za pokretanje računalnog programa PROGRAM PREHRANE 5.0 START?", "Program Prehrane 5.0", _
                                       MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                                    = DialogResult.Yes Then
                System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
            End If
        End If


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        On Error Resume Next

        'trajna licenca
        If My.Settings.PP5PremiumTrajnaLicencaAktivacija.ToString = "Da" Then
            My.Settings.PP5PremiumAktivacija = "Da"
            My.Settings.Save()
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            Premium()
            Exit Sub
        End If

        If Me.MaskedTextBox2.Text = My.Settings.PP5PremiumTrajnaLicencaKod Then
            My.Settings.PP5PremiumAktivacija = "Da"   'Disabled Premium verzija
            My.Settings.PP5Premium2MjAktivacija = "Ne"  'Disabled Premium 2 mj
            My.Settings.PP5StartAktivacija = "Ne"   'uspjesna aktivacija
            My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Da"
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5PremiumAktivacija.aspx")
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            Premium()
            Exit Sub
        End If


        'godisnja licenca
        If My.Settings.PP5PremiumAktivacija.ToString = "Da" Then
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            '    Form1.PriručnikToolStripMenuItem1.Enabled = True
            '    Form1.Button39.Enabled = True
            Premium()
            Exit Sub
        End If

        ' godisnja licenca login
        If Me.MaskedTextBox2.Text = My.Settings.PP5PremiumKod1 & "-" & My.Settings.PP5PremiumKod2 & "-" & My.Settings.PP5PremiumKod3 & "-" & My.Settings.PP5PremiumKod4 Then

            My.Settings.PP5StartAktivacija = "Ne"   'Disabled Start verzija
            My.Settings.PP5PremiumAktivacija = "Da"   'Uspjesna aktivacija Premium verzije
            My.Settings.PP5Premium2MjAktivacija = "Ne"  'Disabled Premium 2 mj
            My.Settings.PP5PremiumAktivacijaDatum = Today.Date    'datum aktivacije

            My.Settings.PP5PremiumKod1 = My.Settings.PP5PremiumKod1.ToString + 43
            My.Settings.PP5PremiumKod2 = My.Settings.PP5PremiumKod2.ToString + 62
            My.Settings.PP5PremiumKod3 = My.Settings.PP5PremiumKod3.ToString + 54
            My.Settings.PP5PremiumKod4 = My.Settings.PP5PremiumKod4.ToString + 68
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5PremiumAktivacija.aspx")
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            '  Form1.PriručnikToolStripMenuItem1.Enabled = True
            '   Form1.Button39.Enabled = True
            Premium()
            Exit Sub
        End If


        '2 godisnja licenca
        If My.Settings.PP5Premium2MjAktivacija.ToString = "Da" Then
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            '    Form1.PriručnikToolStripMenuItem1.Enabled = True
            '    Form1.Button39.Enabled = True
            Premium()
            Exit Sub
        End If

        '2 mjesecna licenca login
        If Me.MaskedTextBox2.Text = My.Settings.PP5Premium2MjKod1 & "-" & My.Settings.PP5Premium2MjKod2 & "-" & My.Settings.PP5Premium2MjKod3 & "-" & My.Settings.PP5Premium2MjKod4 Then

            My.Settings.PP5StartAktivacija = "Ne"   'Disabled Start verzija
            My.Settings.PP5PremiumAktivacija = "Ne"   'Uspjesna aktivacija Premium verzije
            My.Settings.PP5Premium2MjAktivacija = "Da"
            My.Settings.PP5Premium2MjAktivacijaDatum = Today.Date    'datum aktivacije

            My.Settings.PP5Premium2MjKod1 = My.Settings.PP5Premium2MjKod1.ToString + 43
            My.Settings.PP5Premium2MjKod2 = My.Settings.PP5Premium2MjKod2.ToString + 62
            My.Settings.PP5Premium2MjKod3 = My.Settings.PP5Premium2MjKod3.ToString + 54
            My.Settings.PP5Premium2MjKod4 = My.Settings.PP5Premium2MjKod4.ToString + 68
            My.Settings.Save()
            System.Diagnostics.Process.Start("http://www.programprehrane.com/PP5PremiumAktivacija.aspx")
            Me.Hide()
            PricekajteTrenutakPremiumForma.Show()
            Form1.Show()
            Form1.Text = "Program Prehrane 5.0 Premium"
            '  Form1.PriručnikToolStripMenuItem1.Enabled = True
            '   Form1.Button39.Enabled = True
            Premium()
            Exit Sub
        End If




        If MessageBox.Show("Pogrešan unos." _
                    & vbCrLf & "Želite li naručiti aktivacijski kod za pokretanje računalnog programa PROGRAM PREHRANE 5.0 PREMIUM?", "Program Prehrane 5.0", _
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) _
                    = DialogResult.Yes Then
            System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")
        End If


    End Sub



    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        System.Diagnostics.Process.Start("http://www.programprehrane.com/Narudzba.aspx")

    End Sub

  
End Class