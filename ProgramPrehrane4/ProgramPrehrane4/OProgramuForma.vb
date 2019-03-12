Public Class OProgramuForma

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        System.Diagnostics.Process.Start("http://www.programprehrane.com")

    End Sub

    Private Sub OProgramuForma_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '   Me.LabelVersion.Text = String.Format("Version {0}", My.Application.Info.Version.ToString)
        '  Dim Verzija As String = String.Format("Verzija:", My.Application.Info.Version.ToString)
        ' Dim Verzija As String = My.Application.Info.Version.ToString
        Dim Verzija As String = "5.0"

        'DEMO
        If My.Settings.PP5StartAktivacija = "Ne" And My.Settings.PP5PremiumAktivacija = "Ne" And My.Settings.PP5Premium2MjAktivacija = "Ne" Then
            Me.Label4.Text = "Ovaj program nije licenciran."
            Me.Label2.Text = "Verzija: " & Verzija & " Demo"
            Me.ProgressBar1.Visible = False
            Me.Label7.Text = ""
        End If

        'START
        If My.Settings.PP5StartAktivacija = "Da" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Ne" Then
            Dim Istek As Integer = 365 - DateDiff(DateInterval.Day, My.Settings.PP5StartAktivacijaDatum, Now)
            Me.ProgressBar1.Value = (Istek / 365) * 100
            Me.Label7.Text = Format((Istek / 365) * 100, "0.0") & "%"
            Me.Label4.Text = "Ovaj program je licenciran." & vbCrLf & "Do isteka licence je ostalo još " & Istek & " dana."
            Me.Label2.Text = "Verzija: " & Verzija & " Start"
        End If

        'PREMIUM
        If My.Settings.PP5PremiumAktivacija = "Da" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Ne" Then
            Dim Istek1 As Integer = 365 - DateDiff(DateInterval.Day, My.Settings.PP5PremiumAktivacijaDatum, Now)
            Me.ProgressBar1.Value = (Istek1 / 365) * 100
            Me.Label7.Text = Format((Istek1 / 365) * 100, "0.0") & "%"
            Me.Label4.Text = "Ovaj program je licenciran." & vbCrLf & "Do isteka licence je ostalo još " & Istek1 & " dana."
            Me.Label2.Text = "Verzija: " & Verzija & " Premium"
        End If


        'START - trajna licenca
        If My.Settings.PP5StartAktivacija = "Da" And My.Settings.PP5StartTrajnaLicencaAktivacija = "Da" Then
            '  Dim Istek As Integer = 365 - DateDiff(DateInterval.Day, My.Settings.PP5StartAktivacijaDatum, Now)
            '  Me.ProgressBar1.Value = (Istek / 365) * 100
            '   Me.Label7.Text = Format((Istek / 365) * 100, "0.0") & "%"
            Me.Label4.Text = "Ovaj program je trajno licenciran."
            Me.Label2.Text = "Verzija: " & Verzija & " Start"
            Me.ProgressBar1.Visible = False
            Me.Label7.Text = ""
        End If

        'PREMIUM - trajna licenca
        If My.Settings.PP5PremiumAktivacija = "Da" And My.Settings.PP5PremiumTrajnaLicencaAktivacija = "Da" Then
            '   Dim Istek1 As Integer = 365 - DateDiff(DateInterval.Day, My.Settings.PP5PremiumAktivacijaDatum, Now)
            '   Me.ProgressBar1.Value = (Istek1 / 365) * 100
            '  Me.Label7.Text = Format((Istek1 / 365) * 100, "0.0") & "%"
            Me.Label4.Text = "Ovaj program je trajno licenciran."
            Me.Label2.Text = "Verzija: " & Verzija & " Premium"
            Me.ProgressBar1.Visible = False
            Me.Label7.Text = ""
        End If

        'PREMIUM - 2 mjesecna licenca
        If My.Settings.PP5Premium2MjAktivacija = "Da" Then
            Dim Istek1 As Integer = 61 - DateDiff(DateInterval.Day, My.Settings.PP5Premium2MjAktivacijaDatum, Now)
            Me.ProgressBar1.Value = (Istek1 / 61) * 100
            Me.Label7.Text = Format((Istek1 / 61) * 100, "0.0") & "%"
            Me.Label4.Text = "Ovaj program je licenciran." & vbCrLf & "Do isteka licence je ostalo još " & Istek1 & " dana."
            Me.Label2.Text = "Verzija: " & Verzija & " Premium"
        End If



    End Sub
End Class