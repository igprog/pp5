Module MjeraEditiranjeNamirniceModul
    Sub MjeraEditiranjeNamirnice()
        With EditiranjeNamirniceForma
            '    If .Label230.Text = "jušna žlica" Then
            'If .TextBox12.Text <= 1 Then .Label103.Text = "jušna žlica"
            '     If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "jušne žlice"
            '    If .TextBox12.Text >= 5 Then .Label103.Text = "jušnih žlica"
            '   If .TextBox12.Text < 1 Then .Label103.Text = "jušne žlice"
            '  End If
            Dim Kolicina As Double = .Label7.Text
            Dim StaraMjera As String = .Label15.Text   'stara mjera
            Dim NovaMjera As Label = .Label14    'nova mjera

            If StaraMjera = "jušna žlica" Or StaraMjera = "jušne žlice" Or StaraMjera = "jušnih žlica" Then
                If Kolicina = 1 Then NovaMjera.Text = "jušna žlica"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "jušne žlice"
                If Kolicina >= 5 Then NovaMjera.Text = "jušnih žlica"
                If Kolicina < 1 Then NovaMjera.Text = "jušne žlice"
            End If

            If StaraMjera = "šalica" Or StaraMjera = "šalice" Then
                If Kolicina = 1 Then NovaMjera.Text = "šalica"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "šalice"
                If Kolicina >= 5 Then NovaMjera.Text = "šalica"
                If Kolicina < 1 Then NovaMjera.Text = "šalice"
            End If

            If StaraMjera = "komad" Or StaraMjera = "komada" Then
                If Kolicina = 1 Then NovaMjera.Text = "komad"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "komada"
                If Kolicina >= 5 Then NovaMjera.Text = "komada"
                If Kolicina < 1 Then NovaMjera.Text = "komada"
            End If

            If StaraMjera = "mali komad" Or StaraMjera = "malih komada" Or StaraMjera = "malog komada" Or StaraMjera = "mala komada" Then
                If Kolicina = 1 Then NovaMjera.Text = "mali komad"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "mala komada"
                If Kolicina >= 5 Then NovaMjera.Text = "malih komada"
                If Kolicina < 1 Then NovaMjera.Text = "malog komada"
            End If

            If StaraMjera = "plod" Or StaraMjera = "plodova" Or StaraMjera = "ploda" Then
                If Kolicina = 1 Then NovaMjera.Text = "plod"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "ploda"
                If Kolicina >= 5 Then NovaMjera.Text = "plodova"
                If Kolicina < 1 Then NovaMjera.Text = "ploda"
            End If

            If StaraMjera = "limenka" Or StaraMjera = "limenke" Or StaraMjera = "limenki" Then
                If Kolicina = 1 Then NovaMjera.Text = "limenka"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "limenke"
                If Kolicina >= 5 Then NovaMjera.Text = "limenki"
                If Kolicina < 1 Then NovaMjera.Text = "limenke"
            End If

            If StaraMjera = "porcija" Or StaraMjera = "porcije" Then
                If Kolicina = 1 Then NovaMjera.Text = "porcija"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "porcije"
                If Kolicina >= 5 Then NovaMjera.Text = "porcija"
                If Kolicina < 1 Then NovaMjera.Text = "porcije"
            End If

            If StaraMjera = "čajna žličica" Or StaraMjera = "čajnih žličica" Or StaraMjera = "čajne žličice" Then
                If Kolicina = 1 Then NovaMjera.Text = "čajna žličica"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "čajne žličice"
                If Kolicina >= 5 Then NovaMjera.Text = "čajnih žličica"
                If Kolicina < 1 Then NovaMjera.Text = "čajne žličice"
            End If

            If StaraMjera = "zrno" Or StaraMjera = "zrna" Then
                If Kolicina = 1 Then NovaMjera.Text = "zrno"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "zrna"
                If Kolicina >= 5 Then NovaMjera.Text = "zrna"
                If Kolicina < 1 Then NovaMjera.Text = "zrna"
            End If

            If StaraMjera = "veliki plod" Or StaraMjera = "velikih plodova" Or StaraMjera = "velika ploda" Or StaraMjera = "velikog ploda" Then
                If Kolicina = 1 Then NovaMjera.Text = "veliki plod"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "velika ploda"
                If Kolicina >= 5 Then NovaMjera.Text = "velikih plodova"
                If Kolicina < 1 Then NovaMjera.Text = "velikog ploda"
            End If

            If StaraMjera = "kriška" Or StaraMjera = "kriške" Or StaraMjera = "kriški" Then
                If Kolicina = 1 Then NovaMjera.Text = "kriška"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "kriške"
                If Kolicina >= 5 Then NovaMjera.Text = "kriški"
                If Kolicina < 1 Then NovaMjera.Text = "kriške"
            End If

            If StaraMjera = "mala kriška" Or StaraMjera = "malih kriški" Or StaraMjera = "male kriške" Then
                If Kolicina = 1 Then NovaMjera.Text = "mala kriška"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "male kriške"
                If Kolicina >= 5 Then NovaMjera.Text = "malih kriški"
                If Kolicina < 1 Then NovaMjera.Text = "male kriške"
            End If

            If StaraMjera = "čaša" Or StaraMjera = "čaše" Then
                If Kolicina = 1 Then NovaMjera.Text = "čaša"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "čaše"
                If Kolicina >= 5 Then NovaMjera.Text = "čaša"
                If Kolicina < 1 Then NovaMjera.Text = "čaše"
            End If

            If StaraMjera = "boca" Or StaraMjera = "boce" Then
                If Kolicina = 1 Then NovaMjera.Text = "boca"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "boce"
                If Kolicina >= 5 Then NovaMjera.Text = "boca"
                If Kolicina < 1 Then NovaMjera.Text = "boce"
            End If

            If StaraMjera = "polovica" Or StaraMjera = "polovice" Then
                If Kolicina = 1 Then NovaMjera.Text = "polovica"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "polovice"
                If Kolicina >= 5 Then NovaMjera.Text = "polovica"
                If Kolicina < 1 Then NovaMjera.Text = "polovice"
            End If

            If StaraMjera = "limenka" Or StaraMjera = "limenke" Or StaraMjera = "limenki" Then
                If Kolicina = 1 Then NovaMjera.Text = "limenka"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "limenke"
                If Kolicina >= 5 Then NovaMjera.Text = "limenki"
                If Kolicina < 1 Then NovaMjera.Text = "limenke"
            End If

            If StaraMjera = "listić" Or StaraMjera = "listića" Then
                If Kolicina = 1 Then NovaMjera.Text = "listić"
                If Kolicina > 1 And Kolicina < 5 Then NovaMjera.Text = "listića"
                If Kolicina >= 5 Then NovaMjera.Text = "listića"
                If Kolicina < 1 Then NovaMjera.Text = "listića"
            End If




        End With
    End Sub
End Module
