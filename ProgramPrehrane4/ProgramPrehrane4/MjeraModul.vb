Module MjeraModul
    Sub Mjera()
        With Form1
            If .Label230.Text = "jušna žlica" Then
                If .TextBox12.Text <= 1 Then .Label103.Text = "jušna žlica"
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "jušne žlice"
                If .TextBox12.Text >= 5 Then .Label103.Text = "jušnih žlica"
                If .TextBox12.Text < 1 Then .Label103.Text = "jušne žlice"
            End If

            If .Label230.Text = "šalica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "šalice"
                If .TextBox12.Text = 1 Or .TextBox12.Text >= 5 Then .Label103.Text = "šalica"
                If .TextBox12.Text < 1 Then .Label103.Text = "šalice"
            End If

            If .Label230.Text = "plod" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "ploda"
                If .TextBox12.Text >= 5 Then .Label103.Text = "plodova"
                If .TextBox12.Text = 1 Then .Label103.Text = "plod"
                If .TextBox12.Text < 1 Then .Label103.Text = "ploda"
            End If

            If .Label230.Text = "čajna žličica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "čajne žličice"
                If .TextBox12.Text >= 5 Or .TextBox12.Text < 1 Then .Label103.Text = "čajnih žličica"
                If .TextBox12.Text = 1 Then .Label103.Text = "čajna žličica"
                If .TextBox12.Text < 1 Then .Label103.Text = "čajne žličice"
            End If

            If .Label230.Text = "porcija" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "porcije"
                If .TextBox12.Text >= 5 Then .Label103.Text = "porcija"
                If .TextBox12.Text = 1 Then .Label103.Text = "porcija"
                If .TextBox12.Text < 1 Then .Label103.Text = "porcije"
            End If

            If .Label230.Text = "limenka" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "limenke"
                If .TextBox12.Text >= 5 Then .Label103.Text = "limenki"
                If .TextBox12.Text = 1 Then .Label103.Text = "limenka"
                If .TextBox12.Text < 1 Then .Label103.Text = "limenke"
            End If

            If .Label230.Text = "krug" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "kruga"
                If .TextBox12.Text >= 5 Then .Label103.Text = "krugova"
                If .TextBox12.Text = 1 Then .Label103.Text = "krug"
                If .TextBox12.Text < 1 Then .Label103.Text = "kruga"
            End If

            If .Label230.Text = "bobica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "bobice"
                If .TextBox12.Text >= 5 Then .Label103.Text = "bobica"
                If .TextBox12.Text = 1 Then .Label103.Text = "bobica"
                If .TextBox12.Text < 1 Then .Label103.Text = "bobice"
            End If

            If .Label230.Text = "kriška" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "kriške"
                If .TextBox12.Text >= 5 Then .Label103.Text = "kriški"
                If .TextBox12.Text = 1 Then .Label103.Text = "kriška"
                If .TextBox12.Text < 1 Then .Label103.Text = "kriške"
            End If

            If .Label230.Text = "boca" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "boce"
                If .TextBox12.Text >= 5 Then .Label103.Text = "boca"
                If .TextBox12.Text = 1 Then .Label103.Text = "boca"
                If .TextBox12.Text < 1 Then .Label103.Text = "boce"
            End If

            If .Label230.Text = "čaša" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "čaše"
                If .TextBox12.Text >= 5 Then .Label103.Text = "čaša"
                If .TextBox12.Text = 1 Then .Label103.Text = "čaša"
                If .TextBox12.Text < 1 Then .Label103.Text = "čaše"
            End If

            If .Label230.Text = "kockica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "kockice"
                If .TextBox12.Text >= 5 Then .Label103.Text = "kockica"
                If .TextBox12.Text = 1 Then .Label103.Text = "kockica"
                If .TextBox12.Text < 1 Then .Label103.Text = "kockice"
            End If

            If .Label230.Text = "loptica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "loptice"
                If .TextBox12.Text >= 5 Then .Label103.Text = "loptica"
                If .TextBox12.Text = 1 Then .Label103.Text = "loptica"
                If .TextBox12.Text < 1 Then .Label103.Text = "loptice"
            End If

            If .Label230.Text = "polovica" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "polovice"
                If .TextBox12.Text >= 5 Then .Label103.Text = "polovica"
                If .TextBox12.Text = 1 Then .Label103.Text = "polovica"
                If .TextBox12.Text < 1 Then .Label103.Text = "polovice"
            End If

            If .Label230.Text = "mali komad" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "mala komada"
                If .TextBox12.Text >= 5 Then .Label103.Text = "malih komada"
                If .TextBox12.Text = 1 Then .Label103.Text = "mali komad"
                If .TextBox12.Text < 1 Then .Label103.Text = "malog komada"
            End If

            If .Label230.Text = "listić" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "listića"
                If .TextBox12.Text >= 5 Then .Label103.Text = "listića"
                If .TextBox12.Text = 1 Then .Label103.Text = "listić"
                If .TextBox12.Text < 1 Then .Label103.Text = "listića"
            End If

            If .Label230.Text = "serviranje" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "serviranja"
                If .TextBox12.Text >= 5 Then .Label103.Text = "serviranja"
                If .TextBox12.Text = 1 Then .Label103.Text = "serviranje"
                If .TextBox12.Text < 1 Then .Label103.Text = "serviranja"
            End If

            If .Label230.Text = "zrno" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "zrna"
                If .TextBox12.Text >= 5 Then .Label103.Text = "zrna"
                If .TextBox12.Text = 1 Then .Label103.Text = "zrno"
                If .TextBox12.Text < 1 Then .Label103.Text = "zrna"
            End If

            If .Label230.Text = "veliki plod" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "velika ploda"
                If .TextBox12.Text >= 5 Then .Label103.Text = "velikih plodova"
                If .TextBox12.Text = 1 Then .Label103.Text = "veliki plod"
                If .TextBox12.Text < 1 Then .Label103.Text = "velikog ploda"
            End If

            If .Label230.Text = "velika" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "velike"
                If .TextBox12.Text >= 5 Then .Label103.Text = "velikih"
                If .TextBox12.Text = 1 Then .Label103.Text = "velika"
                If .TextBox12.Text < 1 Then .Label103.Text = "velike"
            End If

            If .Label230.Text = "veći plod" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "veća ploda"
                If .TextBox12.Text >= 5 Then .Label103.Text = "većih plodova"
                If .TextBox12.Text = 1 Then .Label103.Text = "veći plod"
                If .TextBox12.Text < 1 Then .Label103.Text = "većeg ploda"
            End If

            If .Label230.Text = "mali plod" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "mala ploda"
                If .TextBox12.Text >= 5 Then .Label103.Text = "malih plodova"
                If .TextBox12.Text = 1 Then .Label103.Text = "mali plod"
                If .TextBox12.Text < 1 Then .Label103.Text = "malog ploda"
            End If

            If .Label230.Text = "srednji plod" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "srednja ploda"
                If .TextBox12.Text >= 5 Then .Label103.Text = "srednjih plodova"
                If .TextBox12.Text = 1 Then .Label103.Text = "srednji plod"
                If .TextBox12.Text < 1 Then .Label103.Text = "srednjeg ploda"
            End If

            If .Label230.Text = "veliki komad" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "velika komada"
                If .TextBox12.Text >= 5 Then .Label103.Text = "velikih komada"
                If .TextBox12.Text = 1 Then .Label103.Text = "veliki komad"
                If .TextBox12.Text < 1 Then .Label103.Text = "velikog komada"
            End If

            If .Label230.Text = "mala kriška" Then
                If .TextBox12.Text > 1 And .TextBox12.Text < 5 Then .Label103.Text = "male kriške"
                If .TextBox12.Text >= 5 Then .Label103.Text = "malih kriški"
                If .TextBox12.Text = 1 Then .Label103.Text = "mala kriška"
                If .TextBox12.Text < 1 Then .Label103.Text = "male kriške"
            End If


            If .Label230.Text = "komad" Then
                If .TextBox12.Text = 1 Then
                    .Label103.Text = "komad"
                Else
                    .Label103.Text = "komada"
                End If
            End If

            If .Label230.Text = "gram" Then
                If .TextBox12.Text = 1 Then
                    .Label103.Text = "gram"
                Else
                    .Label103.Text = "grama"
                End If
            End If

            '   If .Label230.Text = "mala kriška" Then
            'If .TextBox12.Text = 1 Then
            '.Label103.Text = "mala kriška"
            '    Else
            '   .Label103.Text = "male kriške"
            '  End If
            ' End If

            If .Label230.Text = "prutić" Then
                If .TextBox12.Text = 1 Then
                    .Label103.Text = "prutić"
                Else
                    .Label103.Text = "prutića"
                End If
            End If


        End With
    End Sub
End Module
