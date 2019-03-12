Module NutrijentiPreporuceniPostociModul
    Sub NutrijentiPreporuceniPostoci()
        With Form1
            Dim Umin As Integer = .Label88.Text   'ugljikohidrati min
            Dim Umax As Integer = .Label91.Text   'ugljikohidrati max
            Dim Bmin As Integer = .Label89.Text   'bjelancevine min
            Dim Bmax As Integer = .Label92.Text   'bjelancevine max
            Dim Mmin As Integer = .Label90.Text   'masti min
            Dim Mmax As Integer = .Label93.Text   'masti max

            .ListBox9.Items.Clear()    'briši list boks
            .ListBox9.Items.Insert(0, .TextBox3.Text & " kcal")   'energija kcal
            .ListBox9.Items.Insert(1, "")   'prazan red
            .ListBox9.Items.Insert(2, Umin & "-" & Umax & "%")   'ugljikohidrati
            .ListBox9.Items.Insert(3, Bmin & "-" & Bmax & "%")  'bjelancevine
            .ListBox9.Items.Insert(4, Mmin & "-" & Mmax & "%")   'masti
            '      .ListBox9.Items.Insert(1, "")   'prazan red
            '     .ListBox9.Items.Insert(2, .Label83.Text)   'ugljikohidrati
            '    .ListBox9.Items.Insert(3, .Label84.Text)  'bjelancevine
            '   .ListBox9.Items.Insert(4, .Label85.Text)   'masti

        End With
    End Sub
End Module
