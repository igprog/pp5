Module PrepBrojServGlikogenModul
    Sub PrepBrojServGlikogen()
        On Error Resume Next
        With Form1
            Dim TEE As Integer = .TextBox3.Text
            Dim K As Integer = 1   'broj korisnika jelovnika
            '  TEE = .Label81.Text
            .ListBox2.Items.Clear()

            If TEE = 0 Then
                .ListBox2.Items.Insert(0, 0)   'Žitarice"
                .ListBox2.Items.Insert(1, 0)  'Povrće
                .ListBox2.Items.Insert(2, 0)   'Voće
                .ListBox2.Items.Insert(3, 0)   'Meso
                .ListBox2.Items.Insert(4, 0) 'Mlijeko
                .ListBox2.Items.Insert(5, 0)   'Masti
            End If

            If TEE > 0 And TEE <= 1900 Then
                .ListBox2.Items.Insert(0, 17 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 2 * K)  'Povrće
                .ListBox2.Items.Insert(2, 2 * K)   'Voće
                .ListBox2.Items.Insert(3, 0 * K)   'Meso
                .ListBox2.Items.Insert(4, 1 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 3 * K)   'Masti
            End If

            If TEE > 1900 And TEE <= 2100 Then
                .ListBox2.Items.Insert(0, 17 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 4 * K)  'Povrće
                .ListBox2.Items.Insert(2, 3 * K)   'Voće
                .ListBox2.Items.Insert(3, 0 * K)   'Meso
                .ListBox2.Items.Insert(4, 2 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 3 * K)   'Masti
            End If

            If TEE > 2100 And TEE <= 2300 Then
                .ListBox2.Items.Insert(0, 17 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 5 * K)  'Povrće
                .ListBox2.Items.Insert(2, 4 * K)   'Voće
                .ListBox2.Items.Insert(3, 1 * K)   'Meso
                .ListBox2.Items.Insert(4, 2 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 4 * K)   'Masti
            End If

            If TEE > 2300 And TEE <= 2500 Then
                .ListBox2.Items.Insert(0, 18 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 5 * K)  'Povrće
                .ListBox2.Items.Insert(2, 5 * K)   'Voće
                .ListBox2.Items.Insert(3, 1 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 4 * K)   'Masti
            End If

            If TEE > 2500 And TEE <= 2700 Then
                .ListBox2.Items.Insert(0, 18 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 5 * K)  'Povrće
                .ListBox2.Items.Insert(2, 5 * K)   'Voće
                .ListBox2.Items.Insert(3, 1 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 5 * K)   'Masti
            End If

            If TEE > 2700 And TEE <= 2900 Then
                .ListBox2.Items.Insert(0, 19 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 6 * K)  'Povrće
                .ListBox2.Items.Insert(2, 6 * K)   'Voće
                .ListBox2.Items.Insert(3, 2 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 5 * K)   'Masti
            End If

            If TEE > 2900 And TEE <= 3100 Then
                .ListBox2.Items.Insert(0, 19 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 6 * K)  'Povrće
                .ListBox2.Items.Insert(2, 7 * K)   'Voće
                .ListBox2.Items.Insert(3, 2 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 5 * K)   'Masti
            End If

            If TEE > 3100 And TEE <= 3300 Then
                .ListBox2.Items.Insert(0, 19 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 6 * K)  'Povrće
                .ListBox2.Items.Insert(2, 7 * K)   'Voće
                .ListBox2.Items.Insert(3, 3 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 6 * K)   'Masti
            End If

            If TEE > 3300 And TEE <= 3500 Then
                .ListBox2.Items.Insert(0, 19 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 7 * K)  'Povrće
                .ListBox2.Items.Insert(2, 7 * K)   'Voće
                .ListBox2.Items.Insert(3, 4 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 7 * K)   'Masti
            End If

            If TEE > 3500 And TEE <= 3700 Then
                .ListBox2.Items.Insert(0, 20 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 7 * K)  'Povrće
                .ListBox2.Items.Insert(2, 7 * K)   'Voće
                .ListBox2.Items.Insert(3, 4 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 7 * K)   'Masti
            End If

            If TEE > 3700 And TEE <= 3900 Then
                .ListBox2.Items.Insert(0, 21 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 8 * K)  'Povrće
                .ListBox2.Items.Insert(2, 7 * K)   'Voće
                .ListBox2.Items.Insert(3, 4 * K)   'Meso
                .ListBox2.Items.Insert(4, 3 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 7 * K)   'Masti
            End If

            If TEE > 3900 Then
                .ListBox2.Items.Insert(0, 22 * K)   'Žitarice"
                .ListBox2.Items.Insert(1, 9 * K)  'Povrće
                .ListBox2.Items.Insert(2, 8 * K)   'Voće
                .ListBox2.Items.Insert(3, 4 * K)   'Meso
                .ListBox2.Items.Insert(4, 4 * K) 'Mlijeko
                .ListBox2.Items.Insert(5, 7 * K)   'Masti
            End If

        End With
    End Sub
End Module
