﻿Module DanUTjednu1Modul
    Sub DanUTjednu1()
        On Error Resume Next
        With Form1
            Dim DanUTjednuEng As Integer
            DanUTjednuEng = .DateTimePicker1.Value.Date.DayOfWeek
            Dim Dan As String = ""

            If DanUTjednuEng = 0 Then
                Dan = "nedjelju"
            End If
            If DanUTjednuEng = 1 Then
                Dan = "ponedjeljak"
            End If
            If DanUTjednuEng = 2 Then
                Dan = "utorak"
            End If
            If DanUTjednuEng = 3 Then
                Dan = "srijedu"
            End If
            If DanUTjednuEng = 4 Then
                Dan = "četvrtak"
            End If
            If DanUTjednuEng = 5 Then
                Dan = "petak"
            End If
            If DanUTjednuEng = 6 Then
                Dan = "subotu"
            End If
          

            .TextBox13.Text = "Jelovnik za " & Dan & ", " & .DateTimePicker1.Value.Date   'dan u tjednu

        End With
    End Sub

End Module
