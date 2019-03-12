Module PracenjeStanjaGrafModul
    Sub PracenjeStanjaGraf()
        On Error Resume Next
        With Form1

            Dim i As Integer

            .GroupBox49.Text = .TextBox14.Text & " " & .TextBox15.Text   'korisnik

            .Chart1.Series(0).Points.Clear()
            .Chart1.Series(1).Points.Clear()
            .Chart1.Series(2).Points.Clear()

            .Chart1.Series(0).Name = .ComboBox9.Text   'Series name
            .Chart1.Series(1).Name = ""
            .Chart1.Series(2).Name = ""

            .Chart1.ChartAreas(0).AxisX.MajorGrid.LineColor = Color.LightSkyBlue
            .Chart1.ChartAreas(0).AxisY.MajorGrid.LineColor = Color.LightSkyBlue

            ' .Chart1.ChartAreas(0).AxisX.MinorGrid.LineColor = Color.LightGray
            ' .Chart1.ChartAreas(0).AxisY.MinorGrid.LineColor = Color.LightGray

            .Chart1.Series(0).BorderWidth = 5
            .Chart1.Series(0).Color = Color.GreenYellow
            .Chart1.Series(1).BorderWidth = 2
            .Chart1.Series(1).Color = Color.BlueViolet
            .Chart1.Series(2).BorderWidth = 2
            .Chart1.Series(2).Color = Color.IndianRed

            Dim DGV As DataGridView = .DataGridView10
            Dim a As Integer = 6

            If .ComboBox9.Text = "Visina (cm)" Then
                a = 5
                .Chart1.ChartAreas(0).AxisY.Minimum = Val(.TextBox16.Text) - 10
                .Chart1.ChartAreas(0).AxisY.Maximum = Val(.TextBox16.Text) + 10
                .Chart1.Series(1).Name = ""
                .Chart1.Series(2).Name = ""
            End If

            If .ComboBox9.Text = "Masa (kg)" Then
                a = 6
                .Chart1.ChartAreas(0).AxisY.Minimum = DGV.Rows(0).Cells(11).Value - 10
                .Chart1.ChartAreas(0).AxisY.Maximum = DGV.Rows(0).Cells(12).Value + 40
                .Chart1.Series(1).Name = "Donja primjerena granica (kg)"
                .Chart1.Series(2).Name = "Gornja primjerena granica (kg)"
                .Chart1.Series(1).IsVisibleInLegend = True
                .Chart1.Series(2).IsVisibleInLegend = True
                'primjerena masa
                For i = 0 To DGV.RowCount - 1
                    .Chart1.Series(1).Points.AddY(DGV.Rows(i).Cells(11).Value)
                    .Chart1.Series(2).Points.AddY(DGV.Rows(i).Cells(12).Value)
                Next i
            Else
                .Chart1.Series(1).Name = ""
                .Chart1.Series(2).Name = ""
                .Chart1.Series(1).IsVisibleInLegend = False
                .Chart1.Series(2).IsVisibleInLegend = False
            End If

            If .ComboBox9.Text = "Opseg struka (cm)" Then
                a = 7
                .Chart1.ChartAreas(0).AxisY.Minimum = Val(.TextBox18.Text) - 10
                .Chart1.ChartAreas(0).AxisY.Maximum = Val(.TextBox18.Text) + 10
                .Chart1.Series(1).Name = ""
                .Chart1.Series(2).Name = ""
            End If

            If .ComboBox9.Text = "Opseg bokova (cm)" Then
                a = 8
                .Chart1.ChartAreas(0).AxisY.Minimum = Val(.TextBox63.Text) - 10
                .Chart1.ChartAreas(0).AxisY.Maximum = Val(.TextBox63.Text) + 10
                .Chart1.Series(1).Name = ""
                .Chart1.Series(2).Name = ""
            End If

            If .ComboBox9.Text = "WHR (omjer opsega struka i bokova)" Then
                a = 9
                .Chart1.ChartAreas(0).AxisY.Minimum = 0.8
                .Chart1.ChartAreas(0).AxisY.Maximum = 1.2
                .Chart1.Series(1).Name = ""
                .Chart1.Series(2).Name = ""
            End If
            If .ComboBox9.Text = "BMI (kg/m2)" Then
                a = 10
                For i = 0 To DGV.RowCount - 1
                    .Chart1.Series(1).Points.AddY(18.5)
                    .Chart1.Series(2).Points.AddY(25)
                Next i
                .Chart1.ChartAreas(0).AxisY.Minimum = 12
                .Chart1.ChartAreas(0).AxisY.Maximum = 41
                .Chart1.Series(1).IsVisibleInLegend = True
                .Chart1.Series(2).IsVisibleInLegend = True
                .Chart1.Series(1).Name = "Donja primjerena granica (kg/m2)"
                .Chart1.Series(2).Name = "Gornja primjerena granica (kg/m2)"
            End If

            For i = 0 To DGV.RowCount - 1
                .Chart1.Series(0).Points.AddXY(Format(DGV.Rows(i).Cells(13).Value, "dd.MM.yyyy"), DGV.Rows(i).Cells(a).Value)
            Next i

        End With

    End Sub
End Module
