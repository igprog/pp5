Public Class SortiranjeIspisaForm

    Private Sub SortiranjeIspisaForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.ListBox1.Items.Clear()
        'ListBox1.Items.AddRange(Split(textbox1.text, vbCrLf))
        '    Me.ListBox1.Items.AddRange(Form1.TextBox69.Text.ToString.Split(vbNewLine.ToString))
        Me.ListBox1.Items.AddRange(Form1.TextBox69.Text.Split(vbNewLine))
        '  Me.ListBox1.Items.AddRange(Form1.RichTextBoxPrintCtrl1.Text.Split(vbNewLine))

        'Listbox1.items.add(textbox1.text)
        '  Me.ListBox1.Items.Add(Form1.TextBox69.Text)
        '      Dim NewArray() As String = Form1.TextBox69.Text.Split(vbNewLine)
        '     For Each item As String In NewArray
        'ListBox1.Items.Add(item)
        'Next

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        On Error Resume Next
        With Me
            Dim index As Integer = .ListBox1.SelectedIndex
            If index = 0 OrElse index = -1 Then Return ' cannot move up anymore!
            Dim item As Object = .ListBox1.Items(index)
            .ListBox1.Items.RemoveAt(index)
            .ListBox1.Items.Insert(index - 1, item)
            .ListBox1.SelectedIndex = index - 1
        End With

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        On Error Resume Next
        With Me
            Dim index As Integer = .ListBox1.SelectedIndex
            If index >= .ListBox1.Items.Count - 1 Then Exit Sub
            Dim item As Object = .ListBox1.Items(index)
            .ListBox1.Items.RemoveAt(index)
            .ListBox1.Items.Insert(index + 1, item)
            .ListBox1.SelectedIndex = index + 1
        End With

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        On Error Resume Next

        Dim a As Integer
        Dim Ispis As TextBox = Form1.TextBox69
        Ispis.Text = ""
        For a = 0 To Me.ListBox1.Items.Count - 1
            If Ispis.Text = "" Then
                Ispis.Text = Me.ListBox1.Items(a).ToString
            Else
                Ispis.Text = Ispis.Text.ToString & vbCrLf & Me.ListBox1.Items(a).ToString
                Form1.RichTextBoxPrintCtrl1.Text = Ispis.Text
                '  Form1.TextBox69.Text = Form1.TextBox69.Text.ToString & Me.ListBox1.Items(a).ToString
            End If
        Next

       'Form1.RichTextBoxPrintCtrl1.Text = Ispis.Text

        'RICH TEXT BOX 1
        With Form1
            .RichTextBoxPrintCtrl1.Clear()
            ' .RichTextBox1.Text = .TextBox69.Text    'rich text box
            '  .RichTextBoxPrintCtrl1.Text = .TextBox69.Text    'rich text box
            .RichTextBoxPrintCtrl1.Text = Ispis.Text    'rich text box
            .RichTextBoxPrintCtrl1.Lines = .RichTextBoxPrintCtrl1.Text.Split(New Char() {ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)


            ' Dim Rtb As RichTextBox = .RichTextBox1
            Dim Rtb As RichTextBox = .RichTextBoxPrintCtrl1

             Dim Dorucak As String = .TabPage7.Text
            '  Rtb.Text = Replace(Rtb.Text, .TabPage7.Text, vbCrLf & .TabPage7.Text)
            Rtb.Select(Rtb.Text.IndexOf(Dorucak), Dorucak.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim JutarnjaUzina As String = .TabPage8.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage8.Text, vbCrLf & .TabPage8.Text)
            Rtb.Select(Rtb.Text.IndexOf(JutarnjaUzina), JutarnjaUzina.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim Rucak As String = .TabPage9.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage9.Text, vbCrLf & .TabPage9.Text)
            Rtb.Select(Rtb.Text.IndexOf(Rucak), Rucak.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim PopodnevnaUzina As String = .TabPage10.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage10.Text, vbCrLf & .TabPage10.Text)
            Rtb.Select(Rtb.Text.IndexOf(PopodnevnaUzina), PopodnevnaUzina.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim Vecera As String = .TabPage11.Text
            '   Rtb.Text = Replace(Rtb.Text, .TabPage11.Text, vbCrLf & .TabPage11.Text)
            Rtb.Select(Rtb.Text.IndexOf(Vecera), Vecera.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim ObrokPredSpavanje As String = .TabPage12.Text
            '    Rtb.Text = Replace(Rtb.Text, .TabPage12.Text, vbCrLf & .TabPage12.Text)
            Rtb.Select(Rtb.Text.IndexOf(ObrokPredSpavanje), ObrokPredSpavanje.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim UkupneVrijednosti As String = "UKUPNE VRIJEDNOSTI JELOVNIKA:"
            '   Rtb.Text = Replace(Rtb.Text, "UKUPNE VRIJEDNOSTI JELOVNIKA:", vbCrLf & "UKUPNE VRIJEDNOSTI JELOVNIKA:")
            Rtb.Select(Rtb.Text.IndexOf(UkupneVrijednosti), UkupneVrijednosti.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim CijenaJelovnika As String = "Cijena jelovnika:"
            '   Rtb.Text = Replace(Rtb.Text, "Cijena jelovnika:", vbCrLf & "Cijena jelovnika:")
            Rtb.Select(Rtb.Text.IndexOf(CijenaJelovnika), CijenaJelovnika.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)

            Dim DodatnaTjelesnaAktivnost As String = "DODATNA TJELESNA AKTIVNOST:"
            '  Rtb.Text = Replace(Rtb.Text, "DODATNA TJELESNA AKTIVNOST:", vbCrLf & "DODATNA TJELESNA AKTIVNOST:")
            Rtb.Select(Rtb.Text.IndexOf(DodatnaTjelesnaAktivnost), DodatnaTjelesnaAktivnost.Length)
            Rtb.SelectionFont = New Font("Arial", 10, FontStyle.Bold)


            'Namirnice italic
            '  Dim txtSelection As RichTextBox = Me.RichTextBox1
            Dim txtSearch As String = "Namirnice:"
            Dim textEnd As Integer = Rtb.TextLength
            Dim index As Integer = 0
            Dim lastIndex As Integer = Rtb.Text.LastIndexOf(txtSearch)

            Dim myStyle As FontStyle
            myStyle = FontStyle.Bold + FontStyle.Italic
            '  Me.RichTextBox1.SelectionFont = New Font("Arial", 8, myStyle)

            While index < lastIndex
                Rtb.Find(txtSearch, index, textEnd, RichTextBoxFinds.None)
                'txtSelection.SelectionBackColor = Color.Yellow
                Rtb.SelectionFont = New Font("Arial", 8, myStyle)
                index = Rtb.Text.IndexOf(txtSearch, index) + 1
            End While


            'Prazno polje
            Dim txtSearch1 As String = "~"
            Dim textEnd1 As Integer = Rtb.TextLength
            Dim index1 As Integer = 0
            Dim lastIndex1 As Integer = Rtb.Text.LastIndexOf(txtSearch1)
            While index1 < lastIndex1
                Rtb.Find(txtSearch1, index1, textEnd1, RichTextBoxFinds.None)
                Rtb.SelectionColor = Color.White
                index1 = Rtb.Text.IndexOf(txtSearch1, index1) + 1
            End While

        End With

        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Close()

    End Sub
End Class