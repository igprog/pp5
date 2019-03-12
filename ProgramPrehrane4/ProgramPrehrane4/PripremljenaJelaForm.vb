Public Class PripremljenaJelaForm


   

  

    Private Sub GotovaJelaForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PP5DataSet.PripremljenaJela' table. You can move, or remove it, as needed.
        Me.PripremljenaJelaTableAdapter.Fill(Me.PP5DataSet.PripremljenaJela)

    End Sub
End Class