Public Class Form5

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile l�dt Daten in die Tabelle "MessungenDataSet.DuplTestNumbersAndTimes". Sie k�nnen sie bei Bedarf verschieben oder entfernen.
        Me.DuplTestNumbersAndTimesTableAdapter.Fill(Me.MessungenDataSet.DuplTestNumbersAndTimes)

    End Sub
End Class