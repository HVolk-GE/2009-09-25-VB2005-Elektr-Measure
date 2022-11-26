Public Class Form5

    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "MessungenDataSet.DuplTestNumbersAndTimes". Sie können sie bei Bedarf verschieben oder entfernen.
        Me.DuplTestNumbersAndTimesTableAdapter.Fill(Me.MessungenDataSet.DuplTestNumbersAndTimes)

    End Sub
End Class