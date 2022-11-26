Public Class Form3

    Private Sub TemplateBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TemplateBindingNavigatorSaveItem.Click
        '        Me.Validate()
        '        Me.TemplateBindingSource.EndEdit()
        '        Me.TemplateTableAdapter.Update(Me.MessungenDataSet.Template)
        UpdateData()

    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "MessungenDataSet.Template". Sie können sie bei Bedarf verschieben oder entfernen.
        Me.TemplateTableAdapter.Fill(Me.MessungenDataSet.Template)
    End Sub

    Private Sub UpdateData()
        Me.Validate()
        Me.TemplateBindingSource.EndEdit()

        Using updateTransaction As New Transactions.TransactionScope

            DeleteOrders()
            AddNewOrders()

            updateTransaction.Complete()
            MessungenDataSet.AcceptChanges()

        End Using
    End Sub

    Private Sub DeleteOrders()
        Dim deletedOrders As MessungenDataSet.TemplateDataTable
        deletedOrders = CType(MessungenDataSet.Template.GetChanges(Data.DataRowState.Deleted), _
             MessungenDataSet.TemplateDataTable)


        If Not IsNothing(deletedOrders) Then
            Try

                TemplateTableAdapter.Update(deletedOrders)

            Catch ex As Exception
                MessageBox.Show("Delete Failed")
            End Try
        End If
    End Sub

    Private Sub AddNewOrders()

        Dim newOrders As MessungenDataSet.TemplateDataTable
        newOrders = CType(MessungenDataSet.Template.GetChanges(Data.DataRowState.Added), _
             MessungenDataSet.TemplateDataTable)

        If Not IsNothing(newOrders) Then
            Try
                TemplateTableAdapter.Update(newOrders)

            Catch ex As Exception
                MessageBox.Show("AddNew Failed" & vbCrLf & ex.Message)
            End Try
        End If
    End Sub

End Class