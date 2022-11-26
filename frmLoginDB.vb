Public Class frmLoginDB


    Private Sub frmLoginDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "MessungenDataSet.Template". Sie können sie bei Bedarf verschieben oder entfernen.
        '*Me.TemplateTableAdapter.Fill(Me.MessungenDataSet.Template)
        'TODO: Diese Codezeile lädt Daten in die Tabelle "MessungenDataSet.VorlagenFormat". Sie können sie bei Bedarf verschieben oder entfernen.
        '*Me.VorlagenFormatTableAdapter.Fill(Me.MessungenDataSet.VorlagenFormat)

        Dim i As Integer

        Me.ToolStripStatusLabel1.Text = "created by H. Volk"

        CntTmpName = 0 : CntTestNr = 0 : cntBack0 = 0

        con.ConnectionString = strCon

        sqladapterVor = New OleDb.OleDbDataAdapter("Select * from VorlagenFormat", con)
        sqladapterVor0 = New OleDb.OleDbDataAdapter("Select * from VorlagenFormat", con)
        sqladapterTemp = New OleDb.OleDbDataAdapter("Select * from Template", con)
        sqladapterTemp0 = New OleDb.OleDbDataAdapter("Select * from Template", con)
        'sqlAdapterCombo1 = New OleDb.OleDbDataAdapter("Select Testnumber from Template", con)
        sqlAdapterCombo1 = New OleDb.OleDbDataAdapter("Select * from DuplTemplateName", con)
        sqlAdapterCombo2 = New OleDb.OleDbDataAdapter("Select Templatename from VorlagenFormat", con)
        sqlAdapterCntTimes = New OleDb.OleDbDataAdapter("Select * from DuplTestNumbersAndTimes", con)

        con.Open()

        sqlAdapterCombo1.Fill(DS0)
        sqlAdapterCombo2.Fill(DS1)
        sqladapterVor.Fill(Me.MessungenDataSet.VorlagenFormat)
        sqladapterVor0.Fill(DS2)
        sqladapterTemp.Fill(Me.MessungenDataSet.Template)
        sqladapterTemp0.Fill(DS3)
        sqlAdapterCntTimes.Fill(DS4)

        Me.ComboBox1.Items.Clear()
        Me.ComboBox2.Items.Clear()

        '* Frage was schonmal gemessen wurde :
        For i = 0 To DS4.Tables.Item(0).Rows.Count - 1
            ReDim Preserve tstArrayNumbers(i)
            ReDim Preserve tstArrayCnt(i)
            ReDim Preserve tstArrayTimes(i)
            ReDim Preserve tstArrayPadInOut(i)
            ReDim Preserve tstArrayTemplatenam(i)
            tstArrayNumbers(i) = DS4.Tables.Item(0).Rows(i).Item(0).ToString  ' Testnummer
            tstArrayCnt(i) = DS4.Tables.Item(0).Rows(i).Item(1).ToString      ' Anzahl (1)
            tstArrayTimes(i) = DS4.Tables.Item(0).Rows(i).Item(2).ToString    ' Messungszeitpunkt
            tstArrayPadInOut(i) = DS4.Tables.Item(0).Rows(i).Item(3).ToString ' Pad In/Out
            tstArrayTemplatenam(i) = DS4.Tables.Item(0).Rows(i).Item(4).ToString ' Templatename
        Next

        For i = 0 To DS2.Tables.Item(0).Rows.Count - 1
            ReDim Preserve orgArrayTemplatenam(i)
            ReDim Preserve orgArrayPadTimes(i)
            ReDim Preserve orgArrayDiscTimes(i)
            orgArrayTemplatenam(i) = DS2.Tables.Item(0).Rows(i).Item(0).ToString()
            orgArrayPadTimes(i) = DS2.Tables.Item(0).Rows(i).Item(3).ToString()
            orgArrayDiscTimes(i) = DS2.Tables.Item(0).Rows(i).Item(4).ToString()
        Next

        Dim txtTestnumbers0 As String, ConterTime, xx, xy, vwPad, vwDisc As Integer
        Dim TmpTstNumber() As String

        txtTestnumbers0 = ""
        ConterTime = 0
        vwPad = 0
        vwDisc = 0
        xx = 0

        For i = 0 To DS4.Tables.Item(0).Rows.Count - 1
            If txtTestnumbers0 = "" Then
                txtTestnumbers0 = tstArrayNumbers(i)
            End If
            If txtTestnumbers0 <> "" Then
                If txtTestnumbers0 = tstArrayNumbers(i) Then
                    ReDim Preserve TmpTstNumber(xx)
                    ConterTime = ConterTime + 1
                    If vwPad = 0 Or vwDisc = 0 Then
                        For xy = 0 To DS2.Tables.Item(0).Rows.Count - 1
                            If tstArrayTemplatenam(i) = orgArrayTemplatenam(xy) Then
                                vwPad = CInt(orgArrayPadTimes(xy))
                                vwDisc = CInt(orgArrayDiscTimes(xy))
                                vwPad = vwPad * 2
                                vwDisc = vwDisc
                                Exit For
                            End If
                        Next
                    End If
                    If tstArrayPadInOut(i) <> "" Then
                        If ConterTime >= vwPad Then
                            TmpTstNumber(xx) = tstArrayNumbers(i)
                            xx = xx + 1
                            txtTestnumbers0 = ""
                            ConterTime = 0
                            vwPad = 0
                            vwDisc = 0
                        End If
                    ElseIf tstArrayPadInOut(i) = "" Then
                        If ConterTime >= vwDisc Then
                            TmpTstNumber(xx) = tstArrayNumbers(i)
                            xx = xx + 1
                            txtTestnumbers0 = ""
                            ConterTime = 0
                            vwPad = 0
                            vwDisc = 0
                        End If
                    End If
                ElseIf txtTestnumbers0 <> tstArrayNumbers(i) Then
                    txtTestnumbers0 = ""
                    ConterTime = 0
                    vwPad = 0
                    vwDisc = 0
                End If
            End If
        Next
        '###
        For i = 0 To DS0.Tables.Item(0).Rows.Count - 1
            If xx > 0 Then
                For xy = 0 To xx - 1
                    If DS0.Tables.Item(0).Rows(i).Item(0).ToString <> TmpTstNumber(xy) Then
                        Me.ComboBox1.Items.Add(DS0.Tables.Item(0).Rows(i).Item(0).ToString)
                        CntTmpName = i
                    End If
                Next
            ElseIf xx = 0 Then
                Me.ComboBox1.Items.Add(DS0.Tables.Item(0).Rows(i).Item(0).ToString)
                CntTmpName = i
            End If
        Next

        For i = 0 To DS1.Tables.Item(0).Rows.Count - 1
            Me.ComboBox2.Items.Add(DS1.Tables.Item(0).Rows(i).Item(0).ToString)
            CntTestNr = i
        Next i

        IniTal()

        If State = "DE" Then
            Me.Text = "Anmelden"
            Me.Label1.Text = "Name :"
            Me.Label2.Text = "Test Nummer :"
            Me.Label3.Text = "Template Name :"
            Me.CheckBox1.Text = "Neu Anlegen"
            Me.Button1.Text = "Ok"
            Me.Button2.Text = "Abbrechen"
        Else
            Me.Text = "Login"
            Me.Label1.Text = "Initials :"
            Me.Label2.Text = "Test Number :"
            Me.Label3.Text = "Template Name :"
            Me.CheckBox1.Text = "Create New"
            Me.Button1.Text = "Ok"
            Me.Button2.Text = "Close"
        End If

        'If kopieview = "True" Then
        Me.CheckBox1.Visible = True
        'Else
        'Me.CheckBox1.Visible = False
        'End If

        Me.Label3.Visible = False
        Me.ComboBox2.Visible = False
        Me.ComboBox1.Enabled = False
        Me.CheckBox1.Enabled = False

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Dim i As Integer

        If Me.CheckBox1.Checked = True Then
            Me.Label3.Visible = True
            Me.ComboBox1.Items.Clear()
            DBTestnr = Me.ComboBox1.Text
            Me.ComboBox2.Visible = True
            'DBTemplateName = Me.ComboBox2.Text
        Else
            Me.Label3.Visible = False

            For i = 0 To DS0.Tables.Item(0).Rows.Count - 1
                Me.ComboBox1.Items.Add(DS0.Tables.Item(0).Rows(i).Item(0).ToString)
            Next

            Me.ComboBox2.Visible = False

        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        cntBack0 = 0
        Usernam = Me.TextBox1.Text
        strUsername = Me.TextBox1.Text
        DBTestnr = ""
        DBTemplateName = ""

        If Me.CheckBox1.Checked = True Then
            NewTest = "Y"
            DBTemplateName = Me.ComboBox2.Text
            SearchCounts()
            DBTestnr = Me.ComboBox1.Text
        Else
            DBTestnr = Me.ComboBox1.Text
            SearchCounts()
            NewTest = "N"
        End If

        Me.Hide()
        Form2.Show()

    End Sub

    Private Sub VorlagenFormatBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VorlagenFormatBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.VorlagenFormatBindingSource.EndEdit()
        Me.VorlagenFormatTableAdapter.Update(Me.MessungenDataSet.VorlagenFormat)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

        If Len(Me.TextBox1.Text) = 2 Then

            Me.ComboBox1.Enabled = True
            Me.CheckBox1.Enabled = True
            Me.TextBox1.Text = UCase(Me.TextBox1.Text)

        ElseIf Len(Me.TextBox1.Text) > 3 Then

            If State = "DE" Then
                MsgBox("Bitte, initialien eingeben (2 Buchstaben).", MsgBoxStyle.Information, "Eingabe fehlt !")
                Exit Sub
            Else
                MsgBox("Please, insert you initials (2 Characters).", MsgBoxStyle.Information, "Failed insert !")
                Exit Sub
            End If

        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        MainMenue.Show()
    End Sub

    Private Sub frmLoginDB_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        Me.TextBox1.Text = ""
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""

        If cntBack0 = 2 Then

            Me.Validate()
            Me.TemplateBindingSource.EndEdit()

            'Me.VorlagenFormatBindingSource.EndEdit()

            UpdateData()

            Me.MessungenDataSet.AcceptChanges()

            cntBack0 = 0

        End If
    End Sub

    Private Sub UpdateData()

        Me.Validate()
        Me.TemplateBindingSource.EndEdit()
        'Me.VorlagenFormatBindingSource.EndEdit()

        Using updateTransaction As New Transactions.TransactionScope

            '        DeleteOrders()
            AddNewOrders()

            updateTransaction.Complete()

            Me.MessungenDataSet.AcceptChanges()

        End Using
    End Sub

    Private Sub AddNewOrders()
        Dim newOrders As MessungenDataSet.TemplateDataTable
        newOrders = CType(MessungenDataSet.Template.GetChanges(Data.DataRowState.Added), _
             MessungenDataSet.TemplateDataTable)

        If Not IsNothing(newOrders) Then
            Try
                'sqladapterVor.Update(newOrders)
                'VorlagenFormatTableAdapter.Update(newOrders)
                TemplateTableAdapter.Update(newOrders)
                'VorlagenFormatTableAdapter.Insert(newOrders)

            Catch ex As Exception
                MessageBox.Show("AddNew Failed" & vbCrLf & ex.Message)
            End Try
        End If
    End Sub

End Class