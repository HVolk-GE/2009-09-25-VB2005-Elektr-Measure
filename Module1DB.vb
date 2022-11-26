Module Module1DB
    Public PadDisc, DBTestnr, DBTemplateName, NewTest, DBUser As String
    Public MeasurmentTimesPads, MesurementPointPads As Integer
    Public MeasurmentTimesDisc, MesurementPointDisc As Integer

    Public strCon As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Messungen.mdb"
    Public con As New OleDb.OleDbConnection
    Public DS0 As New DataSet
    Public DS1 As New DataSet
    Public DS2 As New DataSet
    Public DS3 As New DataSet
    Public DS4 As New DataSet
    Public cntBack0 As Integer

    Public sqladapterVor As OleDb.OleDbDataAdapter
    Public sqladapterVor0 As OleDb.OleDbDataAdapter
    Public sqladapterTemp As OleDb.OleDbDataAdapter
    Public sqladapterTemp0 As OleDb.OleDbDataAdapter
    Public sqladapterTemp1 As OleDb.OleDbDataAdapter
    Public sqlAdapterCombo1 As OleDb.OleDbDataAdapter
    Public sqlAdapterCombo2 As OleDb.OleDbDataAdapter
    Public sqlAdapterCntTimes As OleDb.OleDbDataAdapter

    Public tstArrayNumbers(), tstArrayCnt(), tstArrayTimes(), tstArrayPadInOut() As String
    Public tstArrayTemplatenam() As String
    Public orgArrayTemplatenam(), orgArrayPadTimes(), orgArrayDiscTimes() As String

    Public CntTmpName, CntTestNr, CntTests As Integer

    Sub SearchCounts()
        Dim i, inita0, cntY As Integer

        cntY = 0

        'If old Test with Number:
        If DBTestnr <> "" And frmLoginDB.CheckBox1.Checked = False Then

            For i = 0 To DS3.Tables.Item(0).Rows.Count - 1
                If DBTestnr = DS3.Tables.Item(0).Rows(i).Item(1).ToString() Then
                    ' Set Testnumber 
                    DBTemplateName = DS3.Tables.Item(0).Rows(i).Item(0).ToString()

                    PadDisc = DS3.Tables.Item(0).Rows(i).Item(3).ToString()

                    DiscID = DS3.Tables.Item(0).Rows(i).Item(16).ToString()

                    Runout = DS3.Tables.Item(0).Rows(i).Item(17).ToString()

                    ' Set Templatename
                    For inita0 = 0 To DS2.Tables.Item(0).Rows.Count - 1
                        If DBTemplateName = DS2.Tables.Item(0).Rows(inita0).Item(0).ToString() Then
                            'Set now Measurmentpoint of Pads and Disc
                            MesurementPointPads = DS2.Tables.Item(0).Rows(inita0).Item(1).ToString()
                            MesurementPointDisc = DS2.Tables.Item(0).Rows(inita0).Item(2).ToString()
                            ' Set now Measurment times for Pads and Disc
                            MeasurmentTimesPads = DS2.Tables.Item(0).Rows(inita0).Item(3).ToString()
                            MeasurmentTimesDisc = DS2.Tables.Item(0).Rows(inita0).Item(4).ToString()
                            cntY = 1
                            Exit For
                        End If
                    Next
                    ' If all settings, go exit this sub
                    If cntY = 1 Then Exit For
                End If
            Next

            ' ElseIf New Test without Testnumber:
        ElseIf DBTemplateName <> "" And frmLoginDB.CheckBox1.Checked = True Then

            DBTemplateName = DS3.Tables.Item(0).Rows(i).Item(0).ToString()
            ' Set Templatename
            For inita0 = 0 To DS2.Tables.Item(0).Rows.Count - 1
                If DBTemplateName = DS2.Tables.Item(0).Rows(inita0).Item(0).ToString() Then
                    'Set now Measurmentpoint of Pads and Disc
                    MesurementPointPads = DS2.Tables.Item(0).Rows(inita0).Item(1).ToString()
                    MesurementPointDisc = DS2.Tables.Item(0).Rows(inita0).Item(2).ToString()
                    ' Set now Measurment times for Pads and Disc
                    MeasurmentTimesPads = DS2.Tables.Item(0).Rows(inita0).Item(3).ToString()
                    MeasurmentTimesDisc = DS2.Tables.Item(0).Rows(inita0).Item(4).ToString()
                    cntY = 1
                    Exit For
                End If
            Next

        End If

    End Sub

    Sub ValuesToDB()
        Dim i As Integer

        Dim saveNow As DateTime = DateTime.Now
        Dim TmpTxt01, TmpTxt02, TmpTxt03, TmpTxt04, TmpTxt05, TmpTxt06 As String
        Dim TmpTxt07, TmpTxt08, TmpTxt09, TmpTxt10, TmpTxt11, TmpTxt12 As String
        Dim TmpTxt13, TmpTxt14, TmpTxt15, TmpTxt16, TmpTxt17, TmpTxt18 As String
        Dim TmpTxt19, TmpTxt20, TmpTxt21, TmpTxt22, TmpTxt23, TmpTxt24 As String

        TmpTxt01 = "" : TmpTxt02 = "" : TmpTxt03 = "" : TmpTxt04 = "" : TmpTxt05 = ""
        TmpTxt06 = "" : TmpTxt07 = "" : TmpTxt08 = "" : TmpTxt09 = "" : TmpTxt10 = ""
        TmpTxt11 = "" : TmpTxt12 = "" : TmpTxt13 = "" : TmpTxt14 = "" : TmpTxt15 = ""
        TmpTxt16 = "" : TmpTxt17 = "" : TmpTxt18 = "" : TmpTxt19 = "" : TmpTxt20 = ""
        TmpTxt21 = "" : TmpTxt22 = "" : TmpTxt23 = "" : TmpTxt24 = ""

        CntTests = 0

        ' Debug.Print(Pads.Text1.Text)

        If Disc.Text1.Text <> "" Then
            If Disc.Text1.Text <> "" Then TmpTxt01 = Disc.Text1.Text
            If Disc.Text2.Text <> "" Then TmpTxt02 = Disc.Text2.Text
            If Disc.Text3.Text <> "" Then TmpTxt03 = Disc.Text3.Text
            If Disc.Text4.Text <> "" Then TmpTxt04 = Disc.Text4.Text
            If Disc.Text5.Text <> "" Then TmpTxt05 = Disc.Text5.Text
            If Disc.Text6.Text <> "" Then TmpTxt06 = Disc.Text6.Text
            If Disc.Text7.Text <> "" Then TmpTxt07 = Disc.Text7.Text
            If Disc.Text8.Text <> "" Then TmpTxt08 = Disc.Text8.Text
            If Disc.Text9.Text <> "" Then TmpTxt09 = Disc.Text9.Text
            If Disc.Text10.Text <> "" Then TmpTxt10 = Disc.Text10.Text
            If Disc.Text11.Text <> "" Then TmpTxt11 = Disc.Text11.Text
            If Disc.Text12.Text <> "" Then TmpTxt12 = Disc.Text12.Text
        ElseIf Pads.Text1.Text <> "" Then
            If Pads.Text1.Text <> "" Then TmpTxt01 = Pads.Text1.Text
            If Pads.Text2.Text <> "" Then TmpTxt02 = Pads.Text2.Text
            If Pads.Text3.Text <> "" Then TmpTxt03 = Pads.Text3.Text
            If Pads.Text4.Text <> "" Then TmpTxt04 = Pads.Text4.Text
            If Pads.Text5.Text <> "" Then TmpTxt05 = Pads.Text5.Text
            If Pads.Text6.Text <> "" Then TmpTxt06 = Pads.Text6.Text
            If Pads.Text7.Text <> "" Then TmpTxt07 = Pads.Text7.Text
            If Pads.Text8.Text <> "" Then TmpTxt08 = Pads.Text8.Text
            If Pads.Text9.Text <> "" Then TmpTxt09 = Pads.Text9.Text
            If Pads.Text10.Text <> "" Then TmpTxt10 = Pads.Text10.Text
            If Pads.Text11.Text <> "" Then TmpTxt11 = Pads.Text11.Text
            If Pads.Text12.Text <> "" Then TmpTxt12 = Pads.Text12.Text
        End If

        If Disc.Text1.Text <> "" Or Pads.Text1.Text <> "" Then
            If lngNumCols = 4 Then

                'Gewicht
                If Disc.Text13.Text <> "" Then TmpTxt13 = Disc.Text13.Text
                If Pads.Text13.Text <> "" Then TmpTxt13 = Pads.Text13.Text

                'Name
                If Disc.Text19.Text <> "" Then TmpTxt14 = strUsername
                If Pads.Text19.Text <> "" Then TmpTxt14 = strUsername

                If Disc.TextBox3.Text <> "" Then TmpTxt15 = Disc.TextBox3.Text
                If Pads.TextBox3.Text <> "" Then TmpTxt15 = Pads.TextBox3.Text

                If Disc.TextBox2.Text <> "" Then TmpTxt16 = Disc.TextBox2.Text
                If Pads.TextBox2.Text <> "" Then TmpTxt16 = Pads.TextBox2.Text

                If Pads.Option1.Checked = True Then TmpTxt17 = "---->"
                If Pads.Option2.Checked = True Then TmpTxt17 = "<----"
                If Pads.Text4.Text <> "" Then TmpTxt18 = saveNow

                ' Runout and Disc ID
                If Disc.TextBox4.Text <> "" Then TmpTxt19 = Disc.TextBox4.Text

                If Disc.TextBox5.Text <> "" Then TmpTxt20 = Disc.TextBox5.Text
                If Disc.Text4.Text <> "" Then TmpTxt18 = saveNow

            ElseIf lngNumCols = 6 Then
                If Disc.Text13.Text <> "" Then TmpTxt13 = Disc.Text13.Text
                If Pads.Text13.Text <> "" Then TmpTxt13 = Pads.Text13.Text

                If Disc.Text19.Text <> "" Then TmpTxt14 = strUsername
                If Pads.Text19.Text <> "" Then TmpTxt14 = strUsername

                If Disc.TextBox3.Text <> "" Then TmpTxt15 = Disc.TextBox3.Text
                If Pads.TextBox3.Text <> "" Then TmpTxt15 = Pads.TextBox3.Text

                If Disc.TextBox2.Text <> "" Then TmpTxt16 = Disc.TextBox2.Text
                If Pads.TextBox2.Text <> "" Then TmpTxt16 = Pads.TextBox2.Text

                If Pads.Option1.Checked = True Then TmpTxt17 = "---->"
                If Pads.Option2.Checked = True Then TmpTxt17 = "<----"
                If Pads.Text6.Text <> "" Then TmpTxt18 = saveNow

                If Disc.TextBox4.Text <> "" Then TmpTxt19 = Disc.TextBox4.Text

                If Disc.TextBox5.Text <> "" Then TmpTxt20 = Disc.TextBox5.Text
                If Disc.Text6.Text <> "" Then TmpTxt18 = saveNow

            ElseIf lngNumCols = 8 Then
                If Disc.Text13.Text <> "" Then TmpTxt13 = Disc.Text13.Text
                If Pads.Text13.Text <> "" Then TmpTxt13 = Pads.Text13.Text

                If Disc.Text19.Text <> "" Then TmpTxt14 = strUsername
                If Pads.Text19.Text <> "" Then TmpTxt14 = strUsername

                If Disc.TextBox3.Text <> "" Then TmpTxt15 = Disc.TextBox3.Text
                If Pads.TextBox3.Text <> "" Then TmpTxt15 = Pads.TextBox3.Text

                If Disc.TextBox2.Text <> "" Then TmpTxt16 = Disc.TextBox2.Text
                If Pads.TextBox2.Text <> "" Then TmpTxt16 = Pads.TextBox2.Text

                If Pads.Option1.Checked = True Then TmpTxt17 = "---->"
                If Pads.Option2.Checked = True Then TmpTxt17 = "<----"
                If Pads.Text8.Text <> "" Then TmpTxt18 = saveNow

                If Disc.TextBox4.Text <> "" Then TmpTxt19 = Disc.TextBox4.Text

                If Disc.TextBox5.Text <> "" Then TmpTxt20 = Disc.TextBox5.Text
                If Disc.Text8.Text <> "" Then TmpTxt18 = saveNow

            ElseIf lngNumCols = 9 Then
                If Disc.Text13.Text <> "" Then TmpTxt13 = Disc.Text13.Text
                If Pads.Text13.Text <> "" Then TmpTxt13 = Pads.Text13.Text

                If Disc.Text19.Text <> "" Then TmpTxt14 = strUsername
                If Pads.Text19.Text <> "" Then TmpTxt14 = strUsername

                If Disc.TextBox3.Text <> "" Then TmpTxt15 = Disc.TextBox3.Text
                If Pads.TextBox3.Text <> "" Then TmpTxt15 = Pads.TextBox3.Text

                If Disc.TextBox2.Text <> "" Then TmpTxt16 = Disc.TextBox2.Text
                If Pads.TextBox2.Text <> "" Then TmpTxt16 = Pads.TextBox2.Text

                If Pads.Option1.Checked = True Then TmpTxt17 = "---->"
                If Pads.Option2.Checked = True Then TmpTxt17 = "<----"
                If Pads.Text9.Text <> "" Then TmpTxt18 = saveNow

                If Disc.TextBox4.Text <> "" Then TmpTxt19 = Disc.TextBox4.Text
                If Disc.TextBox5.Text <> "" Then TmpTxt20 = Disc.TextBox5.Text
                If Disc.Text9.Text <> "" Then TmpTxt18 = saveNow

            ElseIf lngNumCols = 12 Then
                If Disc.Text13.Text <> "" Then TmpTxt13 = Disc.Text13.Text
                If Pads.Text13.Text <> "" Then TmpTxt13 = Pads.Text13.Text

                If Disc.Text19.Text <> "" Then TmpTxt14 = strUsername
                If Pads.Text19.Text <> "" Then TmpTxt14 = strUsername

                If Disc.TextBox3.Text <> "" Then TmpTxt15 = Disc.TextBox3.Text
                If Pads.TextBox3.Text <> "" Then TmpTxt15 = Pads.TextBox3.Text

                If Disc.TextBox2.Text <> "" Then TmpTxt16 = Disc.TextBox2.Text
                If Pads.TextBox2.Text <> "" Then TmpTxt16 = Pads.TextBox2.Text

                If Pads.Option1.Checked = True Then TmpTxt17 = "---->"
                If Pads.Option2.Checked = True Then TmpTxt17 = "<----"
                If Pads.Text12.Text <> "" Then TmpTxt18 = saveNow

                If Disc.TextBox4.Text <> "" Then TmpTxt19 = Disc.TextBox4.Text
                If Disc.TextBox5.Text <> "" Then TmpTxt20 = Disc.TextBox5.Text
                If Disc.Text12.Text <> "" Then TmpTxt18 = saveNow

            End If
        End If

        'Public table As DataTable = DS3.Tables.Item(0)  'New DataTable("table")
        Dim table As DataTable = frmLoginDB.MessungenDataSet.Template 'DS3.Tables.Item(0)
        Dim relation As DataRow
        Dim rowArray(26) As Object

        If Pads.TextBox1.Text <> "" Then
            TmpTxt21 = Pads.TextBox1.Text
        ElseIf Disc.TextBox1.Text <> "" Then
            TmpTxt21 = Disc.TextBox1.Text
        End If

        If Padoutside = 1 Then
            TmpTxt22 = "Outside"
        ElseIf PadInside = 1 Then
            TmpTxt22 = "Inside"
        Else
            TmpTxt22 = "Disc"
        End If

        For i = 0 To frmLoginDB.MessungenDataSet.Template.Rows.Count - 1 'DS3.Tables.Item(0).Rows.Count - 1
            CntTests = i
        Next

        If TmpTxt01 = "" Then TmpTxt01 = "0,0"
        If TmpTxt02 = "" Then TmpTxt02 = "0,0"
        If TmpTxt03 = "" Then TmpTxt03 = "0,0"
        If TmpTxt04 = "" Then TmpTxt04 = "0,0"
        If TmpTxt05 = "" Then TmpTxt05 = "0,0"
        If TmpTxt06 = "" Then TmpTxt06 = "0,0"
        If TmpTxt07 = "" Then TmpTxt07 = "0,0"
        If TmpTxt08 = "" Then TmpTxt08 = "0,0"
        If TmpTxt09 = "" Then TmpTxt09 = "0,0"
        If TmpTxt10 = "" Then TmpTxt10 = "0,0"
        If TmpTxt11 = "" Then TmpTxt11 = "0,0"
        If TmpTxt12 = "" Then TmpTxt12 = "0,0"
        If TmpTxt16 = "" Then TmpTxt16 = "0"
        If TmpTxt15 = "" Then TmpTxt15 = "0"


        'rowArray(0) = ""
        rowArray(0) = DBTemplateName
        rowArray(1) = DBTestnr
        rowArray(2) = TmpTxt21
        rowArray(3) = PadDisc
        rowArray(4) = CDbl(TmpTxt01)
        rowArray(5) = CDbl(TmpTxt02)
        rowArray(6) = CDbl(TmpTxt03)
        rowArray(7) = CDbl(TmpTxt04)
        rowArray(8) = CDbl(TmpTxt05)
        rowArray(9) = CDbl(TmpTxt06)
        rowArray(10) = CDbl(TmpTxt07)
        rowArray(11) = CDbl(TmpTxt08)
        rowArray(12) = CDbl(TmpTxt09)
        rowArray(13) = CDbl(TmpTxt10)
        rowArray(14) = CDbl(TmpTxt11)
        rowArray(15) = CDbl(TmpTxt12)
        rowArray(16) = TmpTxt19
        rowArray(17) = TmpTxt20
        rowArray(18) = ""
        rowArray(19) = TmpTxt22
        rowArray(20) = TmpTxt17
        rowArray(21) = CDbl(TmpTxt13)
        rowArray(22) = CInt(TmpTxt16)
        rowArray(23) = CInt(TmpTxt15)
        rowArray(24) = TmpTxt18
        rowArray(25) = "Nein" ' -1
        rowArray(26) = strUsername

        relation = frmLoginDB.MessungenDataSet.Template.NewRow()
        'relation = table.NewRow()
        relation.ItemArray = rowArray
        'table.Rows.Add(relation)

        frmLoginDB.MessungenDataSet.Template.Rows.Add(relation)

        PrintTable(frmLoginDB.MessungenDataSet.Template)

        AddNewRow()

    End Sub

    Private Sub PrintTable(ByVal table As DataTable)
        Dim row As DataRow
        Dim column As DataColumn
        For Each row In table.Rows
            For Each column In table.Columns
                Console.WriteLine(row(column))
            Next
        Next
    End Sub

    Private Sub AddNewRow()

        '        Dim newOrders As NorthwindDataSet.OrdersDataTable
        Dim newData As MeasureAndWeigh.MessungenDataSet.TemplateDataTable
        'newOrders = CType(NorthwindDataSet.Orders.GetChanges(Data.DataRowState.Added), _
        '    NorthwindDataSet.OrdersDataTable)

        newData = CType(frmLoginDB.MessungenDataSet.Template.GetChanges(Data.DataRowState.Added), _
                 MessungenDataSet.TemplateDataTable)

        If Not IsNothing(newData) Then
            Try
                frmLoginDB.TemplateTableAdapter.Update(newData)
                'OrdersTableAdapter.Update(newOrders)
            Catch ex As Exception
                MessageBox.Show("AddNew Failed" & vbCrLf & ex.Message)
            End Try
        End If
    End Sub

 
End Module
