Module ReadandWirteExcel

    Public xlAppl As Object
    Public xlWB As Object
    Public xlWS As Object

    Public xlApplLiefNicht As Boolean
    Public boolWBOffen As Boolean
    Public wb As Object

    Public lngNumCols, lngNumRows, lngTotalNumCols As Integer

    Public Sub readExcel()
        Dim i As Short

        On Error Resume Next

        If DBOrExcel = 2 Then

            If Err.Number <> 0 Then xlApplLiefNicht = True

            Err.Clear()

            If xlAppl Is Nothing Then
                xlAppl = CType(CreateObject("Excel.Application"), _
                           Microsoft.Office.Interop.Excel.Application)
                If Err.Number <> 0 Then
                    MsgBox("Konnte keine Verbindung zu Excel herstellen !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    GoTo err_Handler
                End If
            End If

            If Not xlApplLiefNicht Then
                If xlAppl.Workbooks.Count > 0 Then
                    For Each wb In xlAppl.Workbooks
                        If LCase(wb.Name) = LCase(xlDateiName) Then
                            'wb.Activate
                            boolWBOffen = True
                            Exit For
                        End If
                    Next wb
                End If
            End If

            'LVCVTimer: 1 = CV ; 2 = LV
            If frmLogin.Option1.Visible = True Then
                If LVCVTimer = 1 Then
                    PathNam = CVPath
                ElseIf LVCVTimer = 2 Then
                    PathNam = LVPath
                End If
            Else
                PathNam = CVPath
            End If

            On Error Resume Next
            If Not boolWBOffen Then
                Err.Clear()

                xlWB = xlAppl.Workbooks.Open(FileName:=PathNam & xlDateiName)

                If Err.Number <> 0 Then
                    MsgBox("Die Arbeitsmappe '" & xlDateiName & "' konnte nicht geöffnet werden !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    If xlApplLiefNicht Then xlAppl.Application.Quit()
                    xlAppl = Nothing
                    GoTo err_Handler
                Else
                    boolWBOffen = True
                End If

            Else
                xlWB = xlAppl.Workbooks(xlDateiName)
                boolWBOffen = True
            End If

            On Error GoTo 0
            xlwscnt = xlWB.Worksheets.Count
            For i = 1 To xlWB.Worksheets.Count
                If Sheets01 <> "" Then
                    If xlWB.Worksheets(i).Name = Sheets01 Then
                        xlWS = xlWB.Worksheets(Sheets01)
                        xlWB.Worksheets(Sheets01).Select()
                        lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                        lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                        Exit For
                    End If
                ElseIf Sheets02 <> "" Then
                    If xlWB.Worksheets(i).Name = Sheets02 Then
                        xlWS = xlWB.Worksheets(Sheets02)
                        xlWB.Worksheets(Sheets02).Select()
                        lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                        lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                        Exit For
                    End If
                ElseIf Sheets03 <> "" Then
                    If xlWB.Worksheets(i).Name = Sheets03 Then
                        xlWS = xlWB.Worksheets(Sheets03)
                        xlWB.Worksheets(Sheets03).Select()
                        lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                        lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                        Exit For
                    End If
                End If
            Next

            If lngNumRows >= 2 Then
                xlSpaltenEinlesen()

                If State = "DE" Then
                    Form2.Label1.Text = "Messpunkte : "
                    Form2.Command2.Text = "Schliessen"
                Else
                    Form2.Label1.Text = "Measure points : "
                    Form2.Command2.Text = "Close"
                End If

                lngTotalNumCols = lngNumCols

                If Disccnt = 1 Then
                    lngNumCols = lngNumCols - 8
                Else
                    lngNumCols = lngNumCols - 7
                End If

                Form2.Label2.Text = CStr(lngNumCols)

            End If

            'Excel sichtbar machen:
            If excelview = "True" Then
                xlAppl.Application.Visible = True
            Else
                xlAppl.Application.Visible = False
            End If
        End If
        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        'ausser Beenden-Button deaktivieren:
        Dim ctl As System.Windows.Forms.Control

        For Each ctl In Form2.Controls
            If ctl.Name <> "cmdBeenden" Then
                ctl.Enabled = False
            End If
        Next ctl

        Exit Sub
    End Sub

    Private Sub xlSpaltenEinlesen()

        If DBOrExcel = 2 Then

            Dim lngNumRows As Integer
            Dim lngRowIndex As Integer
            Dim intColIndex As Short
            Dim strTemp As String

            Form2.cmbAuswahl.Items.Clear()
            Form2.cmbAuswahl.Items.Add("")

            lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            intColIndex = 1
            For lngRowIndex = 2 To lngNumRows
                If CStr(xlWS.Cells(lngRowIndex, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 1, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 2, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 3, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 4, intColIndex + 1).Value) = "" Then
                    strTemp = xlWS.Cells(lngRowIndex, intColIndex).Value
                    If strTemp <> "" Then
                        Form2.cmbAuswahl.Items.Add(strTemp)
                    Else
                        Exit For
                    End If
                End If
            Next lngRowIndex
        End If
    End Sub

    Sub ValuesWritetoexcel()
        Dim saveNow As DateTime = DateTime.Now
        Dim StartTestNr0 As String
        Dim StartTestNr1 As String

        StartTestNr0 = Microsoft.VisualBasic.Left(xlDateiName, 1)
        StartTestNr1 = Microsoft.VisualBasic.Left(DBTestnr, 1)

        If DBOrExcel = 2 Then

            If Microsoft.VisualBasic.Left(xlDateiName, 1) = StartCVTestNr _
                Or Microsoft.VisualBasic.Left(DBTestnr, 1) = StartCVTestNr Then

                If Disc.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Disc.Text1.Text)
                If Disc.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Disc.Text2.Text)
                If Disc.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Disc.Text3.Text)
                If Disc.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Disc.Text4.Text)
                If Disc.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Disc.Text5.Text)
                If Disc.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Disc.Text6.Text)
                If Disc.Text7.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Disc.Text7.Text)
                If Disc.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = CDbl(Disc.Text8.Text)
                If Disc.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Disc.Text9.Text)
                If Disc.Text10.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = CDbl(Disc.Text10.Text)
                If Disc.Text11.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = CDbl(Disc.Text11.Text)
                If Disc.Text12.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = CDbl(Disc.Text12.Text)

            ElseIf Microsoft.VisualBasic.Left(xlDateiName, 1) = StartLVTestNr _
                   Or Microsoft.VisualBasic.Left(DBTestnr, 1) = StartLVTestNr Then

                If xlWS.Cells(40, 14).Value = "" Then xlWS.Cells(40, 14).Value = Disc.ComboBox3.Text '  Disc.ComboBox3.Text

                If lngNumCols = 4 Then
                    If Disc.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Disc.Text1.Text)
                    If Disc.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Disc.Text2.Text)
                    If Disc.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Disc.Text3.Text)
                    If Disc.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Disc.Text4.Text)
                Else
                    If Disc.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Disc.Text1.Text)
                    If Disc.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Disc.Text2.Text)
                    If Disc.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Disc.Text3.Text)
                    If Disc.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Disc.Text4.Text)
                    If Disc.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Disc.Text5.Text)
                    If Disc.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Disc.Text6.Text)
                    If Disc.Text7.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Disc.Text7.Text)
                    If Disc.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = CDbl(Disc.Text8.Text)
                    If Disc.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Disc.Text9.Text)
                    If Disc.Text10.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = CDbl(Disc.Text10.Text)
                    If Disc.Text11.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = CDbl(Disc.Text11.Text)
                    If Disc.Text12.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = CDbl(Disc.Text12.Text)
                End If
            End If

            If Pads.Text1.Text <> "" Then

                If lngNumCols = 4 Then
                    If Pads.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Pads.Text1.Text)
                    If Pads.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Pads.Text2.Text)
                    If Pads.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Pads.Text3.Text)
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Pads.Text4.Text)
                ElseIf lngNumCols = 6 Then
                    If Pads.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Pads.Text1.Text)
                    If Pads.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Pads.Text2.Text)
                    If Pads.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Pads.Text3.Text)
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Pads.Text4.Text)
                    If Pads.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Pads.Text5.Text)
                    If Pads.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Pads.Text6.Text)
                ElseIf lngNumCols = 8 Then
                    If Pads.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Pads.Text1.Text)
                    If Pads.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Pads.Text2.Text)
                    If Pads.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Pads.Text3.Text)
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Pads.Text4.Text)
                    If Pads.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Pads.Text5.Text)
                    If Pads.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Pads.Text6.Text)
                    If Pads.Text7.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Pads.Text7.Text)
                    If Pads.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = CDbl(Pads.Text8.Text)
                ElseIf lngNumCols = 9 Then
                    If Pads.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Pads.Text1.Text)
                    If Pads.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Pads.Text2.Text)
                    If Pads.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Pads.Text3.Text)
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Pads.Text4.Text)
                    If Pads.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Pads.Text5.Text)
                    If Pads.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Pads.Text6.Text)
                    If Pads.Text7.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Pads.Text7.Text)
                    If Pads.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = CDbl(Pads.Text8.Text)
                    If Pads.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Pads.Text9.Text)
                ElseIf lngNumCols = 12 Then
                    If Pads.Text1.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum).Value = CDbl(Pads.Text1.Text)
                    If Pads.Text2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 1).Value = CDbl(Pads.Text2.Text)
                    If Pads.Text3.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 2).Value = CDbl(Pads.Text3.Text)
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 3).Value = CDbl(Pads.Text4.Text)
                    If Pads.Text5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Pads.Text5.Text)
                    If Pads.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = CDbl(Pads.Text6.Text)
                    If Pads.Text7.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Pads.Text7.Text)
                    If Pads.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = CDbl(Pads.Text8.Text)
                    If Pads.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Pads.Text9.Text)
                    If Pads.Text10.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = CDbl(Pads.Text10.Text)
                    If Pads.Text11.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = CDbl(Pads.Text11.Text)
                    If Pads.Text12.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = CDbl(Pads.Text12.Text)
                End If

            End If

            If Disc.Text1.Text <> "" Then

                If lngNumCols = 4 Then
                    If Disc.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Disc.Text13.Text)
                    If Disc.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = strUsername
                    If Disc.TextBox6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = Disc.TextBox6.Text
                    If Disc.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = Disc.TextBox2.Text
                    If Disc.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = Disc.TextBox4.Text
                    If Disc.TextBox5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = Disc.TextBox5.Text
                    If Disc.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = saveNow
                ElseIf lngNumCols = 6 Then
                    If Disc.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Disc.Text13.Text)
                    If Disc.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = strUsername
                    If Disc.TextBox6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = Disc.TextBox6.Text
                    If Disc.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = Disc.TextBox2.Text
                    If Disc.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = Disc.TextBox4.Text
                    If Disc.TextBox5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = Disc.TextBox5.Text
                    If Disc.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = saveNow
                ElseIf lngNumCols = 8 Then
                    If Disc.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Disc.Text13.Text)
                    If Disc.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = strUsername
                    If Disc.TextBox6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = Disc.TextBox6.Text
                    If Disc.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = Disc.TextBox2.Text
                    If Disc.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = Disc.TextBox4.Text
                    If Disc.TextBox5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = Disc.TextBox5.Text
                    If Disc.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 14).Value = saveNow
                ElseIf lngNumCols = 9 Then
                    If Disc.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = CDbl(Disc.Text13.Text)
                    If Disc.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = strUsername
                    If Disc.TextBox6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = Disc.TextBox6.Text
                    If Disc.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = Disc.TextBox2.Text
                    If Disc.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = Disc.TextBox4.Text
                    If Disc.TextBox5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 14).Value = Disc.TextBox5.Text
                    If Disc.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 15).Value = saveNow
                ElseIf lngNumCols = 12 Then
                    If Disc.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = CDbl(Disc.Text13.Text)
                    If Disc.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = strUsername
                    If Disc.TextBox6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 14).Value = Disc.TextBox6.Text
                    If Disc.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 15).Value = Disc.TextBox2.Text
                    If Disc.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 16).Value = Disc.TextBox4.Text
                    If Disc.TextBox5.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 17).Value = Disc.TextBox5.Text
                    If Disc.Text12.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 18).Value = saveNow
                End If

            ElseIf Pads.Text1.Text <> "" Then

                If lngNumCols = 4 Then
                    If Pads.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 4).Value = CDbl(Pads.Text13.Text)
                    If Pads.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 5).Value = strUsername
                    If Pads.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = Pads.TextBox4.Text
                    If Pads.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = Pads.TextBox2.Text
                    If Pads.Option1.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = "---->"
                    If Pads.Option2.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = "<----"
                    If Pads.Text4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = saveNow
                ElseIf lngNumCols = 6 Then
                    If Pads.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 6).Value = CDbl(Pads.Text13.Text)
                    If Pads.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 7).Value = strUsername
                    If Pads.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = Pads.TextBox4.Text
                    If Pads.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = Pads.TextBox2.Text
                    If Pads.Option1.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = "---->"
                    If Pads.Option2.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = "<----"
                    If Pads.Text6.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = saveNow
                ElseIf lngNumCols = 8 Then
                    If Pads.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 8).Value = CDbl(Pads.Text13.Text)
                    If Pads.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = strUsername
                    If Pads.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = Pads.TextBox4.Text
                    If Pads.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = Pads.TextBox2.Text
                    If Pads.Option1.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = "---->"
                    If Pads.Option2.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = "<----"
                    If Pads.Text8.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = saveNow
                ElseIf lngNumCols = 9 Then
                    If Pads.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 9).Value = CDbl(Pads.Text13.Text)
                    If Pads.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 10).Value = strUsername
                    If Pads.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 11).Value = Pads.TextBox4.Text
                    If Pads.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = Pads.TextBox2.Text
                    If Pads.Option1.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = "---->"
                    If Pads.Option2.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = "<----"
                    If Pads.Text9.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 14).Value = saveNow
                ElseIf lngNumCols = 12 Then
                    If Pads.Text13.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 12).Value = CDbl(Pads.Text13.Text)
                    If Pads.Text19.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 13).Value = strUsername
                    If Pads.TextBox4.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 14).Value = Pads.TextBox4.Text
                    If Pads.TextBox2.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 15).Value = Pads.TextBox2.Text
                    If Pads.Option1.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 16).Value = "---->"
                    If Pads.Option2.Checked = True Then xlWS.Cells(startWriteRow, startWriteColum + 16).Value = "<----"
                    If Pads.Text12.Text <> "" Then xlWS.Cells(startWriteRow, startWriteColum + 17).Value = saveNow

                End If
            End If
        End If

        ClearTexts()

    End Sub


    Sub ClearTexts()

        'Clear Pads Formular
        Pads.Text1.Text = ""
        Pads.Text2.Text = ""
        Pads.Text3.Text = ""
        Pads.Text4.Text = ""
        Pads.Text5.Text = ""
        Pads.Text6.Text = ""
        Pads.Text7.Text = ""
        Pads.Text8.Text = ""
        Pads.Text9.Text = ""
        Pads.Text10.Text = ""
        Pads.Text11.Text = ""
        Pads.Text12.Text = ""
        Pads.Text13.Text = ""
        Pads.Text19.Text = ""
        Pads.TextBox4.Text = ""
        Pads.TextBox2.Text = ""

        ' Clear CV Disc Fomular
        Disc.Text1.Text = ""
        Disc.Text2.Text = ""
        Disc.Text3.Text = ""
        Disc.Text4.Text = ""
        Disc.Text5.Text = ""
        Disc.Text6.Text = ""
        Disc.Text7.Text = ""
        Disc.Text8.Text = ""
        Disc.Text9.Text = ""
        Disc.Text10.Text = ""
        Disc.Text11.Text = ""
        Disc.Text12.Text = ""
        Disc.Text13.Text = ""
        Disc.Text19.Text = ""
        Disc.TextBox6.Text = ""
        Disc.TextBox2.Text = ""
        Disc.TextBox4.Text = ""
        Disc.TextBox5.Text = ""

    End Sub

    Sub CheckValuesIO()
        'Check Pads
        If Pads.Text1.Visible = True And Pads.Text1.Text = "" Or _
           Pads.Text2.Visible = True And Pads.Text2.Text = "" Or _
           Pads.Text3.Visible = True And Pads.Text3.Text = "" Or _
           Pads.Text4.Visible = True And Pads.Text4.Text = "" Or _
           Pads.Text5.Visible = True And Pads.Text5.Text = "" Or _
           Pads.Text6.Visible = True And Pads.Text6.Text = "" Or _
           Pads.Text7.Visible = True And Pads.Text7.Text = "" Or _
           Pads.Text9.Visible = True And Pads.Text9.Text = "" Or _
           Pads.Text10.Visible = True And Pads.Text10.Text = "" Or _
           Pads.Text11.Visible = True And Pads.Text11.Text = "" Or _
           Pads.Text12.Visible = True And Pads.Text12.Text = "" Or _
           Pads.Text13.Visible = True And Pads.Text13.Text = "" Then
            valfailed = 1
        End If

        If Disc.Text1.Visible = True And Disc.Text1.Text = "" Or _
           Disc.Text2.Visible = True And Disc.Text2.Text = "" Or _
           Disc.Text3.Visible = True And Disc.Text3.Text = "" Or _
           Disc.Text4.Visible = True And Disc.Text4.Text = "" Or _
           Disc.Text5.Visible = True And Disc.Text5.Text = "" Or _
           Disc.Text6.Visible = True And Disc.Text6.Text = "" Or _
           Disc.Text7.Visible = True And Disc.Text7.Text = "" Or _
           Disc.Text9.Visible = True And Disc.Text9.Text = "" Or _
           Disc.Text10.Visible = True And Disc.Text10.Text = "" Or _
           Disc.Text11.Visible = True And Disc.Text11.Text = "" Or _
           Disc.Text12.Visible = True And Disc.Text12.Text = "" Or _
           Disc.Text13.Visible = True And Disc.Text13.Text = "" Then
            valfailed = 1
        End If

        If valfailed = 1 Then
            If State = "DE" Then
                MsgBox("Ein Wert fehlt, bitte überprüfen Sie die gemessenen Werte !", MsgBoxStyle.Critical, "Wert fehlt !")
            Else
                MsgBox("Measurement not complete, please check the values !", MsgBoxStyle.Critical, "Value failed !")
            End If
        End If
    End Sub

End Module
