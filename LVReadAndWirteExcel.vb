Module LVReadAndWirteExcel

    Sub LVMircofaceExcel()
        '*Dim boolWBOffen As Boolean
        Dim i As Short
        '* Dim wb As Object ' As Excel.Workbook
        '*Dim wb As Microsoft.Office.Interop.Excel.Workbook

        'If cntBack = 1 Then Command2_Click() 'cmdBeenden_Click()

        ' GetObject-Funktionsaufruf ohne erstes Argument gibt einen
        ' Verweis auf eine Instanz der Anwendung zurück. Wenn die
        ' Anwendung nicht ausgeführt wird, tritt ein Fehler auf.
        If DBOrExcel = 2 Then
            'Prüfen, ob Excel ausgeführt wird:
            On Error Resume Next
            '* xlAppl = GetObject(, "Excel.Application")

            '* xlAppl = CType(CreateObject("Excel.Application"), _
            '*           Microsoft.Office.Interop.Excel.Application)


            '*   xlApp = CType(CreateObject("Excel.Application"), _
            '*                     Microsoft.Office.Interop.Excel.Application)

            If Err.Number <> 0 Then xlApplLiefNicht = True

            Err.Clear() ' Err-Objekt im Fehlerfall löschen

            'Wenn Excel nicht ausgeführt wird, Excel starten:
            If xlAppl Is Nothing Then
                '*            xlAppl = CreateObject("Excel.Application")
                xlAppl = CType(CreateObject("Excel.Application"), _
                           Microsoft.Office.Interop.Excel.Application)


                '*    xlApp = CType(CreateObject("Excel.Application"), _
                '*                       Microsoft.Office.Interop.Excel.Application)

                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Konnte keine Verbindung zu Excel herstellen !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    GoTo err_Handler
                End If
            End If

            'Prüfen, ob Arbeitsmappe bereits offen ist:
            '*boolWBOffen = False

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

            'LVPath As String, CVPath As String, PathNam As String
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
                'Wenn Arbeitsmappe nicht offen, öffnen und Verweis setzen:
                Err.Clear()
                xlWB = xlAppl.Workbooks.Open(FileName:=PathNam & xlDateiName)
                '*   xlBook = xlApp.Workbooks.Open(Filename:=PathNam & xlDateiName)
                '#App.Path & "\" & xlDateiName)

                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Die Arbeitsmappe '" & xlDateiName & "' konnte nicht geöffnet werden !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    If xlApplLiefNicht Then xlAppl.Application.Quit()
                    xlAppl = Nothing
                    GoTo err_Handler
                End If

            Else
                'Verweis setzen:
                xlWB = xlAppl.Workbooks(xlDateiName)
                '*     xlBook = xlApp.Workbooks(xlDateiName)
            End If

            xlwscnt = xlWB.Worksheets.Count

            On Error GoTo 0
            If MFCnt <> 1 Then
                'Verweis auf Tabellenblatt setzen:
                For i = 1 To xlWB.Worksheets.Count
                    If Sheets01 <> "" Then
                        If xlWB.Worksheets(i).Name = Sheets01 Then
                            xlWS = xlWB.Worksheets(Sheets01)
                            xlWB.Worksheets(Sheets01).Select()
                            lngNumCols = xlWS.Range("Z2").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                            lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                            Exit For
                        End If
                    ElseIf Sheets02 <> "" Then
                        If xlWB.Worksheets(i).Name = Sheets02 Then
                            xlWS = xlWB.Worksheets(Sheets02)
                            xlWB.Worksheets(Sheets02).Select()
                            lngNumCols = xlWS.Range("Z2").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                            lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                            Exit For
                        End If
                    ElseIf Sheets03 <> "" Then
                        If xlWB.Worksheets(i).Name = Sheets03 Then
                            xlWS = xlWB.Worksheets(Sheets03)
                            xlWB.Worksheets(Sheets03).Select()
                            'xlToRight
                            lngNumCols = xlWS.Range("Z2").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                            lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                            Exit For
                        End If
                    End If
                Next
            ElseIf MFCnt = 1 Then
                If Sheets01 <> "" Then
                    Sheets01 = Sheets01
                End If
                If Sheets02 <> "" Then
                    Sheets01 = Sheets02
                End If
                If Sheets03 <> "" Then
                    Sheets01 = Sheets03
                End If
                Dim xlDateiName00 As String
                xlDateiName00 = ""
                '#For i = 1 To Len(xlDateiName)
                '#xlDateiName00 = Mid(xlDateiName, i, 1)
                '#If xlDateiName00 = "." Then
                '#xlDateiName00 = Mid(xlDateiName, 1, i - 1)
                '#Exit For
                '#End If
                '#Next
                'xlWS = xlWB.Worksheets(xlDateiName00)
                'xlWB.Worksheets(xlDateiName00).Select(1)
                'lngNumCols = xlWS.Range("Z2").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                'lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                If tttemp <> "" Then
                    xlDateiName00 = xlWB.Worksheets(1).Name
                    xlWS = xlWB.Worksheets(xlDateiName00)
                    'xlWB.Worksheets(xlDateiName00).Select()
                    'xlWB.Worksheets(i).Select()
                    lngNumCols = xlWS.Range("Z2").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                    lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                ElseIf tttemp = "" Then
                    For i = 1 To xlWB.Worksheets.Count
                        If xlWB.Worksheets(i).Name = Sheets01 Then
                            xlDateiName00 = xlWB.Worksheets(i).Name
                            xlWS = xlWB.Worksheets(xlDateiName00)
                            'xlWB.Worksheets(xlDateiName00).Select()
                            xlWB.Worksheets(i).Select()
                            lngNumCols = xlWS.Range("A1").End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight).Column
                            lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                        End If
                    Next
                End If
            End If
            '#Set xlWS = xlWB.Worksheets(xlWS_Name)

            'Anzahl Zeilen in Excel-Tabelle ermitteln:
            '#lngNumRows = xlWS.Range("A65536").End(xlUp).Row

            'Wenn Anzahl der Zeilen > 2 (Überschrift + min. einen  Datensatz)...
            If lngNumRows >= 2 Then
                'Excel-Tabelle Spalte 1 in cmbAuswahl einlesen:
                MFSpalteneinlesen()
                'Ersten Eintrag in Combo auswählen:
                ' cmbAuswahl.ListIndex = 0
                If State = "DE" Then
                    Form2.Label1.Text = "Messpunkte : "
                    Form2.Command2.Text = "Schliessen"
                Else
                    Form2.Label1.Text = "Measure points : "
                    Form2.Command2.Text = "Close"
                End If
                '#####################################################################
                ' Messpunkt Berechnung
                If tttemp <> "" Then

                    lngTotalNumCols = lngNumCols

                    If Disccnt = 1 Then
                        'lngNumCols = lngNumCols - 6
                        lngNumCols = lngNumCols - 7
                    Else
                        'lngNumCols = lngNumCols - 4
                        lngNumCols = lngNumCols - 6
                    End If

                    'lngNumCols = lngNumCols - 4
                ElseIf tttemp = "" Then
                    lngTotalNumCols = lngNumCols

                    If Disccnt = 1 Then
                        'lngNumCols = lngNumCols - 7
                        lngNumCols = lngNumCols - 8
                    Else
                        'lngNumCols = lngNumCols - 5
                        lngNumCols = lngNumCols - 7
                    End If

                    ' lngNumCols = lngNumCols - 5
                End If

                Form2.Label2.Text = CStr(lngNumCols)

                '#####################################################################
            Else
                'ComboBox und verschiedene Buttons deaktiveren:
                'EnableControls(False)
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

    Sub MFSpalteneinlesen()
        Dim lngNumRows, lngRowIndex, i, intColIndex, a As Integer
        Dim strTemp As String
        'Inhalt von cmbAuswahl löschen:
        
        If DBOrExcel = 2 Then
            Form2.cmbAuswahl.Items.Clear()
            Form2.cmbAuswahl.Items.Add("")
            'Anzahl Zeilen in Excel-Tabelle ermitteln:
            If tttemp <> "" Then

                'Debug.Print(LVCVTimer)

                lngNumRows = xlWS.Range("Z65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                intColIndex = 26
                intColIndex0 = intColIndex

                For lngRowIndex = 2 To lngNumRows
                    strTemp = xlWS.Cells(lngRowIndex, intColIndex).Value
                    If strTemp = Sheets01 Then
                        txtCnt = lngRowIndex
                        a = 1
                        Do While CStr(xlWS.Cells(lngRowIndex, intColIndex + a).Value) <> ""
                            lngNumCols = intColIndex + a
                            a = a + 1
                        Loop

                        lngNumCols = lngNumCols - intColIndex

                        '#                Debug.Print(CStr(xlWS.Cells(lngRowIndex, intColIndex + a - 1).Value))

                        For i = 1 To MessCnt
                            strTemp = xlWS.Cells(lngRowIndex + i, intColIndex).Value
                            If CStr(xlWS.Cells(lngRowIndex + i, intColIndex + 1).Value) = "" And _
                            CStr(xlWS.Cells(lngRowIndex + i + 1, intColIndex + 1).Value) = "" Or _
                            CStr(xlWS.Cells(lngRowIndex + i + 1, intColIndex + 1).Value) = "1" Then
                                If CStr(xlWS.Cells(lngRowIndex + i, intColIndex + 1).Value) = "" Then
                                    Form2.cmbAuswahl.Items.Add(strTemp)
                                End If
                            End If
                        Next
                        Exit For
                    End If

                Next lngRowIndex

            ElseIf tttemp = "" Then
                lngNumRows = xlWS.Range("A65536").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                intColIndex = 1
                intColIndex0 = intColIndex
                'Excel-Tabelle/Spalte ab Zeile 2, in ComboBox cmbAuswahl einlesen:
                intColIndex = 1
                For lngRowIndex = 2 To lngNumRows
                    If CStr(xlWS.Cells(lngRowIndex, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 1, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 2, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 3, intColIndex + 1).Value) = "" And _
                CStr(xlWS.Cells(lngRowIndex + 4, intColIndex + 1).Value) = "" Then
                        strTemp = xlWS.Cells(lngRowIndex, intColIndex).Value
                        ' &
                        '", " & xlWS.Cells(lngRowIndex, intColIndex + 1).Value
                        If strTemp <> "" Then
                            Form2.cmbAuswahl.Items.Add(strTemp)
                        Else
                            Exit For
                        End If

                    End If
                Next lngRowIndex

            End If
        End If
    End Sub

    Sub SearchPPGNumInExcel()
        '*Dim boolWBOffen As Boolean
        Dim i As Short
        '* Dim wb As Object ' As Excel.Workbook

        'If cntBack = 1 Then Command2_Click() 'cmdBeenden_Click()

        ' GetObject-Funktionsaufruf ohne erstes Argument gibt einen
        ' Verweis auf eine Instanz der Anwendung zurück. Wenn die
        ' Anwendung nicht ausgeführt wird, tritt ein Fehler auf.
        If DBOrExcel = 2 Then


            'Prüfen, ob Excel ausgeführt wird:
            On Error Resume Next
            xlAppl = GetObject(, "Excel.Application")
            If Err.Number <> 0 Then xlApplLiefNicht = True

            Err.Clear() ' Err-Objekt im Fehlerfall löschen

            'Wenn Excel nicht ausgeführt wird, Excel starten:
            If xlAppl Is Nothing Then
                xlAppl = CreateObject("Excel.Application")

                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Konnte keine Verbindung zu Excel herstellen !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    GoTo err_Handler
                End If
            End If

            'Prüfen, ob Arbeitsmappe bereits offen ist:
            '*boolWBOffen = False

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

            'LVPath As String, CVPath As String, PathNam As String
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

            'On Error Resume Next
            If Not boolWBOffen Then
                'Wenn Arbeitsmappe nicht offen, öffnen und Verweis setzen:
                Err.Clear()
                xlWB = xlAppl.Workbooks.Open(FileName:=PathNam & xlDateiName)

                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Die Arbeitsmappe '" & xlDateiName & "' konnte nicht geöffnet werden !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    If xlApplLiefNicht Then xlAppl.Application.Quit()
                    xlAppl = Nothing
                    GoTo err_Handler
                End If

            Else
                'Verweis setzen:
                xlWB = xlAppl.Workbooks(xlDateiName)
            End If
            Dim xlDateiName00 As String
            xlDateiName00 = ""

            xlWB = xlAppl.Workbooks(xlDateiName)

            If xlWB.Workbooks.count < 2 Then
                For i = 1 To xlWB.Worksheets.count
                    xlWS = xlWB.Worksheets(i)
                Next
                xlDateiName00 = xlWS.Cells(39, 14).Value
                i = 1
                ReDim Preserve Prg(i)
                Prg(i) = xlDateiName00
                tttemp = xlDateiName00
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
    ' Hier wird geprüft ob es sich um eine Microface Datei handelt oder ob es sich 
    ' um eine Excelarbeitmappe mit mehr als einem 
    ' Tabellenblatt handelt ->
    ' 1 Blatt = LVCVTimer = 1
    ' > 1 Blatt = LVCVTimer = 2

    Sub WhatsIt()
        '*Dim boolWBOffen As Boolean
        Dim i As Short
        '*        Dim wb As Object ' As Excel.Workbook

        'If cntBack = 1 Then Command2_Click() 'cmdBeenden_Click()

        ' GetObject-Funktionsaufruf ohne erstes Argument gibt einen
        ' Verweis auf eine Instanz der Anwendung zurück. Wenn die
        ' Anwendung nicht ausgeführt wird, tritt ein Fehler auf.

        'Prüfen, ob Excel ausgeführt wird:
        On Error Resume Next
        If DBOrExcel = 2 Then
            xlAppl = GetObject(, "Excel.Application")
            If Err.Number <> 0 Then xlApplLiefNicht = True

            Err.Clear() ' Err-Objekt im Fehlerfall löschen

            'Wenn Excel nicht ausgeführt wird, Excel starten:
            If xlAppl Is Nothing Then
                xlAppl = CreateObject("Excel.Application")

                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Konnte keine Verbindung zu Excel herstellen !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    GoTo err_Handler
                End If
            End If

            'Prüfen, ob Arbeitsmappe bereits offen ist:
            '*boolWBOffen = False

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

            'LVPath As String, CVPath As String, PathNam As String
            '1 = CV ; 2 = LV
            'LVCVTimer = 1
            'LVCVTimer: 1 = CV ; 2 = LV
            If frmLogin.Option1.Visible = True Then
                If LVCVTimer = 1 Then
                    PathNam = CVPath
                ElseIf LVCVTimer = 2 Then
                    PathNam = LVPath
                End If
            Else
                'If frmLogin.File1.Path = "" Then
                PathNam = CVPath
                'Else
                'PathNam = frmLogin.File1.Path
                'End If
            End If

            'On Error Resume Next
            If Not boolWBOffen Then
                'Wenn Arbeitsmappe nicht offen, öffnen und Verweis setzen:
                Err.Clear()
                xlWB = xlAppl.Workbooks.Open(FileName:=PathNam & xlDateiName)
                'xlAppl.Application.Visible = True
                'Wenn ein Fehler aufgetreten ist...
                If Err.Number <> 0 Then
                    MsgBox("Die Arbeitsmappe '" & xlDateiName & "' konnte nicht geöffnet werden !", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, Form2.Text)
                    If xlApplLiefNicht Then xlAppl.Application.Quit()
                    xlAppl = Nothing
                    GoTo err_Handler
                Else
                    boolWBOffen = True
                End If

            Else
                'Verweis setzen:
                xlWB = xlAppl.Workbooks(xlDateiName)
                boolWBOffen = True
            End If

            'xlWB = xlAppl.Workbooks(xlDateiName)

            'xlAppl.Application.Visible = True

            If xlWB.WorkSheets.Count < 2 Then
                ' LVCVTimer: 1 = CV ; 2 = LV
                LVCVTimer = 2
            ElseIf xlWB.WorkSheets.Count > 2 Then
                LVCVTimer = 1
            End If

            'xlAppl.Application.Quit()
        End If
        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        'ausser Beenden-Button deaktivieren:
        'Dim ctl As System.Windows.Forms.Control

        '        For Each ctl In Form2.Controls
        'If ctl.Name <> "cmdBeenden" Then
        'ctl.Enabled = False
        'End If
        'Next ctl
        Exit Sub

    End Sub
End Module
