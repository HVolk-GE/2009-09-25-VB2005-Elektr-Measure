Public Class ApplConfig
    Public ashort As Short
    Public Auswahlfehlt As String

    Private Sub ApplConfig_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim I As Short
        Dim MessagesPort, MessagesInstID, MessagesPath, MessagesFiletyp As String
        Dim MessagesStandardKonf, MessagesMFSystem As String
        Dim ININame As String
        Dim PortNames() As String = IO.Ports.SerialPort.GetPortNames()

        For I = PortNames.Length - 1 To 0 Step -1
            lstCommPort.Items.Add(PortNames(I))
        Next

        Me.ToolStripStatusLabel1.Text = Creatxt

        INIPath = My.Application.Info.DirectoryPath ' "C:\"
        ININame = "\Resources\config.ini"
        INIPath = INIPath & ININame

        MessagesPort = ""
        MessagesInstID = ""
        MessagesPath = ""
        MessagesFiletyp = ""
        MessagesStandardKonf = ""
        MessagesMFSystem = ""

        IniTal()

        If AdminUser <> Adminpwd Then
            Me.CheckBox8.Enabled = False
        End If

        Me.TextBox1.Text = CVPath
        Me.TextBox2.Text = LVTempPath

        Me.CheckBox1.Checked = False
        Me.CheckBox2.Checked = False
        Me.CheckBox7.Checked = False
        Me.Button1.Visible = False

        Me.Button7.Enabled = False

        Me.lstCommPort.Enabled = False

        If State = "DE" Then
            Me.Text = "Konfiguration Anwendung"
            Me.TabPage1.Text = "Meßmittel Konfiguration"
            Me.TabPage2.Text = "Pfad angaben"
            Me.TabPage3.Text = "Standard konfiguration"
            Me.TabPage4.Text = "Microface konfiguration"
            Me.TabPage6.Text = "DTV Einstellungen"
            Me.TabPage5.Text = "Nicht erwünschte Zeichen"
            Me.Label39.Text = "Hier die Zeichen Eintragen, die im Stream auftauchen, jedoch nicht nach Excel tranferiert werden soll" & Chr(10) & _
                              "Beispiel: Kern Waage sendet = ""+    22.(3)    g"" " & Chr(10) & _
                              "Es soll jedoch nicht, das ""+"" Zeichen, und die ""("" und "")"" Zeichen auch das ""g"" soll nicht gesendet werden."
            Me.Label40.Text = "Char 1"
            Me.Label41.Text = "Char 2"
            Me.Label42.Text = "Char 3"
            Me.Label43.Text = "Char 4"
            Me.Label44.Text = "Waage :"

            Me.Button7.Text = "Speichern"
            Me.Label1.Text = "Aktuelle Konfiguration :"
            Me.Label2.Text = "Bearbeitbare (aktuelle Tests) Dateien Pfad :" & Chr(10) & _
                             "(Abschliessend '\' eingeben)"
            Me.Label3.Text = "Template Dateien Pfad :" & Chr(10) & _
                             "(Abschliessend '\' eingeben)"
            Me.Label49.Text = "Pfad zur Dtv_measure.EXE" & Chr(10) & _
                             "(Abschliessend '\' eingeben)"
            Me.GroupBox4.Text = "Gefundene Port(s) :"
            Me.ComboBox1.Items.Add("Anschluss Mitutoyo USB-Box :(" & PortMitutoyo & ")") ' "PORTMitutoyo"
            Me.ComboBox1.Items.Add("Anschluss Waage :(" & PortWeigh & ")") ' "PORTWeigh"
            MessagesPort = "Serielle Schnittstelle konfiguration." & Chr(10) & _
                               "In dem Kombinationsfeld (links) werden die" & Chr(10) & _
                               "aktuellen Anschlüsse und die dazu" & Chr(10) & _
                               "gehörenden Geräte angezeigt (soweit konfiguriert)." & Chr(10) & _
                               Chr(10) & _
                               "In dem Listenfeld (links), werden die am" & Chr(10) & _
                               "Computer verfügbaren Anschlüsse angezeigt." & Chr(10) & _
                               "Wählen Sie 'Edit' und dann das Geräte und den" & Chr(10) & _
                               "dazu gehörigen Anschluss aus, damit diese" & Chr(10) & _
                               "gespeichert werden können, klicken Sie auf 'Speichern'."
            MessagesInstID = "Textfelder (links), werden die angeschlossenen Instrumente" & Chr(10) & _
                             "mittels ID angezeigt, wählen Sie 'Edit' und geben Sie ggf." & Chr(10) & _
                             "die aktuelle Messinstrument ID ein, mit einem klick" & Chr(10) & _
                             "auf 'Speichern' wird Ihre Änderung dann übernommen."

            MessagesPath = "Textfeld (links 1 von oben), hier wird der Pfad zu den aktuell" & Chr(10) & _
                           "zu bearbeiteten Versuchen angezeigt/editiert" & Chr(10) & _
                            Chr(10) & _
                            "Textfeld (links 2 von oben), hier wird der Pfad zu den Templates" & Chr(10) & _
                            "angezeigt/editiert"
            MessagesFiletyp = "Auswahlfeld (link 3 und 4 von oben), hier kann ausgewählt werden," & Chr(10) & _
                              "welches Dateienformat, der Anwender in der Auswahl angezeigt" & Chr(10) & _
                              "bekommt"
            MessagesStandardKonf = "Auf dieser Seite werden verschiedenste konfigurationen editiert," & Chr(10) & _
                                   "mehr Informationen hierfür finden Sie in der Online-Hilfe."
            MessagesMFSystem = "Hier werden Microface templates für" & Chr(10) & _
                               "Test Programm erstellt. Wenn es eine" & Chr(10) & _
                               "Vorlage gibt, kann man diese zuerst" & Chr(10) & _
                               "links Auswählen, dann auf 'Erstelle" & Chr(10) & _
                               "neuen Eintrag' klicken." & Chr(10) & _
                               "Somit werden die Vorgabewerte aus den" & Chr(10) & _
                               "linken Feldern in die neu anzulegenden" & Chr(10) & _
                               "Daten auf der rechten Seite kopiert und" & Chr(10) & _
                               "man erhält so eine 1:1 kopie der Vorgabe" & Chr(10) & _
                               "aus der linken Seite. Jedoch darf man" & Chr(10) & _
                               "nicht vergessen den 'Template Namen" & Chr(10) & _
                               "zu ändern sonst bekommt man später" & Chr(10) & _
                               "Probleme !" & Chr(10) & _
                               "Die Templates - Excel Dateien müssen ein" & Chr(10) & _
                               "bestimmtes Format vorweisen," & Chr(10) & _
                               "mehr Informationen hierzu finden Sie" & Chr(10) & _
                               "in der Online-Hilfe."
            Auswahlfehlt = "Auswahl fehlt !"
            Me.Label5.Text = "Datei Typ CV"
            Me.Label6.Text = "Datei Typ LV"
            Me.CheckBox1.Text = "Erstelle neuen Eintrag"
            Me.CheckBox2.Text = "Edit"
            Me.CheckBox3.Text = "Edit"
            Me.CheckBox4.Text = "Excel sichtbar"
            Me.CheckBox5.Text = "Edit"
            Me.CheckBox6.Text = "Neu erstellen sichbar"
            Me.CheckBox7.Text = "Edit"
            Me.CheckBox8.Text = "Edit"
            Me.Label7.Text = "Anfangszeichen, CV Testnummer:"
            Me.Label8.Text = "Anfangszeichen, LV Testnummer:"
            Me.Label48.Text = "Belag Messrichtung"
            Me.Label9.Text = "Insturment ID Anschluss Mitutoyo 1"
            Me.Label10.Text = "Insturment ID Anschluss Mitutoyo 2"
            Me.Label47.Text = "Insturment ID Anschluss Mitutoyo 3"
            Me.Label46.Text = "Insturment ID Anschluss Mitutoyo 4"
            Me.Label11.Text = "Insturment ID Anschluss Waage 1"
            Me.Label45.Text = "Insturment ID Anschluss Waage 2"
            Me.Label50.Text = "Exe Datei für DTV - Measurement (FM-Standard Prg.)"
            Me.Label12.Text = "Ländereinstellung (DE für deutsch)"
            Me.Label13.Text = "Tabellenblattnamen" & Chr(10) & _
                              "Für Exceltabellen mit 3 (oder mehr) Tabellenblättern."
            Me.Label14.Text = "CV Tabellennamen:"
            Me.Label15.Text = "Tabelle 1"
            Me.Label16.Text = "Tabelle 2"
            Me.Label17.Text = "Tabelle 3"

            Me.Label18.Text = "LV Tabellennamen:"
            Me.Label19.Text = "Tabelle 1"
            Me.Label20.Text = "Tabelle 2"
            Me.Label21.Text = "Tabelle 3"

            Me.Label22.Text = "Template Name:"
            Me.Label23.Text = "Beschriftung Pad inside"
            Me.Label24.Text = "Beschriftung Pad outside"
            Me.Label25.Text = "Beschriftung Disc"
            Me.Label26.Text = "Anzahl Messzeitpunkt Pad"
            Me.Label27.Text = "Anzahl Messzeitpunkt Disc"

            Me.Label28.Text = "Neuer Template Name:"
            Me.Label29.Text = "Beschriftung Pad inside"
            Me.Label30.Text = "Beschriftung Pad outside"
            Me.Label31.Text = "Beschriftung Disc"
            Me.Label32.Text = "Anzahl Messzeitpunkt Pad"
            Me.Label33.Text = "Anzahl Messzeitpunkt Disc"

            Me.Button2.Text = "Schliessen"
            Me.Button3.Text = "Speichern"
            Me.Button4.Text = "Speichern"
            Me.Button5.Text = "Speichern"
            Me.Button6.Text = "Speichern"
        Else
            Me.Text = "Application Config"
            Me.TabPage1.Text = "Measurement equipment config"
            Me.TabPage2.Text = "Path information"
            Me.TabPage3.Text = "Standard config"
            Me.TabPage4.Text = "Microface config"
            Me.TabPage6.Text = "DTV Settings"
            Me.TabPage5.Text = "Not used Characters"
            Me.Label39.Text = "Here insert the characters, if inside the Stream, but not send to Excel" & Chr(10) & _
                              "Example: Kern Waage send = ""+    22.(3)    g"" " & Chr(10) & _
                              "Send this without, the ""+"", ""("",  "")"" and ""g"" Characters."
            Me.Label40.Text = "Char 1"
            Me.Label41.Text = "Char 2"
            Me.Label42.Text = "Char 3"
            Me.Label43.Text = "Char 4"
            Me.Label44.Text = "Weigh instrument:"

            Me.Label1.Text = "Actual Configuration :"
            Me.Label2.Text = "Work (actual Tests) File Path :" & Chr(10) & _
                             "(end '\' insert)"
            Me.Label3.Text = "Template File Path :" & Chr(10) & _
                             "(end '\' insert)"
            Me.Label49.Text = "Path to Dtv_measure.EXE" & Chr(10) & _
                             "(end '\' insert)"
            Me.Label50.Text = "Exe file for DTV - Measurement (FM-Standard Prg.)"
            Me.GroupBox4.Text = "Found Port(s) :"
            Me.ComboBox1.Items.Add("Connect Mitutoyo USB-Box on :(" & PortMitutoyo & ")") ' "PORTMitutoyo"
            Me.ComboBox1.Items.Add("Connect Weighing Machine on :(" & PortWeigh & ")") ' "PORTWeigh"
            MessagesPort = "Serial port configuration." & Chr(10) & _
                               "In the Combinationfield (left) can you" & Chr(10) & _
                               "see the actual ports and the dazu" & Chr(10) & _
                               "Instruments, that have connect to the port." & Chr(10) & _
                               Chr(10) & _
                               "In the Listfield (left), can you see the" & Chr(10) & _
                               "ports for these Computer." & Chr(10) & _
                               "Select 'Edit' and you can select a Instrument" & Chr(10) & _
                               "and one port to connect this Instrument after that," & Chr(10) & _
                               "click 'Save' and the changes go to save."
            MessagesInstID = "Textfields (on the left side), can you see the connected Instruments" & Chr(10) & _
                             "with ID, select 'Edit' and you can edit that" & Chr(10) & _
                             "Save these, with Button 'save"""
            MessagesPath = "Textfield (left first from top), here can you edit the Path to" & Chr(10) & _
                           " actual/working tests Files." & Chr(10) & _
                            Chr(10) & _
                            "Textfeld (left second from top), here can you edit the Path to Templates." & Chr(10) & _
                            "With that can Operator make new test."
            MessagesFiletyp = "Combinationfield (left third and fourth from top), here can you select," & Chr(10) & _
                              "the file format, that the Operator can see in Combinationfield."
            MessagesStandardKonf = "In this side can you edit different configurationen," & Chr(10) & _
                                   "more Informationen can you found in the Online-Help."
            MessagesMFSystem = "Here the config for Microface templates" & Chr(10) & _
                               "Test Programs create." & Chr(10) & _
                               "If you can found a basic template in the" & Chr(10) & _
                               "combination - field on the left side." & Chr(10) & _
                               "you can make a copy of that, to right" & Chr(10) & _
                               "side and you can edit that. With the" & Chr(10) & _
                               "select 'create new'. You can but to" & Chr(10) & _
                               "a complete new without select a test" & Chr(10) & _
                               "on the left side. Select only the selectbox" & Chr(10) & _
                               "'create new'. Please, attend not the template-" & Chr(10) & _
                               "name must different have(Not 2 different config" & Chr(10) & _
                               "with 1 Templatename)." & Chr(10) & _
                               "Please, attend these Excel File for Microface" & Chr(10) & _
                               "must have a special format inside." & Chr(10) & _
                               "view for more Information the Online-" & Chr(10) & _
                               "Help !"
            Auswahlfehlt = "Selection failed !"
            Me.Label5.Text = "File Typ CV"
            Me.Label6.Text = "File Typ LV"
            Me.CheckBox1.Text = "Create new"
            Me.CheckBox2.Text = "Edit"
            Me.CheckBox3.Text = "Edit"
            Me.CheckBox4.Text = "Excel view"
            Me.CheckBox5.Text = "Edit"
            Me.CheckBox6.Text = "Create a new test view"
            Me.CheckBox7.Text = "Edit"
            Me.CheckBox8.Text = "Edit"
            Me.Label7.Text = "CV Testnumber begins with :"
            Me.Label8.Text = "CV Testnumber begins with :"
            Me.Label48.Text = "Pads Measurement directory"
            Me.Label9.Text = "Insturment ID port Mitutoyo 1"
            Me.Label10.Text = "Insturment ID port Mitutoyo 2"
            Me.Label47.Text = "Insturment ID port Mitutoyo 3"
            Me.Label46.Text = "Insturment ID port Mitutoyo 4"
            Me.Label11.Text = "Insturment ID port Weighing Machine 1"
            Me.Label45.Text = "Insturment ID port Weighing Machine 2"

            Me.Label12.Text = "Country (DE for german)"
            Me.Label13.Text = "Sheetname" & Chr(10) & _
                              "For Excelfile with > 3 Sheets."

            Me.Label14.Text = "CV Sheetname:"
            Me.Label15.Text = "Sheet 1"
            Me.Label16.Text = "Sheet 2"
            Me.Label17.Text = "Sheet 3"

            Me.Label18.Text = "LV Sheetname:"
            Me.Label19.Text = "Sheet 1"
            Me.Label20.Text = "Sheet 2"
            Me.Label21.Text = "Sheet 3"

            Me.Label22.Text = "Template Name:"
            Me.Label23.Text = "Description Pad inside"
            Me.Label24.Text = "Description Pad outside"
            Me.Label25.Text = "Description Disc"
            Me.Label26.Text = "Number of verification Pad"
            Me.Label27.Text = "Number of verification Disc"

            Me.Label28.Text = "Template Name:"
            Me.Label29.Text = "Description Pad inside"
            Me.Label30.Text = "Description Pad outside"
            Me.Label31.Text = "Description Disc"
            Me.Label32.Text = "Number of verification Pad"
            Me.Label33.Text = "Number of verification Disc"

            Me.Button2.Text = "Close"
            Me.Button3.Text = "Save"
            Me.Button4.Text = "Save"
            Me.Button5.Text = "Save"
            Me.Button6.Text = "Save"
            Me.Button7.Text = "Save"
        End If

        Me.Label28.Enabled = False
        Me.Label29.Enabled = False
        Me.Label30.Enabled = False
        Me.Label31.Enabled = False
        Me.Label32.Enabled = False
        Me.Label33.Enabled = False

        Me.TextBox20.Enabled = False
        Me.TextBox21.Enabled = False
        Me.TextBox22.Enabled = False
        Me.TextBox23.Enabled = False
        Me.TextBox24.Enabled = False
        Me.TextBox25.Enabled = False

        Me.Button3.Enabled = False
        Me.Button4.Enabled = False

        Me.TextBox1.Enabled = False
        Me.TextBox2.Enabled = False
        Me.ComboBox2.Enabled = False
        Me.ComboBox3.Enabled = False
        Me.Button5.Enabled = False

        Me.TextBox3.Enabled = False
        Me.TextBox4.Enabled = False
        Me.TextBox8.Enabled = False
        Me.TextBox9.Enabled = False
        Me.TextBox10.Enabled = False
        Me.TextBox11.Enabled = False
        Me.TextBox12.Enabled = False
        Me.TextBox13.Enabled = False
        Me.TextBox14.Enabled = False
        Me.Button6.Enabled = False

        Me.ComboBox2.Items.Add("*.xl*")
        Me.ComboBox2.Items.Add("*.xlt")
        Me.ComboBox2.Items.Add("*.xls")
        Me.ComboBox2.Text = CVFile

        Me.ComboBox3.Items.Add("*.xl*")
        Me.ComboBox3.Items.Add("*.xlt")
        Me.ComboBox3.Items.Add("*.xls")
        Me.ComboBox3.Text = LVFile

        Me.Label4.Text = MessagesPort
        Me.Label34.Text = MessagesInstID
        Me.Label35.Text = MessagesPath
        Me.Label36.Text = MessagesFiletyp
        Me.Label37.Text = MessagesStandardKonf
        Me.Label38.Text = MessagesMFSystem

        Me.TextBox3.Text = StartCVTestNr
        Me.TextBox4.Text = StartLVTestNr

        Me.TextBox33.Text = DTVPath
        Me.TextBox34.Text = DTVAnw

        If excelview = "True" Then
            Me.CheckBox4.Checked = True
        Else
            Me.CheckBox4.Checked = False
        End If

        If kopieview = "True" Then
            Me.CheckBox6.Checked = True
        Else
            Me.CheckBox6.Checked = False
        End If

        Me.TextBox5.Text = MitutoyoInstr1ID
        Me.TextBox5.Enabled = False
        Me.TextBox6.Text = MitutoyoInstr2ID
        Me.TextBox6.Enabled = False

        Me.TextBox32.Text = MitutoyoInstr3ID
        Me.TextBox32.Enabled = False
        Me.TextBox31.Text = MitutoyoInstr4ID
        Me.TextBox31.Enabled = False

        Me.TextBox7.Text = weightID
        Me.TextBox7.Enabled = False
        Me.TextBox30.Text = weight1ID
        Me.TextBox30.Enabled = False

        Me.ComboBox6.Text = PadsDir
        Me.ComboBox6.Enabled = False

        Me.TextBox8.Text = State
        Me.TextBox9.Text = CVSheets01
        Me.TextBox10.Text = CVSheets02
        Me.TextBox11.Text = CVSheets03
        Me.TextBox12.Text = LVSheets01
        Me.TextBox13.Text = LVSheets02
        Me.TextBox14.Text = LVSheets03

        Me.TextBox26.Text = Chr1
        Me.TextBox27.Text = Chr2
        Me.TextBox28.Text = Chr3
        Me.TextBox29.Text = Chr4

        Me.TextBox26.Enabled = False
        Me.TextBox27.Enabled = False
        Me.TextBox28.Enabled = False
        Me.TextBox29.Enabled = False

        ReDim Prg00(0)

        For I = 1 To 999
            If I < 10 Then
                ReDim Preserve Prg00(I)
                LastProj1 = GetIniString("MFSystem", "Prg0" & I, INIPath) '***
                Prg00(I) = LastProj1
            Else
                ReDim Preserve Prg00(I)
                LastProj1 = GetIniString("MFSystem", "Prg" & I, INIPath) '***
                Prg00(I) = LastProj1
            End If

            If LastProj1 = "" Then
                ashort = I
                Exit For
            End If
        Next

        For I = 1 To ashort - 1
            Me.ComboBox4.Items.Add(Prg00(I))
        Next

        If DBOrExcel = 0 Then
            Me.Button1.Visible = False
            Me.ComboBox5.Visible = False
        ElseIf DBOrExcel = 2 Then
            Me.ComboBox5.Text = "Excel"
            Me.Button1.Visible = False
            Me.ComboBox5.Visible = True
        ElseIf DBOrExcel = 1 Then
            Me.ComboBox5.Text = "DB"
            Me.Button1.Visible = True
            Me.ComboBox5.Visible = True
        End If

        'Me.CheckBox2.Enabled = False
        Me.ComboBox5.Enabled = False

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        MainMenue.Show()
    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedValueChanged

        Dim ttttemp As String

        ttttemp = Me.ComboBox4.Text

        LastProj1 = GetIniString(ttttemp, "PadInSide", INIPath) '***
        Me.TextBox15.Text = LastProj1

        LastProj1 = GetIniString(ttttemp, "PadOutSide", INIPath) '***
        Me.TextBox16.Text = LastProj1

        LastProj1 = GetIniString(ttttemp, "Disc", INIPath) '***
        Me.TextBox17.Text = LastProj1

        LastProj1 = GetIniString(ttttemp, "PadMeasurePoints", INIPath) '***
        Me.TextBox18.Text = LastProj1

        LastProj1 = GetIniString(ttttemp, "DiscMeasurePoints", INIPath) '***
        Me.TextBox19.Text = LastProj1

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If Me.CheckBox1.Checked = True Then
            Me.TextBox20.Enabled = True
            Me.TextBox21.Enabled = True
            Me.TextBox22.Enabled = True
            Me.TextBox23.Enabled = True
            Me.TextBox24.Enabled = True
            Me.TextBox25.Enabled = True

            If Me.CheckBox4.Text <> "" Then
                Me.TextBox20.Text = Me.ComboBox4.Text
                Me.TextBox21.Text = Me.TextBox15.Text
                Me.TextBox22.Text = Me.TextBox16.Text
                Me.TextBox23.Text = Me.TextBox17.Text
                Me.TextBox24.Text = Me.TextBox18.Text
                Me.TextBox25.Text = Me.TextBox19.Text
            End If
           
            Me.Label28.Enabled = True
            Me.Label29.Enabled = True
            Me.Label30.Enabled = True
            Me.Label31.Enabled = True
            Me.Label32.Enabled = True
            Me.Label33.Enabled = True

            Me.Button3.Enabled = True
        Else
            Me.Label28.Enabled = False
            Me.Label29.Enabled = False
            Me.Label30.Enabled = False
            Me.Label31.Enabled = False
            Me.Label32.Enabled = False
            Me.Label33.Enabled = False

            Me.TextBox20.Enabled = False
            Me.TextBox21.Enabled = False
            Me.TextBox22.Enabled = False
            Me.TextBox23.Enabled = False
            Me.TextBox24.Enabled = False
            Me.TextBox25.Enabled = False

            Me.TextBox20.Text = ""
            Me.TextBox21.Text = ""
            Me.TextBox22.Text = ""
            Me.TextBox23.Text = ""
            Me.TextBox24.Text = ""
            Me.TextBox25.Text = ""

            Me.Button3.Enabled = False
        End If

    End Sub

    Public Section, Key, sValue As String

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim i As Integer
        Dim selNam As String

        Section = "Config"

        If Me.CheckBox3.Checked = True Then
            If Me.ComboBox1.Text <> "" Then
                i = Len(ComboBox1.Text)
                selNam = Mid(ComboBox1.Text, 11, 1)
                If selNam = "M" Then
                    Key = "PORTMitutoyo"
                ElseIf selNam = "W" Then
                    Key = "PORTWeigh"
                End If

                sValue = Me.lstCommPort.SelectedItem.ToString()
                If sValue <> "" Then
                    WriteString(Section, Key, sValue)
                Else
                    MsgBox(Auswahlfehlt, MsgBoxStyle.Information)
                    Exit Sub
                End If
            End If
        End If

        If Me.CheckBox2.Checked = True Then
            If Me.TextBox5.Text <> MitutoyoInstr1ID Then
                Key = "MitutoyoInstrumentNr1"
                sValue = Me.TextBox5.Text
                WriteString(Section, Key, sValue)
            End If

            If Me.TextBox6.Text <> MitutoyoInstr2ID Then
                Key = "MitutoyoInstrumentNr2"
                sValue = Me.TextBox6.Text
                WriteString(Section, Key, sValue)
            End If

            If Me.TextBox32.Text <> MitutoyoInstr3ID Then
                Key = "MitutoyoInstrumentNr3"
                sValue = Me.TextBox32.Text
                WriteString(Section, Key, sValue)
            End If

            If Me.TextBox31.Text <> MitutoyoInstr4ID Then
                Key = "MitutoyoInstrumentNr4"
                sValue = Me.TextBox31.Text
                WriteString(Section, Key, sValue)
            End If

            If Me.TextBox7.Text <> weightID Then
                Key = "WaageInstrumentNr1"
                sValue = Me.TextBox7.Text
                WriteString(Section, Key, sValue)
            End If

            If Me.TextBox30.Text <> weight1ID Then
                Key = "WaageInstrumentNr2"
                sValue = Me.TextBox30.Text
                WriteString(Section, Key, sValue)
            End If

        End If


        If Me.CheckBox2.Checked = True Then Me.CheckBox2.Checked = False
        If Me.CheckBox3.Checked = True Then Me.CheckBox3.Checked = False

    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged

        If Me.CheckBox2.Checked = True Then
            Me.TextBox5.Enabled = True
            Me.TextBox6.Enabled = True
            Me.TextBox7.Enabled = True
            Me.Button4.Enabled = True
            Me.TextBox30.Enabled = True
            Me.TextBox31.Enabled = True
            Me.TextBox32.Enabled = True
        Else
            Me.TextBox5.Enabled = False
            Me.TextBox6.Enabled = False
            Me.TextBox7.Enabled = False
            Me.Button4.Enabled = False
            Me.TextBox30.Enabled = False
            Me.TextBox31.Enabled = False
            Me.TextBox32.Enabled = False
        End If

    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

        If Me.CheckBox3.Checked = True Then
            Me.lstCommPort.Enabled = True
            Me.Button4.Enabled = True
        Else
            Me.lstCommPort.Enabled = False
            Me.Button4.Enabled = False
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        'LastProj1 = GetIniString("Config", "excelview", INIPath)

        Section = "Config"
        Key = "excelview"

        If Me.CheckBox4.Checked = True Then
            sValue = "True"
        Else
            sValue = "False"
        End If

        WriteString(Section, Key, sValue)

    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        Section = "Config"
        Key = "Kopieview"

        If Me.CheckBox6.Checked = True Then
            sValue = "True"
        Else
            sValue = "False"
        End If

        WriteString(Section, Key, sValue)

    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged

        If Me.CheckBox5.Checked = True Then
            Me.TextBox1.Enabled = True
            Me.TextBox2.Enabled = True
            Me.ComboBox2.Enabled = True
            Me.ComboBox3.Enabled = True
            Me.Button5.Enabled = True
        Else
            Me.TextBox1.Enabled = False
            Me.TextBox2.Enabled = False
            Me.ComboBox2.Enabled = False
            Me.ComboBox3.Enabled = False
            Me.Button5.Enabled = False
        End If

    End Sub

    Private Sub Button5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.Click

        If Me.TextBox1.Text <> CVPath Then
            Section = "Path"
            Key = "FilePath"
            sValue = Me.TextBox1.Text
            WriteString(Section, Key, sValue)
            Key = "LVPath"
            sValue = Me.TextBox1.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox2.Text <> LVTempPath Then
            Section = "Path"
            Key = "LVTempPath"
            sValue = Me.TextBox2.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.ComboBox2.Text <> CVFile And Me.ComboBox2.Text <> "" Then
            Section = "Files"
            Key = "CVFile"
            sValue = Me.ComboBox2.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.ComboBox3.Text <> LVFile And Me.ComboBox3.Text <> "" Then
            Section = "Files"
            Key = "LVFile"
            sValue = ComboBox3.Text
            WriteString(Section, Key, sValue)
        End If

    End Sub

    Private Sub CheckBox7_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox7.CheckedChanged

        If Me.CheckBox7.Checked = True Then
            Me.TextBox3.Enabled = True
            Me.TextBox4.Enabled = True
            Me.TextBox8.Enabled = True
            Me.TextBox9.Enabled = True
            Me.TextBox10.Enabled = True
            Me.TextBox11.Enabled = True
            Me.TextBox12.Enabled = True
            Me.TextBox13.Enabled = True
            Me.TextBox14.Enabled = True
            Me.Button6.Enabled = True
            Me.ComboBox6.Enabled = True
        Else
            Me.TextBox3.Enabled = False
            Me.TextBox4.Enabled = False
            Me.TextBox8.Enabled = False
            Me.TextBox9.Enabled = False
            Me.TextBox10.Enabled = False
            Me.TextBox11.Enabled = False
            Me.TextBox12.Enabled = False
            Me.TextBox13.Enabled = False
            Me.TextBox14.Enabled = False
            Me.Button6.Enabled = False
            Me.ComboBox6.Enabled = False
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click

        ' Location :
        If Me.TextBox8.Text <> State Then
            Section = "Location"
            Key = "State"
            sValue = Me.TextBox8.Text
            WriteString(Section, Key, sValue)
        End If

        ' Testnumber begins with:
        If Me.TextBox3.Text <> StartCVTestNr Then
            Section = "Config"
            Key = "BeginCVTestNr"
            sValue = Me.TextBox3.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox4.Text <> StartLVTestNr Then
            Section = "Config"
            Key = "BeginLVTestNr"
            sValue = Me.TextBox4.Text
            WriteString(Section, Key, sValue)
        End If

        ' CV Sheets :
        If Me.TextBox9.Text <> CVSheets01 Then
            Section = "Sheets"
            Key = "CVSheets01"
            sValue = Me.TextBox9.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox10.Text <> CVSheets02 Then
            Section = "Sheets"
            Key = "CVSheets02"
            sValue = Me.TextBox10.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox11.Text <> CVSheets03 Then
            Section = "Sheets"
            Key = "CVSheets03"
            sValue = Me.TextBox11.Text
            WriteString(Section, Key, sValue)
        End If

        ' LV Sheets:
        If Me.TextBox12.Text <> LVSheets01 Then
            Section = "Sheets"
            Key = "LVSheets01"
            sValue = Me.TextBox12.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox13.Text <> LVSheets02 Then
            Section = "Sheets"
            Key = "LVSheets02"
            sValue = Me.TextBox13.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.TextBox14.Text <> LVSheets03 Then
            Section = "Sheets"
            Key = "LVSheets03"
            sValue = Me.TextBox14.Text
            WriteString(Section, Key, sValue)
        End If

        ' Database or Excel :
        If Me.ComboBox5.Visible = True Then
            If Me.ComboBox5.Text <> "" Then
                Section = "Config"
                Key = "DBOrExcel"
                sValue = Me.ComboBox5.Text
                WriteString(Section, Key, sValue)
            End If
        End If

        ' Directory by Pad's round or updown ?
        If Me.ComboBox6.Text <> PadsDir Then
            Section = "Config"
            Key = "Padsdirector"
            sValue = Me.ComboBox6.Text
            WriteString(Section, Key, sValue)
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim ttttemp As String

        Section = "MFSystem"
        Key = "LVSheets03"
        sValue = Me.TextBox20.Text
        ttttemp = sValue

        If ashort < 10 Then
            Key = "Prg0" & ashort
        Else
            Key = "Prg" & ashort
        End If

        WriteString(Section, Key, sValue)

        '###########################################################################################

        Section = ttttemp
        Key = "PadInSide"
        sValue = Me.TextBox21.Text
        WriteString(Section, Key, sValue)

        LastProj1 = GetIniString(ttttemp, "PadInSide", INIPath) '***

        Section = ttttemp
        Key = "PadOutSide"
        sValue = Me.TextBox22.Text
        WriteString(Section, Key, sValue)

        Section = ttttemp
        Key = "Disc"
        sValue = Me.TextBox23.Text
        WriteString(Section, Key, sValue)

        Section = ttttemp
        Key = "PadMeasurePoints"
        sValue = Me.TextBox24.Text
        WriteString(Section, Key, sValue)

        Section = ttttemp
        Key = "DiscMeasurePoints"
        sValue = Me.TextBox25.Text
        WriteString(Section, Key, sValue)

        Me.CheckBox1.Checked = False

    End Sub

    Private Sub ComboBox5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.TextChanged
        If Me.ComboBox5.Text = "DB" Then
            Me.Button1.Visible = True
            If State = "DE" Then
                Me.Button1.Text = "Datenbank konfig"
            Else
                Me.Button1.Text = "Database config"
            End If
        Else
            Me.Button1.Visible = False
        End If

        If Me.ComboBox5.Text <> "" Then
            Section = "Config"
            Key = "DBOrExcel"
            sValue = Me.ComboBox5.Text
            WriteString(Section, Key, sValue)
        End If

        If Me.ComboBox5.Text = "Excel" Then
            DBOrExcel = 2
        ElseIf Me.ComboBox5.Text = "DB" Then
            DBOrExcel = 1
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub CheckBox8_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox8.CheckedChanged

        If Me.CheckBox8.Checked = True Then
            Me.Button7.Enabled = True
            Me.TextBox26.Enabled = True
            Me.TextBox27.Enabled = True
            Me.TextBox28.Enabled = True
            Me.TextBox29.Enabled = True
        Else
            Me.Button7.Enabled = False
            Me.TextBox26.Enabled = False
            Me.TextBox27.Enabled = False
            Me.TextBox28.Enabled = False
            Me.TextBox29.Enabled = False
        End If

    End Sub

    Private Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click

        Section = "NotUsedChar"

        Key = "Char1"
        sValue = Me.TextBox26.Text
        WriteString(Section, Key, sValue)

        Key = "Char2"
        sValue = Me.TextBox27.Text
        WriteString(Section, Key, sValue)

        Key = "Char3"
        sValue = Me.TextBox28.Text
        WriteString(Section, Key, sValue)

        Key = "Char4"
        sValue = Me.TextBox29.Text
        WriteString(Section, Key, sValue)

    End Sub

End Class