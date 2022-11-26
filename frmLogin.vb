' FMLogin Form:
' Hier wird der Anwender um Informationen gebeten, er Muss dieses Formular
' komplett ausfüllen hierraus ergibt sich welchen Weg das weitere Programm
' geht.

Option Strict Off
Option Explicit On
Friend Class frmLogin
	Inherits System.Windows.Forms.Form
	Public LoginSucceeded As Boolean
    ' Die Checkboxen wurden im Zuge des ersten Updates heraus genommen,
    ' es sollte dem Anwender keine Auswhal möglichkeit gegeben werden
    ' zwischen CV und LV zu wählen.
    ' Die anderen Checkboxen wurden für die Auswahl Belag / Scheibe
    ' Belag innen / Belag aussen angezeigt, diese sind nach dem ersten
    ' Update in den zweiten Dialog gerutscht.
    ' Die folgenden Checkboxen sind hier noch enthalten weil Sie zu beginn
    ' schon gelesen werden und einen definierten Wert darstellen.

    ' Checkbox Variablen:
    ' Padcnt = 0     -> Auswahl Belag         (1=Ja/0=Nein)
    ' Disccnt = 0    -> Auswahl Scheibe       (1=Ja/0=Nein)
    ' PadInside = 0  -> Auswahl Belag innen   (1=Ja/0=Nein)
    ' Padoutside = 0 -> Auswhal Belage aussen (1=Ja/0=Nein)
    ' LVCVTimer : 1 = CV ; 2 = LV

    Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
        Padcnt = 1
        Disccnt = 0
        If Me.Check1.Checked = True Then
            Me.Check2.Enabled = True
            Me.Check3.Enabled = True
            Me.Check4.Checked = False
        Else
            Me.Check2.Checked = False
            Me.Check3.Checked = False
            Me.Check2.Enabled = False
            Me.Check3.Enabled = False
        End If
    End Sub

    Private Sub Check2_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check2.CheckStateChanged
        PadInside = 1
        Padoutside = 0
        Me.Check3.CheckState = False
        Me.Check4.CheckState = False
    End Sub
	
    Private Sub Check3_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check3.CheckStateChanged
        Padoutside = 1
        PadInside = 0
        Me.Check2.CheckState = False
        Me.Check4.CheckState = False
    End Sub
	
    Private Sub Check4_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check4.CheckStateChanged
        Disccnt = 1
        Padoutside = 0
        PadInside = 0
        Padcnt = 0
        Me.Check1.CheckState = False
        Me.Check2.CheckState = False
        Me.Check3.CheckState = False
        Me.Check2.Enabled = False
        Me.Check3.Enabled = False

    End Sub

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        ' Schliesst das aktuelle Formular
        canchelPad = 0
        Me.Hide()
        MainMenue.Show()

    End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        ' Geht zum nächsten Dialog
        Dim i, intfaildfile As Integer, foundFile0, foundFile1 As String

        LastProj1 = "" : Sheets01 = "" : Sheets02 = "" : Sheets03 = "" : MFCnt = 0 : MessCnt = 0
        intfaildfile = 0 : foundFile0 = "" : foundFile1 = ""
        MFCnt = 0 : tttemp = "" : PrgCnt = 0

        If Me.txtUserName.Text = "" Then
            If State = "DE" Then
                MsgBox("Benutzername fehlt !", MsgBoxStyle.Information, "Benutzername fehlt !")
            Else
                MsgBox("Please, not insert a Username !", MsgBoxStyle.Information, "Username failed !")
            End If
            Exit Sub
        End If
        ' Anwendername, für Messung, hier nur die Initalien
        Usernam = Me.txtUserName.Text

        ' Die Testnummer wird aus der ComboBox gelesen, wenn ein neuer Test
        ' angelegt wurde, dann liest dieser Schritt die Testnummer aus dem
        ' entsprechenden Textfeld aus.
        If Me.CheckBox1.Checked = True Then
            testnumber = Me.TextBox1.Text
        ElseIf Me.CheckBox1.Checked = False Then
            testnumber = Me.Combo1.Text
        End If

        ' Fehlermeldung, wenn Anwender nicht alle Informationen ausgewählt bzw.
        ' Eingegeben hat.
        If Me.Option1.Checked = False And Me.Option2.Checked = False Then
            If State = "DE" Then
                MsgBox("CV/LV Auswahl nicht komplett !", MsgBoxStyle.Information, "Auswahl nicht komplett !")
            Else
                MsgBox("Please, CV/LV Selection not complete !", MsgBoxStyle.Information, "Selection not complete !")
            End If
            Exit Sub
        End If

        ' Wenn neuer Test angelegt wird, dann durchläuft das Programm dieses
        ' Prozedur:
        If LVNewCnt = 1 And Me.CheckBox1.Checked = True And Me.TextBox1.Text <> "" Then
            ' Testnummer = Dateinamen
            foundFile0 = Me.Combo1.Text
            foundFile1 = UCase(Me.TextBox1.Text & ".xls")
            xlDateiName = foundFile1
            ' Ist die Datei mit dem Namen schon vorhanden ?
            If My.Computer.FileSystem.FileExists(PathNam & foundFile1) = True Then
                intfaildfile = 1
            End If
            ' Wenn die Datei noch nicht vorhanden ist, wird dieser hier kopiert
            If intfaildfile = 0 Then
                My.Computer.FileSystem.CopyFile(LVTempPath & foundFile0, PathNam & foundFile1)
                SearchPPGNumInExcel()
                MFCnt = 1
                PrgCnt = 1
                ' Fals die Datei schon im Verzeichnis liegt, wird hier eine Fehlermeldung
                ' zur Anzeige gebracht:
            ElseIf intfaildfile = 1 Then
                If State = "DE" Then
                    MsgBox("Testnummer " & foundFile1 & " besteht bereits, bitte Test auswählen oder einen anderen Anlegen !", MsgBoxStyle.Critical, "Testnummer gefunden !")
                Else
                    MsgBox("Testnumber " & foundFile1 & " found, please select these Test or insert a new testnumber !", MsgBoxStyle.Critical, "Testnumber found !")
                End If
                Exit Sub
            End If
            ' Wenn es kein Neuer Test ist, geht es hier weiter:
        ElseIf LVNewCnt = 0 And Me.CheckBox1.Checked = False And Me.TextBox1.Text = "" Then
            foundFile1 = Me.Combo1.Text
            xlDateiName = foundFile1
            SearchPPGNumInExcel()
            MFCnt = 1
            PrgCnt = 1
        End If

        ' Um sicher zu stellen das xlDateiName auch wirklich den richtigen Namen enthält:
        If Me.Combo1.Text <> "" Then
            If xlDateiName = "" Then
                xlDateiName = Me.Combo1.Text
            End If
        Else
            ' Entsprechende Fehlermeldung:
            If State = "DE" Then
                MsgBox("Bitte Datei auswählen !", MsgBoxStyle.Information, "Messblatt fehlt !")
            Else
                MsgBox("Select a file !", MsgBoxStyle.Information, "Template failed !")
            End If
            Exit Sub
        End If
        ' Die Fehlermeldungen über die Checkboxen siehe oben,
        ' die heraus genommen wurden.
        If Me.Check1.Checked = False And Me.Check4.Checked = False Then
            If State = "DE" Then
                MsgBox("Auswahl nicht komplett !", MsgBoxStyle.Information, "Auswahl nicht komplett !")
            Else
                MsgBox("Please, Selection not complete !", MsgBoxStyle.Information, "Selection not complete !")
            End If
            Exit Sub
        End If

        If Me.Check1.Checked = True Then
            If Me.Check2.Checked = False And Me.Check3.Checked = False Then
                If State = "DE" Then
                    MsgBox("Belagauswahl nicht komplett !", MsgBoxStyle.Information, "Auswahl nicht komplett !")
                Else
                    MsgBox("Please, Pads selection not complete !", MsgBoxStyle.Information, "Selection not complete !")
                End If
                Exit Sub
            End If
        End If

        ' Verweis Neu, hier wird LVCVTimer gesetzt auf Grund der Anzahl der Tabellenblätter
        If Me.Option1.Visible = True And Me.Option2.Visible = True Then
            If Me.Option1.Checked = True Then LVCVTimer = 1
            If Me.Option2.Checked = True Then LVCVTimer = 2
        Else
            ' Wenn man nicht weis was es ist, dann prüfe es:
            WhatsIt()
        End If

        ' Dieses stellt den "normal Fall" dar, hier CV Excel Arbeitmappe 
        ' mit >= 3 Arbeitsblätter:
        If LVCVTimer = 1 Then
            If Disccnt > 0 And Padcnt < 1 Then
                LastProj1 = GetIniString("Sheets", "CVSheets01", INIPath) '***
                Sheets01 = LastProj1
            ElseIf Padcnt > 0 And PadInside > 0 Then
                LastProj1 = GetIniString("Sheets", "CVSheets02", INIPath) '***
                Sheets02 = LastProj1
            ElseIf Padcnt > 0 And Padoutside > 0 Then
                LastProj1 = GetIniString("Sheets", "CVSheets03", INIPath) '***
                Sheets03 = LastProj1
            End If

            ' Bei LV - Microface sieht dieses so aus :
        ElseIf LVCVTimer = 2 Then

            If Prg(0) = "" Then

                If LVNewCnt < 1 And MFCnt < 1 Then
                    i = 1
                    ReDim Preserve Prg(i)
                    LastProj1 = GetIniString("MFSystem", "Prg01", INIPath) '***
                    Prg(i) = LastProj1

                    i = i + 1

                    Do While LastProj1 <> ""
                        MFCnt = 1
                        If i < 10 Then
                            ReDim Preserve Prg(i)
                            LastProj1 = GetIniString("MFSystem", "Prg0" & i, INIPath) '***
                            Prg(i) = LastProj1
                        Else
                            ReDim Preserve Prg(i)
                            LastProj1 = GetIniString("MFSystem", "Prg" & i, INIPath) '***
                            Prg(i) = LastProj1
                        End If
                        If LastProj1 <> "" Then
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    i = i - 1
                    PrgCnt = i
                End If

                If MFCnt = 1 Then

                    tttemp = UCase(tttemp)

                    If tttemp <> "" Then

                        If PrgCnt >= 1 Then

                            For i = 1 To PrgCnt
                                If tttemp = Prg(i) Then
                                    If Disccnt > 0 And Padcnt < 1 Then
                                        LastProj1 = GetIniString(tttemp, "Disc", INIPath) '***
                                        Sheets01 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "DiscMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    ElseIf Padcnt > 0 And PadInside > 0 Then
                                        LastProj1 = GetIniString(tttemp, "PadInSide", INIPath) '***
                                        Sheets02 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    ElseIf Padcnt > 0 And Padoutside > 0 Then
                                        LastProj1 = GetIniString(tttemp, "PadOutSide", INIPath) '***
                                        Sheets03 = LastProj1
                                        LastProj1 = GetIniString(tttemp, "PadMeasurePoints", INIPath) '***
                                        MessCnt = LastProj1
                                    End If
                                End If
                            Next
                        End If
                    Else
                        If Disccnt > 0 And Padcnt < 1 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
                            Sheets01 = LastProj1
                        ElseIf Padcnt > 0 And PadInside > 0 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
                            Sheets02 = LastProj1
                        ElseIf Padcnt > 0 And Padoutside > 0 Then
                            LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
                            Sheets03 = LastProj1
                        End If
                    End If
                End If
                ' hier LV Excel Arbeitmappe mit >= 3 Arbeitsblätter:
            ElseIf MFCnt = 0 Then
                If Disccnt > 0 And Padcnt < 1 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets01", INIPath) '***
                    Sheets01 = LastProj1
                ElseIf Padcnt > 0 And PadInside > 0 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets02", INIPath) '***
                    Sheets02 = LastProj1
                ElseIf Padcnt > 0 And Padoutside > 0 Then
                    LastProj1 = GetIniString("Sheets", "LVSheets03", INIPath) '***
                    Sheets03 = LastProj1
                End If
            End If

        End If

        Form2.Show()
        Me.Hide()

    End Sub

	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		If Me.Combo1.Text <> "" Then
            Me.Check1.Enabled = True
			Me.Check2.Enabled = False
			Me.Check3.Enabled = False
            Me.Check4.Enabled = True
            If Me.CheckBox1.Checked = True Then
                Me.TextBox1.Enabled = True
            End If
        End If
    End Sub

    Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        cntBack = 0
        xlDateiName = ""
        Prg(0) = ""

        Me.ToolStripStatusLabel1.Text = Creatxt
        Me.Combo1.Text = ""
        IniTal()
        PathNam = CVPath
        Me.Check1.Enabled = False
        Me.Check2.Enabled = False
        Me.Check3.Enabled = False
        Me.Check4.Enabled = False
        Me.Label2.Enabled = False
        Me.Combo1.Enabled = False
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Visible = True
        Me.TextBox1.Enabled = False
        Me.TextBox1.Visible = True
        Me.Check1.Checked = True
        Me.Check2.Checked = True

        Disccnt = 0 : LVNewCnt = 0 : Padoutside = 0 : PadInside = 0 : Padcnt = 0

        If State = "DE" Then
            Me.Text = "Anmelden"
            Me.Label1.Text = "&Benutzer Initialen"
            Me.cmdOK.Text = "&OK"
            Me.cmdCancel.Text = "&Beenden"
            Me.Check1.Text = "Belag"
            Me.Check2.Text = "Innen"
            Me.Check3.Text = "Aussen"
            Me.Check4.Text = "Scheibe"
            Me.CheckBox1.Text = "Neu anlegen"
            Me.Label2.Text = "Test Nummer :"
            Me.Label3.Text = "Test Auswahl :"
        Else
            Me.Text = "Login"
            Me.Label1.Text = "Operator's Initials"
            Me.cmdOK.Text = "&OK"
            Me.cmdCancel.Text = "&Exit"
            Me.Check1.Text = "Pads"
            Me.Check2.Text = "Inside"
            Me.Check3.Text = "Outside"
            Me.Check4.Text = "Disc"
            Me.CheckBox1.Text = "Create new"
            Me.Label2.Text = "Test Number :"
            Me.Label3.Text = "Select template :"
        End If

        If ViewCVLV = "True" Then
            Me.Option1.Visible = True
            Me.Option2.Visible = True
            Me.CheckBox1.Visible = True
            Me.Label2.Visible = True
            Me.TextBox1.Visible = True
        Else
            Me.Option1.Visible = False
            Me.Option2.Visible = False
            'Me.CheckBox1.Visible = False
            'Me.Label2.Visible = False
            'Me.TextBox1.Visible = False
            Me.Combo1.Enabled = True
        End If

        If kopieview = "True" Then
            Me.CheckBox1.Visible = True
            Me.Label2.Visible = True
            Me.TextBox1.Visible = True
        Else
            Me.CheckBox1.Visible = False
            Me.Label2.Visible = False
            Me.TextBox1.Visible = False
        End If

        'Me.TopMost = True


        Exit Sub

err_Handler:
        'Wenn ein Fehler aufgetreten ist, alle Controls
        '  ausser Beenden-Button deaktivieren:
        Dim ctl As System.Windows.Forms.Control

        For Each ctl In Me.Controls
            If ctl.Name <> "cmdCancel" Then
                ctl.Enabled = False
            End If
        Next ctl
        Exit Sub

    End Sub
	
    Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
        If eventSender.Checked Then

            Me.Option2.Checked = False

            Me.File1.Path = CVPath
            Me.File1.FileName = CVFile
            Me.File1.Refresh()

            ComboFull()
        End If
    End Sub
	
    Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
        ' Selection LV
        If eventSender.Checked Then

            'IniTal()

            'Me.Option1.Checked = False
            'Me.File1.Path = LVPath
            'Me.File1.FileName = LVFile
            'Me.File1.Refresh()
            'Me.CheckBox1.Enabled = True

            ComboFull()

            If Me.Combo1.Items.Count > 0 Then
                Me.Combo1.Enabled = True
            ElseIf Me.Combo1.Items.Count = 0 Then
                Me.CheckBox1.Checked = True
                Me.Label2.Enabled = True
                Me.TextBox1.Enabled = True
                Me.Combo1.Enabled = True
                '            Else
                'If State = "DE" Then
                'MsgBox("Kann keine Datei finden, erstellen Sie einen neuen Test oder kontaktiern Sie eine Person !", MsgBoxStyle.Information, "Pfad fehlt !")
                'Else
                'MsgBox("Can`t found files in search directory, please contact personal !", MsgBoxStyle.Information, "Path failed !")
                'End If
            End If

        ' 1 = CV ; 2 = LV
            'LVCVTimer = 2

        End If
    End Sub
	
	Sub ComboFull()
        Dim i As Integer
		Dim SourcefileName As String
        Me.Combo1.Items.Clear()
		
        For i = 0 To Me.File1.Items.Count - 1
            SourcefileName = Me.File1.Items(i)
            Me.Combo1.Items.Add(SourcefileName)
        Next i
		
	End Sub
	
    Private Sub txtUserName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUserName.TextChanged

        If Len(Me.txtUserName.Text) = 2 Then
            Me.Option1.Enabled = True
            Me.Option2.Enabled = True
            Me.txtUserName.Text = UCase(Me.txtUserName.Text)
            strUsername = Me.txtUserName.Text
            xlDateiName = ""

        ElseIf Len(Me.txtUserName.Text) > 3 Then
            If State = "DE" Then
                MsgBox("Bitte, initialien eingeben (2 Buchstaben).", MsgBoxStyle.Information, "Eingabe fehlt !")
                Exit Sub
            Else
                MsgBox("Please, insert you initials (2 Characters).", MsgBoxStyle.Information, "Failed insert !")
                Exit Sub
            End If
        End If

        If kopieview = "True" Then
            Me.CheckBox1.Visible = True
            Me.CheckBox1.Enabled = True
        End If

        Me.Combo1.Enabled = True
        Me.Option1.Checked = True
        Me.File1.Refresh()

    End Sub

    Private Sub frmLogin_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.Checked = False
        Me.Option1.Enabled = False
        Me.Option2.Enabled = False
        Me.Check1.Enabled = False
        Me.Check2.Enabled = False
        Me.Check3.Enabled = False
        Me.Check4.Enabled = False
        Me.TextBox1.Enabled = False
        Me.TextBox1.Text = ""

        DiscID = ""

        Me.Option1.Checked = False
        Me.Option2.Checked = False
        Me.Check1.Checked = True
        Me.Check2.Checked = True
        Me.Check3.Checked = False
        Me.Check4.Checked = False

        'Disccnt = 0 : LVNewCnt = 0 : Padoutside = 0 : PadInside = 0 : Padcnt = 0
        'LVCVTimer = 0
        canchelPad = 0
        Me.Combo1.Text = ""
        Me.Combo1.Items.Clear()
        Me.File1.Items.Clear()
        Me.Combo1.Enabled = False
        Me.txtUserName.Text = ""
        Me.cmdOK.Text = "&OK"

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged

        If Me.CheckBox1.Checked = True Then
            LastProj1 = GetIniString("Path", "LVTempPath", INIPath) '***
            LVTempPath = LastProj1
            'LastProj1 = GetIniString("Path", "LVPath", INIPath) '***
            LastProj1 = GetIniString("Path", "FilePath", INIPath) '***
            LVPath = LastProj1

            Me.TextBox1.Enabled = True

            Me.Label2.Enabled = True
            Me.File1.Path = LVTempPath
            Me.File1.Refresh()
            Me.Combo1.Text = ""

            If State = "DE" Then
                Me.Label3.Text = "Vorlage auswahl :"
                Me.cmdOK.Text = "Erstell. u. öffnen"
            Else
                Me.Label3.Text = "Select template :"
                Me.cmdOK.Text = "Create and open"
            End If

            ComboFull()

            If Me.Combo1.Items.Count > 0 Then
                Me.Combo1.Enabled = True
            Else
                If State = "DE" Then
                    MsgBox("Kann keinen Pfad, bitte kontaktiern Sie eine Person !", MsgBoxStyle.Information, "Pfad fehlt !")
                Else
                    MsgBox("Can`t found files in search directory, please contact personal !", MsgBoxStyle.Information, "Path failed !")
                End If
            End If

            LVNewCnt = 1

        ElseIf Me.CheckBox1.Checked = False Then
            Me.TextBox1.Enabled = False
            Me.Label2.Enabled = False
            Me.File1.Path = PathNam
            Me.File1.Refresh()

            ComboFull()

            Me.cmdOK.Text = "&OK"

            ComboFull()

            If Me.Combo1.Items.Count > 0 Then
                Me.Combo1.Enabled = True
            Else
                If State = "DE" Then
                    MsgBox("Kann keinen Pfad oder  Datei finden, erstellen Sie einen neuen Test oder kontaktiern Sie eine Person !", MsgBoxStyle.Information, "Pfad fehlt !")
                Else
                    MsgBox("Can`t found files in search directory. Please, create a new test or contact personal !", MsgBoxStyle.Information, "Path failed !")
                End If
            End If

            LVNewCnt = 0

        End If
    End Sub
End Class
